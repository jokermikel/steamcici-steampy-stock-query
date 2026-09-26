# -*- coding: utf-8 -*-
"""
多平台 CDK 售卖库存查询（图形版 / 双击即用）
================================================

支持平台：
· STEAMCICI —— 鉴权凭证为浏览器 Cookie 里的 Admin-Token
· SteamPY   —— 鉴权凭证为浏览器 Cookie 里的 accessToken

· 双击本文件即可打开窗口（.pyw，无黑色控制台）。
· 顶部下拉框选平台，首次使用需填写一次该平台的凭证，会保存在本文件同目录的
  steamcici_config.json 中，之后打开会自动查询，也可随时点【查询】刷新。
· 凭证失效（报 401）时，点【更换凭证】，重新登录网站复制即可。

依赖：Python3 标准库（tkinter / urllib / zipfile），无需 pip install。
说明：平台对 CDK 明文做了脱敏，本工具显示的是「脱敏 CDK + 库存数量 +
      出库/兑换状态」，与网页上你本人看到的一致。
导出：默认导出 .xlsx —— 真正的两张工作表（页签）：【游戏库存】与【CDK 明细】，
      各占一页，互不混合；也可在保存对话框里改选 .csv，此时两张表上下排列
      在同一个 CSV 文件内（CSV 格式本身不支持多工作表）。
"""

import csv
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
import zipfile
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from urllib import parse as urlparse
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "steamcici_config.json")

PAGE_SIZE = 100
REQUEST_INTERVAL = 0.3
MAX_RETRY = 3
MAX_PAGES = 500   # 翻页硬上限：万一接口忽略翻页参数，避免无限循环刷请求
TIMEOUT = 20

FONT = "Microsoft YaHei UI"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")


# ============================ 通用网络层（两个平台共用） ============================ #

class QueryError(Exception):
    """查询过程中的可预期错误：凭证失效、接口报错、网络失败等。"""


def _get_json(url, params, headers, timeout=TIMEOUT):
    """最底层的一次 GET：拼参数、带重试、解析 JSON。
    与平台无关，只负责「把 JSON 拿回来」，平台语义交给上层包装函数。"""
    if params:
        url = url + ("&" if "?" in url else "?") + urlparse.urlencode(params)
    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            req = urlrequest.Request(url, headers=headers, method="GET")
            with urlrequest.urlopen(req, timeout=timeout) as resp:
                body = resp.read().decode("utf-8", "replace")
            try:
                return json.loads(body)
            except ValueError:
                # 通常是返回了登录页 / 风控提示页而不是数据
                last_err = ("返回的不是 JSON（多半是凭证失效或被风控拦截）：%s"
                            % body.strip().replace("\n", " ")[:100])
        except HTTPError as e:
            if e.code == 401:
                raise QueryError("认证失败(401)：凭证无效或已过期，请更换。")
            last_err = "HTTP %s" % e.code
        except (URLError, TimeoutError, OSError) as e:
            last_err = str(e)
        if attempt < MAX_RETRY:
            time.sleep(0.8 * attempt)
    raise QueryError("请求失败：%s（已重试 %d 次）" % (last_err, MAX_RETRY))


def _paged(fetch_page, base_params, page_param="pageNumber"):
    """通用翻页：fetch_page(params) 需返回 (rows, total)。
    翻页参数名各平台不同（steamcici 用 pageNum，steampy 用 pageNumber），
    因此通过 page_param 配置；停止规则与两边站点一致：返回空 / 已够 total /
    不足一页 即停止。"""
    rows, page, total = [], 1, None
    while True:
        params = dict(base_params)
        params.update({page_param: page, "pageSize": PAGE_SIZE})
        page_rows, total = fetch_page(params)
        page_rows = page_rows or []
        rows.extend(page_rows)
        if (not page_rows or (total is not None and len(rows) >= total)
                or len(page_rows) < PAGE_SIZE or page >= MAX_PAGES):
            break
        page += 1
        time.sleep(REQUEST_INTERVAL)
    return rows


def _to_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def _breakdown(rows, field="分区"):
    """按某列统计条数，返回「国区 320、全球区 250」这样的字符串。
    该列不存在时返回空串（例如 steamcici 的数据没有分区概念）。"""
    if not rows or field not in rows[0]:
        return ""
    order, agg = [], {}
    for r in rows:
        k = r.get(field) or "—"
        if k not in agg:
            agg[k] = 0
            order.append(k)
        agg[k] += 1
    return "、".join("%s %d" % (k, agg[k]) for k in order)


def _sort_desc(rows, field):
    """按指定列的**数值**降序排列（稳定排序，同值保持接口原顺序）。
    field 为空或该列不存在时原样返回，避免把 steamcici 那种没有库存列的表打乱。"""
    if not rows or not field or field not in rows[0]:
        return rows
    return sorted(rows, key=lambda r: _to_int(r.get(field)), reverse=True)


def _sort_by_order(rows, field, order):
    """按**取值优先级**排序：order 里靠前的排前面，不在 order 中的排到最后，
    同一档内保持原顺序。用于「未出库在前、已售出在后」这类非数值排序。"""
    if not rows or not field or field not in rows[0]:
        return rows
    rank = dict((v, i) for i, v in enumerate(order))
    last = len(order)
    return sorted(rows, key=lambda r: rank.get(r.get(field), last))


def _apply_sort(rows, spec):
    """按平台的排序配置排序。spec 支持三种写法：
       · "" 或 None          → 保持接口原顺序
       · "列名"              → 该列数值降序
       · ("列名", [值1, 值2]) → 按取值优先级，值1 在前
    """
    if not spec:
        return rows
    if isinstance(spec, (tuple, list)):
        return _sort_by_order(rows, spec[0], spec[1])
    return _sort_desc(rows, spec)


def _map(table, code):
    if code is None or code == "":
        return ""
    return table.get(str(code), "状态码:%s" % code)


def _min_price(*vals):
    """取多个价格候选值中的最小值（返回原值，保留原始小数位）。
    全部为空时返回空串。对应 SteamPY 网页端该列的算法：
        "¥" + min(steamGame.keyAveAmt, steamGame.keyTxAmt)
    注意：网页端把这列命名为「最新成交价」，但实际语义是**最低售价**
    （已用真实数据核对：min 结果与网页显示值完全一致）。"""
    best, best_num = "", None
    for v in vals:
        if v in (None, ""):
            continue
        try:
            n = float(v)
        except (TypeError, ValueError):
            continue
        if best_num is None or n < best_num:
            best, best_num = v, n
    return best


def clean_time(t):
    if not t or not isinstance(t, str):
        return t or ""
    s = t.replace("T", " ")
    s = re.sub(r"\.\d+", "", s)
    s = re.sub(r"\s*(Z|[+-]\d{2}:?\d{2})$", "", s)
    return s.strip()[:19]


# ============================ 平台一：STEAMCICI ============================ #

SC_BASE = "https://steamcici.com/prod-api"

# ---- 状态映射（均已对照站点确认）----
SELL_STATUS = {"1": "出售中", "2": "已售空", "0": "已下架"}
STOCK_STATUS = {"0": "未出库(在库)", "1": "已售出"}
ORDER_STATUS = {"2": "支付完成"}
EXCHANGE_STATUS = {"1": "兑换成功"}
# CDK 明细的默认排序权重：还没出库的排前面，已经卖掉的垫底
SC_STOCK_ORDER = ("未出库(在库)", "已售出")


def sc_fetch(path, params, token):
    """STEAMCICI 接口调用：Bearer Token + code 校验。"""
    headers = {
        "Authorization": "Bearer " + token,
        "Accept": "application/json, text/plain, */*",
        "User-Agent": UA,
        "Referer": "https://steamcici.com/sell",
    }
    data = _get_json(SC_BASE + path, params, headers)
    code = data.get("code")
    if code == 401:
        raise QueryError("认证失败(401)：Admin-Token 无效或已过期，请更换 Token。")
    if code != 200:
        raise QueryError("接口返回错误 code=%s msg=%s" % (code, data.get("msg")))
    return data


def sc_verify_token(token):
    data = sc_fetch("/user/getInfo", {}, token)
    d = data.get("data") or {}
    user = d.get("user") if isinstance(d, dict) else None
    if isinstance(user, dict):
        return user.get("nickName") or user.get("userName") or "已登录用户"
    return "已登录用户"


def sc_list_all_sells(token):
    """三个状态（在售/售空/下架）各拉一遍，合并成一个列表。"""
    all_rows = []
    for st in ("1", "2", "0"):
        all_rows.extend(_paged(
            lambda p, st=st: _sc_sell_page(token, st, p), {}, page_param="pageNum"))
        time.sleep(REQUEST_INTERVAL)
    return all_rows


def _sc_sell_page(token, status, params):
    data = sc_fetch("/user/system/shopSell/list", dict(params, status=status), token)
    return data.get("rows") or [], data.get("total")


def _sc_stock_page(token, sell_id, params):
    data = sc_fetch("/user/system/shopSell/queryStock",
                    dict(params, sellId=sell_id), token)
    return data.get("rows") or [], data.get("total")


def parse_order(order):
    if not order:
        return "", "", "", "", ""
    if isinstance(order, dict):
        return (order.get("orderNo") or "",
                _map(ORDER_STATUS, order.get("orderStatus")),
                _map(EXCHANGE_STATUS, order.get("exchangeStatus")),
                clean_time(order.get("payTime") or order.get("buyTime")),
                clean_time(order.get("exchangeTime")))
    return str(order), "", "", "", ""


def sc_collect(token, progress=None):
    """STEAMCICI：返回 (nick, summary_rows, detail_rows)。"""
    nick = sc_verify_token(token)
    if progress:
        progress("正在拉取挂售商品 ...")
    sells = sc_list_all_sells(token)
    summary = [{
        "游戏名称": s.get("shopName", ""),
        "游戏ID": s.get("shopId", ""),
        "挂售记录ID": s.get("id", ""),
        "状态": _map(SELL_STATUS, s.get("status")),
        "当前库存": _to_int(s.get("stockNum")),
        "累计库存": _to_int(s.get("sumStockNum")),
        "出售单价": s.get("sellPrice", ""),
        # 注意：接口字段名是 lastDealPrice，但平台实际给的是「最低售价」，
        # 显示名以此为准（不要被字段名误导成「最新成交价」）
        "最低售价": s.get("lastDealPrice", ""),
        "实际到手": s.get("realAmount", ""),
        "Steam原价": s.get("steamPrice", ""),
        "挂售时间": s.get("createTime", ""),
    } for s in sells]

    details = []
    for i, s in enumerate(sells, 1):
        sid = s.get("id")
        name = s.get("shopName", "")
        if progress:
            progress("正在拉取《%s》的 CDK 明细 (%d/%d) ..." % (name, i, len(sells)))
        for c in _paged(lambda p, sid=sid: _sc_stock_page(token, sid, p), {},
                        page_param="pageNum"):
            on, ps, es, pt, et = parse_order(c.get("transactionOrder"))
            details.append({
                "游戏名称": name,
                "挂售记录ID": sid,
                "脱敏CDK": c.get("cdkStr", ""),
                "库存状态": _map(STOCK_STATUS, c.get("stockStatus")),
                "订单号": on,
                "支付状态": ps,
                "兑换状态": es,
                "售出时间": pt,
                "兑换时间": et,
                "入库时间": c.get("createTime", ""),
            })
        time.sleep(REQUEST_INTERVAL)
    return nick, summary, details


# ============================ 平台二：SteamPY ============================ #

PY_BASE = "https://steampy.com/xboot"

# SteamPY 按「区」独立挂售，同一套接口名挂在三个区前缀下：
#   国区   /steamKeySale/...  （对应网页端 pyUserInfo/sellerCDKey 的「国区」页签）
#   全球区 /usKeySale/...
#   俄罗斯区 /ruKeySale/...
#
# 【重要】每个区只取「我自己」的两个接口（已对照前端卖家页面确认）：
#   · listGame —— 我挂售的游戏及其库存（网页端「库存总列表」）
#   · listSelf —— 我的挂单记录（网页端点「查看挂单」后的列表，commonUrl）
# 千万不要用 listSale —— 那是全市场挂单列表，会把别人的挂单和价格也导进来。
PY_REGIONS = (
    ("国区", "steamKeySale"),
    ("全球区", "usKeySale"),
    ("俄罗斯区", "ruKeySale"),
)

# 挂售状态码映射（取自前端 formatStatus 过滤器的完整版）
PY_SALE_STATUS = {"1": "出售", "0": "关闭", "R": "待审批",
                  "U": "等待上货", "D": "上货完成", "S": "全部结单"}


def py_fetch(path, params, token):
    """SteamPY 接口调用：请求头 accessToken 鉴权 + success/code 校验。"""
    headers = {
        "accessToken": token,
        "APP_TOKEN": "",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": UA,
        "Referer": "https://steampy.com/",
    }
    data = _get_json(PY_BASE + path, params, headers)
    if not isinstance(data, dict):
        raise QueryError("接口返回结构异常（期望 JSON 对象）：%r" % (data,))
    code = data.get("code")
    if code == 401:
        raise QueryError("认证失败(401)：accessToken 无效或已过期，请更换。")
    if data.get("success") is True or code == 200:
        return data
    raise QueryError("接口返回错误 code=%s message=%s"
                     % (code, data.get("message") or data.get("msg")))


def _py_page(path, base_params, token):
    """SteamPY 分页包装：数据在 result.content，总数在 result.totalElements。"""
    def fetch(params):
        result = py_fetch(path, params, token).get("result") or {}
        if not isinstance(result, dict):
            return [], None
        return result.get("content") or [], result.get("totalElements")
    return _paged(fetch, base_params)


def _py_game_row(region, g):
    """listGame 的一条：我挂售的某个游戏及其库存。
    字段来自前端渲染确认：stock / total / steamGame{gameName, gameUrl}。"""
    sg = g.get("steamGame") if isinstance(g.get("steamGame"), dict) else {}
    return {
        "分区": region,
        "游戏名称": sg.get("gameName") or g.get("gameName") or "",
        "游戏ID": g.get("gameId", ""),
        "当前库存": _to_int(g.get("stock")),
        "库存总数": _to_int(g.get("total")),
        "Steam链接": sg.get("gameUrl") or g.get("gameUrl") or "",
    }


def _py_self_row(region, c):
    """listSelf 的一条：我的一笔挂单记录。
    字段来自前端渲染确认：createTime / stock / total / keyPrice /
    saleStatus / steamGame{gameName, gameUrl}。"""
    sg = c.get("steamGame") if isinstance(c.get("steamGame"), dict) else {}
    return {
        "分区": region,
        "游戏名称": sg.get("gameName") or c.get("gameName") or "",
        "游戏ID": c.get("gameId", ""),
        "挂售记录ID": c.get("id", ""),
        "库存": _to_int(c.get("stock")),
        "库存总数": _to_int(c.get("total")),
        # 最低售价 = min(keyAveAmt, keyTxAmt)，与网页端显示值一致
        # （网页表头写的是「最新成交价」，语义实为最低售价）
        "最低售价": _min_price(sg.get("keyAveAmt"), sg.get("keyTxAmt")),
        "挂售单价": c.get("keyPrice", ""),
        "挂售状态": _map(PY_SALE_STATUS, c.get("saleStatus")),
        "创建时间": clean_time(c.get("createTime")),
        "Steam链接": sg.get("gameUrl") or "",
    }


def py_collect(token, progress=None):
    """SteamPY：返回 (nick, summary_rows, detail_rows)，结构与 steamcici 版一致。

    两张表都**只取我自己**的数据，与网页端卖家页面完全一致：
      · 汇总 ← /{区}/listGame（我挂售的游戏及库存）
      · 明细 ← /{区}/listSelf（我的挂单记录）
    每个区各 1 次请求，不再逐游戏循环，避免把全市场挂单也拉回来。
    """
    nick = ""
    summary, details = [], []
    seen, dup = set(), 0
    for region, prefix in PY_REGIONS:
        if progress:
            progress("正在拉取 SteamPY【%s】我的游戏库存 ..." % region)
        for g in _py_page("/%s/listGame" % prefix, {}, token):
            summary.append(_py_game_row(region, g))
        time.sleep(REQUEST_INTERVAL)

        if progress:
            progress("正在拉取 SteamPY【%s】我的挂单明细 ..." % region)
        try:
            rows = _py_page("/%s/listSelf" % prefix,
                            {"sort": "keyPrice", "order": "desc"}, token)
        except QueryError:
            rows = []
            if progress:
                progress("SteamPY【%s】：挂单明细拉取失败，已跳过（汇总不受影响）。"
                         % region)
        for c in rows:
            rid = c.get("id")
            if rid not in (None, ""):
                key = (region, str(rid))
                if key in seen:       # 同一条记录被重复返回，丢弃
                    dup += 1
                    continue
                seen.add(key)
            details.append(_py_self_row(region, c))
        time.sleep(REQUEST_INTERVAL)

    notes = []
    if dup:
        notes.append("已自动剔除 %d 条重复的挂单记录。" % dup)
    PLATFORMS["steampy"].last_note = "  ".join(notes)
    return nick, summary, details


# ============================ 平台注册表 ============================ #

class Platform(object):
    """一个可查询平台的完整描述：鉴权、接口、列定义、导出定义。"""

    def __init__(self, key, label, token_name, token_hint, config_key,
                 collect, summary_title, detail_title,
                 summary_cols, detail_cols,
                 export_summary_cols, export_detail_cols,
                 sort_summary="", sort_detail=""):
        self.key = key
        self.label = label
        self.token_name = token_name
        self.token_hint = token_hint
        self.config_key = config_key
        self.collect = collect
        self.summary_title = summary_title
        self.detail_title = detail_title
        self.summary_cols = summary_cols
        self.detail_cols = detail_cols
        self.export_summary_cols = export_summary_cols
        self.export_detail_cols = export_detail_cols
        # 两张表的默认排序：按该列数值降序；留空则保持接口返回顺序
        self.sort_summary = sort_summary
        self.sort_detail = sort_detail
        # 一次查询的附带说明（例如「已去重 N 条重复记录」），由 collect 写入、
        # 界面读取后清空，用于把「数据为什么是这个数」如实告诉用户。
        self.last_note = ""


PASTE_TIP = ("\n【控制台显示 undefined 是正常的】：\n"
             "copy() 只负责写入剪贴板、本身没有返回值，回显 undefined 即代表已复制成功，\n"
             "直接去粘贴即可（可先粘到记事本确认是不是一串长字符）。\n"
             "【若弹出「Don't paste code into the DevTools Console…」安全提示】：\n"
             "这是浏览器的正常防诈骗保护，不是出错。在控制台最下方输入框里\n"
             "手动输入「允许粘贴」四个字并回车，然后再粘贴一次即可。")

SC_HINT = ("获取方法：电脑浏览器登录 steamcici.com → 按 F12 → 点「控制台/Console」\n"
           "→ 粘贴下面这行回车，Admin-Token 会自动复制到剪贴板：\n"
           "copy(document.cookie.split('; ').find(r=>r.startsWith('Admin-Token='))"
           ".split('=').slice(1).join('='))" + PASTE_TIP)

PY_HINT = ("获取方法：电脑浏览器登录 steampy.com（登录卖家账号）→ 按 F12 → 点「控制台/Console」\n"
           "→ 粘贴下面这行回车，accessToken 会自动复制到剪贴板：\n"
           "copy(localStorage.getItem('accessToken') || '没找到，请确认已登录')\n"
           "注意：它存在 localStorage 里，不是 Cookie（同理也可在 F12 → Application\n"
           "→ Storage → Local Storage 下找到 accessToken 这一项）。" + PASTE_TIP)

PLATFORMS = {}

PLATFORMS["steamcici"] = Platform(
    key="steamcici",
    label="STEAMCICI",
    token_name="Admin-Token",
    token_hint=SC_HINT,
    config_key="token",
    collect=sc_collect,
    summary_title="游戏库存",
    detail_title="CDK 明细",
    summary_cols=[
        ("游戏名称", 230), ("状态", 80), ("当前库存", 80), ("累计库存", 80),
        ("出售单价", 90), ("最低售价", 90), ("实际到手", 90), ("挂售时间", 150),
    ],
    detail_cols=[
        ("游戏名称", 200), ("脱敏CDK", 160), ("库存状态", 100), ("订单号", 170),
        ("支付状态", 90), ("兑换状态", 90), ("售出时间", 150), ("兑换时间", 150),
    ],
    export_summary_cols=[
        "游戏名称", "状态", "当前库存", "累计库存", "出售单价", "最低售价",
        "实际到手", "Steam原价", "挂售时间", "游戏ID", "挂售记录ID",
    ],
    export_detail_cols=[
        "游戏名称", "挂售记录ID", "脱敏CDK", "库存状态", "订单号", "支付状态",
        "兑换状态", "售出时间", "兑换时间", "入库时间",
    ],
    # 游戏库存按「当前库存」降序；
    # CDK 明细是逐张 CDK、没有库存数字列，改为按状态优先级：
    # 「未出库(在库)」在前，「已售出」往后排
    sort_summary="当前库存",
    sort_detail=("库存状态", SC_STOCK_ORDER),
)

PLATFORMS["steampy"] = Platform(
    key="steampy",
    label="SteamPY",
    token_name="accessToken",
    token_hint=PY_HINT,
    config_key="steampy_token",
    collect=py_collect,
    summary_title="游戏库存",
    detail_title="挂售明细",
    summary_cols=[
        ("分区", 70), ("游戏名称", 240), ("游戏ID", 90), ("当前库存", 80),
        ("库存总数", 80), ("Steam链接", 260),
    ],
    detail_cols=[
        ("分区", 70), ("游戏名称", 200), ("挂售记录ID", 110), ("库存", 70),
        ("库存总数", 80), ("最低售价", 90), ("挂售单价", 90), ("挂售状态", 90),
        ("创建时间", 150),
    ],
    export_summary_cols=[
        "分区", "游戏名称", "游戏ID", "当前库存", "库存总数", "Steam链接",
    ],
    export_detail_cols=[
        "分区", "游戏名称", "游戏ID", "挂售记录ID", "库存", "库存总数",
        "最低售价", "挂售单价", "挂售状态", "创建时间", "Steam链接",
    ],
    # 游戏库存按「当前库存」降序，挂售明细按「库存」降序
    sort_summary="当前库存",
    sort_detail="库存",
)

PLATFORM_ORDER = ["steamcici", "steampy"]


def platform_by_label(label):
    for p in (PLATFORMS[k] for k in PLATFORM_ORDER):
        if p.label == label:
            return p
    return PLATFORMS[PLATFORM_ORDER[0]]


# ============================ 配置 ============================ #

def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("提示", "配置保存失败：%s\n本次仍可使用。" % e)


def save_token(plat, token):
    """按平台分别保存凭证（各写各的键，互不覆盖），并记住最后使用的平台。
    沿用原文件名与原有的 "token" 键，老的配置文件可直接继续用。"""
    cfg = load_config()
    cfg[plat.config_key] = token
    cfg["last_platform"] = plat.key
    save_config(cfg)


def load_token(plat):
    return load_config().get(plat.config_key, "") or ""


# ============================ 凭证设置对话框 ============================ #

class TokenDialog(tk.Toplevel):
    def __init__(self, master, plat):
        super().__init__(master)
        self.plat = plat
        self.title("设置 %s 的 %s" % (plat.label, plat.token_name))
        self.resizable(False, False)
        self.result = None
        self.transient(master)
        self.grab_set()

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请粘贴浏览器 Cookie 中的 %s：" % plat.token_name,
                  font=(FONT, 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frm, text=plat.token_hint, justify="left", foreground="#555").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 8))

        self.var_show = tk.BooleanVar(value=True)
        self.entry = ttk.Entry(frm, width=52, font=(FONT, 10))
        self.entry.grid(row=2, column=0, columnspan=2, sticky="we")
        tk.Checkbutton(frm, text="显示", variable=self.var_show,
                       command=self._toggle, font=(FONT, 9)).grid(
            row=3, column=0, sticky="w", pady=(4, 0))
        self._toggle()

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="取消", command=self._cancel).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="保存并查询", command=self._ok).pack(side="right")

        self.entry.focus_set()
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.update_idletasks()
        self.geometry("+%d+%d" % (master.winfo_rootx() + 80, master.winfo_rooty() + 120))

    def _toggle(self):
        self.entry.config(show="" if self.var_show.get() else "*")

    def _ok(self):
        v = self.entry.get().strip()
        if not v:
            messagebox.showwarning("提示", "请先粘贴 %s。" % self.plat.token_name,
                                   parent=self)
            return
        if v.lower().startswith("bearer "):
            v = v[7:].strip()
        self.result = v
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ============================ 主窗口 ============================ #

# ============================ 导出 ============================ #

def write_table(writer, title, cols, rows):
    """在同一个 CSV writer 中写出一张独立的表：表名行 + 表头行 + 数据行。
    列顺序严格按 cols 输出，缺失字段补空串，保证字段与数据逐列对应。"""
    writer.writerow([title])
    writer.writerow(cols)
    for r in rows:
        writer.writerow([r.get(c, "") for c in cols])


def export_blocks(path, plat, summary, details):
    """把该平台的【汇总】与【明细】写成同一个 CSV 文件中的两张表，
    两张表之间用一个空行分隔。表为空时仍保留表名与表头，保证结构完整。"""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        write_table(w, plat.summary_title, plat.export_summary_cols, summary)
        w.writerow([])  # 空行：两张表的分隔标记
        write_table(w, plat.detail_title, plat.export_detail_cols, details)


# ============ XLSX 写出（纯标准库手写 OOXML，无需 openpyxl 等第三方库） ============

XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# 这些字段在 Excel 里写成「数字」，可直接求和；其余一律写成文本，
# 以免长 ID 被显示成科学计数法、或 "12.50" 这类价格被抹掉末尾 0。
XLSX_NUMERIC_FIELDS = {"当前库存", "累计库存"}

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xml_text(v):
    """转义为可安全放入 XML 文本节点的字符串，并剔除 XML 不允许的控制字符。"""
    s = "" if v is None else str(v)
    s = _ILLEGAL_XML.sub("", s)
    return (s.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _col_letter(idx):
    """1 -> A, 26 -> Z, 27 -> AA"""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _sheet_xml(cols, rows):
    """生成一张工作表的 sheet xml：第 1 行为表头，其后为数据行。"""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="%s"><sheetData>' % XLSX_NS]

    def cell(ref, value, numeric):
        if numeric:
            return '<c r="%s"><v>%s</v></c>' % (ref, value)
        return ('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                % (ref, _xml_text(value)))

    out.append('<row r="1">')
    for ci, name in enumerate(cols, 1):
        out.append(cell("%s1" % _col_letter(ci), name, False))
    out.append('</row>')

    for ri, r in enumerate(rows, 2):
        out.append('<row r="%d">' % ri)
        for ci, name in enumerate(cols, 1):
            v = r.get(name, "")
            numeric = (name in XLSX_NUMERIC_FIELDS
                       and isinstance(v, int) and not isinstance(v, bool))
            out.append(cell("%s%d" % (_col_letter(ci), ri), v, numeric))
        out.append('</row>')

    out.append('</sheetData></worksheet>')
    return "".join(out)


def create_xlsx(path, sheets):
    """把多张工作表写进一个 .xlsx 文件。
    sheets: [(工作表名, 字段列表, 数据行列表), ...]  —— 每张表各占一个页签。"""
    n = len(sheets)

    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="%s">' % CT_NS,
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
          'package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
          'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(1, n + 1):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType='
                  '"application/vnd.openxmlformats-officedocument.spreadsheetml.'
                  'worksheet+xml"/>' % i)
    ct.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
              'openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    ct.append('</Types>')

    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="%s" xmlns:r="%s"><sheets>' % (XLSX_NS, DOC_REL_NS)]
    for i, (title, _cols, _rows) in enumerate(sheets, 1):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                  % (_xml_text(title), i, i))
    wb.append('</sheets></workbook>')

    wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<Relationships xmlns="%s">' % PKG_REL_NS]
    for i in range(1, n + 1):
        wb_rels.append('<Relationship Id="rId%d" Type="%s/worksheet" '
                       'Target="worksheets/sheet%d.xml"/>' % (i, DOC_REL_NS, i))
    wb_rels.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>'
                   % (n + 1, DOC_REL_NS))
    wb_rels.append('</Relationships>')

    # 最小可用样式表：没有它，部分版本的 Excel/WPS 会提示文件需要修复
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<styleSheet xmlns="%s">'
              '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border/></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
              'borderId="0"/></cellStyleXfs>'
              '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
              'borderId="0" xfId="0"/></cellXfs>'
              '<cellStyles count="1"><cellStyle name="常规" xfId="0" '
              'builtinId="0"/></cellStyles>'
              '</styleSheet>') % XLSX_NS

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="%s">'
                 '<Relationship Id="rId1" Type="%s/officeDocument" '
                 'Target="xl/workbook.xml"/></Relationships>') % (PKG_REL_NS, DOC_REL_NS)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        z.writestr("xl/styles.xml", styles)
        for i, (_title, cols, rows) in enumerate(sheets, 1):
            z.writestr("xl/worksheets/sheet%d.xml" % i, _sheet_xml(cols, rows))


def export_excel(path, plat, summary, details):
    """导出为 .xlsx：两张独立工作表，各占一个页签，数据不混合。"""
    create_xlsx(path, [
        (plat.summary_title, plat.export_summary_cols, summary),
        (plat.detail_title, plat.export_detail_cols, details),
    ])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("多平台 CDK 售卖库存查询")
        self.geometry("1120x700")
        self.minsize(900, 560)
        self.config(bg="#f4f6f8")

        self.summary, self.details = [], []
        cfg = load_config()
        key = cfg.get("last_platform") or PLATFORM_ORDER[0]
        self.plat = PLATFORMS.get(key, PLATFORMS[PLATFORM_ORDER[0]])
        self.token = load_token(self.plat)
        self._busy = False

        self._build_ui()

        if self.token:
            self.after(250, self.do_query)
        else:
            self.after(200, self.change_token)

    # ---------- 界面 ---------- #
    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Treeview", font=(FONT, 9), rowheight=26)
        style.configure("Treeview.Heading", font=(FONT, 9, "bold"))
        style.configure("TButton", font=(FONT, 10))
        style.configure("TLabel", font=(FONT, 10))
        style.configure("TNotebook.Tab", font=(FONT, 10), padding=(16, 6))

        top = tk.Frame(self, bg="#f4f6f8")
        top.pack(fill="x", padx=14, pady=(12, 8))

        # 平台选择：切换后自动读取该平台已保存的凭证并查询
        ttk.Label(top, text="平台", background="#f4f6f8").pack(side="left")
        self.var_platform = tk.StringVar(value=self.plat.label)
        self.cmb_platform = ttk.Combobox(
            top, textvariable=self.var_platform, state="readonly", width=12,
            font=(FONT, 10), values=[PLATFORMS[k].label for k in PLATFORM_ORDER])
        self.cmb_platform.pack(side="left", padx=(4, 10))
        self.cmb_platform.bind("<<ComboboxSelected>>", self._on_platform_change)

        self.btn_query = tk.Button(
            top, text="🔍  查询", font=(FONT, 13, "bold"), fg="white",
            bg="#1f9d63", activebackground="#188050", activeforeground="white",
            relief="flat", cursor="hand2", width=10, height=1,
            command=self.do_query)
        self.btn_query.pack(side="left")

        ttk.Button(top, text="导出表格", command=self.export_data).pack(side="left", padx=(10, 0))
        ttk.Button(top, text="更换凭证", command=self.change_token).pack(side="left", padx=(8, 0))

        self.lbl_info = ttk.Label(top, text="账号：—    在库 CDK：—    最后查询：—",
                                  background="#f4f6f8")
        self.lbl_info.pack(side="right")

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(0, 6))

        self.tab1 = ttk.Frame(self.nb)
        self.tab2 = ttk.Frame(self.nb)
        self.nb.add(self.tab1, text=self.plat.summary_title)
        self.nb.add(self.tab2, text=self.plat.detail_title)
        self.tv1 = self._make_tree(self.tab1, self.plat.summary_cols)
        self.tv2 = self._make_tree(self.tab2, self.plat.detail_cols)

        self.status = tk.StringVar(value="就绪。")
        bar = tk.Label(self, textvariable=self.status, anchor="w",
                       font=(FONT, 9), bg="#232a31", fg="#d7dde3", padx=10)
        bar.pack(fill="x", side="bottom")

    def _make_tree(self, parent, cols):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)
        tv = ttk.Treeview(wrap, columns=[c[0] for c in cols], show="headings")
        for name, width in cols:
            tv.heading(name, text=name)
            tv.column(name, width=width, anchor="w", stretch=False)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tv.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="we")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        tv.tag_configure("odd", background="#f7f9fa")
        tv.tag_configure("sold", background="#fdeeee")
        return tv

    def _rebuild_tables(self):
        """切换平台时按新平台的列定义重建两张表，并清空上一平台的数据。"""
        for frame in (self.tab1, self.tab2):
            for child in frame.winfo_children():
                child.destroy()
        self.tv1 = self._make_tree(self.tab1, self.plat.summary_cols)
        self.tv2 = self._make_tree(self.tab2, self.plat.detail_cols)
        self.summary, self.details = [], []
        self.nb.tab(0, text=self.plat.summary_title)
        self.nb.tab(1, text=self.plat.detail_title)
        self.lbl_info.config(text="账号：—    在库 CDK：—    最后查询：—")

    def _on_platform_change(self, event=None):
        if self._busy:
            self.var_platform.set(self.plat.label)
            return
        self.plat = platform_by_label(self.var_platform.get())
        self.token = load_token(self.plat)
        self._rebuild_tables()
        cfg = load_config()
        cfg["last_platform"] = self.plat.key
        save_config(cfg)
        if self.token:
            self.status.set("已切换到 %s，正在查询 ..." % self.plat.label)
            self.do_query()
        else:
            self.status.set("已切换到 %s，请先填写 %s。"
                            % (self.plat.label, self.plat.token_name))
            self.after(150, self.change_token)

    # ---------- 行为 ---------- #
    def change_token(self):
        dlg = TokenDialog(self, self.plat)
        self.wait_window(dlg)
        if dlg.result:
            self.token = dlg.result
            save_token(self.plat, self.token)
            self.do_query()

    def do_query(self):
        if self._busy:
            return
        if not self.token:
            self.change_token()
            return
        self._busy = True
        self.btn_query.config(state="disabled", bg="#9bb8aa")
        self.status.set("查询中 …")

        plat = self.plat

        def worker():
            try:
                nick, summ, det = plat.collect(
                    self.token, progress=lambda t: self.after(0, self.status.set, t))
                self.after(0, self.on_success, nick, summ, det)
            except QueryError as e:
                self.after(0, self.on_error, str(e))
            except Exception as e:
                self.after(0, self.on_error, "发生未知错误：%s" % e)

        threading.Thread(target=worker, daemon=True).start()

    def on_success(self, nick, summ, det):
        # 分区构成先统计，避免受排序影响（排序后国区未必排在第一个）
        bd = _breakdown(det)
        # 默认排序：见各平台的 sort_summary / sort_detail
        # （steamcici 明细按「未出库在前、已售出在后」；steampy 明细按库存降序）
        # 排序结果同时用于表格显示与导出
        summ = _apply_sort(summ, self.plat.sort_summary)
        det = _apply_sort(det, self.plat.sort_detail)
        self.summary, self.details = summ, det
        self._fill(self.tv1, summ, [c[0] for c in self.plat.summary_cols])
        if self.plat.key == "steamcici":
            self._fill(self.tv2, det, [c[0] for c in self.plat.detail_cols],
                       flag_field="库存状态", flag_values=("已售出",))
        else:
            self._fill(self.tv2, det, [c[0] for c in self.plat.detail_cols],
                       flag_field="挂售状态", flag_values=("关闭",))
        self.nb.tab(0, text="%s（%d）" % (self.plat.summary_title, len(summ)))
        self.nb.tab(1, text="%s（%d）" % (self.plat.detail_title, len(det)))
        in_stock = sum(_to_int(r.get("当前库存")) for r in summ)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.lbl_info.config(text="%s · %s    在库：%d    最后查询：%s"
                             % (self.plat.label, nick or "—", in_stock, now))

        # 明细条数容易让人误判，状态栏明确摊开各分区各占多少，
        # 并带上 collect 的附带说明
        parts = ["完成：%d 个商品，共 %d 条明细。" % (len(summ), len(det))]
        if bd:
            parts.append("明细分区构成：%s。" % bd)
        note = self.plat.last_note
        self.plat.last_note = ""
        if note:
            parts.append(note)
        self.status.set("  ".join(parts))
        self._set_idle()

    def on_error(self, msg):
        self.status.set(msg)
        self._set_idle()
        messagebox.showerror("查询失败", msg)
        if "401" in msg:
            self.after(200, self.change_token)

    def _set_idle(self):
        self._busy = False
        self.btn_query.config(state="normal", bg="#1f9d63")

    def _fill(self, tv, rows, cols, flag_field=None, flag_values=()):
        """填充表格；flag_field 指定的列命中 flag_values 时整行标红底。"""
        tv.delete(*tv.get_children())
        for i, r in enumerate(rows):
            tag = "odd" if i % 2 else "even"
            if flag_field and str(r.get(flag_field, "")) in flag_values:
                tag = "sold"
            tv.insert("", "end", values=[r.get(c, "") for c in cols], tags=(tag,))

    def export_data(self):
        if not self.summary and not self.details:
            messagebox.showinfo("提示", "暂无数据，请先点【查询】。")
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            title="保存库存",
            initialfile="%s_库存_%s.xlsx" % (self.plat.key, stamp),
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿 · 两张工作表", "*.xlsx"),
                       ("CSV 文件 · 两张表上下排列", "*.csv")])
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                export_blocks(path, self.plat, self.summary, self.details)
                how = "1 个 CSV 文件（%s %d 行 + %s %d 行）"
            else:
                export_excel(path, self.plat, self.summary, self.details)
                how = "1 个 Excel 工作簿（工作表：%s %d 行 / %s %d 行）"
            note = how % (self.plat.summary_title, len(self.summary),
                          self.plat.detail_title, len(self.details))
            messagebox.showinfo("已导出", "已保存：\n%s\n\n%s" % (path, note))
            self.status.set("已导出 " + note + "。")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


def main():
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        try:
            messagebox.showerror("启动失败", str(e))
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()

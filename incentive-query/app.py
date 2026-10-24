# -*- coding: utf-8 -*-
"""激励金额自助查询 · 后端（仅服务端可见原始数据）。

数据源：data/incentive_sample.xlsx，三个 Sheet 与三张原表列结构完全一致（末尾加一列「月份」）：
  Sheet「试驾激励」：大区|经销商编码|经销商名称|门店类型|车系|上牌日期|销售名字|主岗位|在职|身份证|激励金额|月份
  Sheet「交付激励」：大区|门店类型|经销商代码|经销商名称|员工工号|姓名|身份证|APP订单号|DMS车系名称|车型标准名称|激励金额|月份
  Sheet「终端激励」：大区|序号|店编号|店简称|店类型|总经理名称|总经理电话|总经理身份证|N60终端目标|N60实绩|激励金额|月份

查询：
  店代码   -> 匹配 经销商编码 / 经销商代码 / 店编号，一次查回三类激励
  身份证号 -> 匹配 身份证(试驾/交付) + 总经理身份证(终端)，一次查回本人全部激励

留痕：时间 / 方式 / 查询人 / 店代码或掩码身份证 / 月份 / 是否查到 / 条数 / 合计 / 该键累计第几次
安全：完整身份证只在服务端比对，永不下发浏览器；无任何导出接口。
"""
import os
import re
import csv
import json
import mimetypes
import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import openpyxl
    OPENPYXL_OK = True
except Exception:
    OPENPYXL_OK = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_XLSX = os.environ.get("INCENTIVE_DATA_XLSX",
                           os.path.join(BASE_DIR, "data", "incentive_sample.xlsx"))
DATA_DIR = os.path.join(BASE_DIR, "data")
PORT = int(os.environ.get("PORT", "3000"))

# 每个 Sheet 的列结构（与原表一致 + 末列月份）
SHEETS = {
    "试驾激励": ["大区", "经销商编码", "经销商名称", "门店类型", "车系", "上牌日期",
                 "销售名字", "主岗位", "在职", "身份证", "激励金额", "月份"],
    "交付激励": ["大区", "门店类型", "经销商代码", "经销商名称", "员工工号", "姓名",
                 "身份证", "APP订单号", "DMS车系名称", "车型标准名称", "激励金额", "月份"],
    "终端激励": ["大区", "序号", "店编号", "店简称", "店类型", "总经理名称", "总经理电话",
                 "总经理身份证", "N60终端目标", "N60实绩", "激励金额", "月份"],
}
TYPES = list(SHEETS.keys())

# 统一记录字段
REC_KEYS = ["激励类型", "大区", "店代码", "门店名称", "门店类型", "月份", "姓名",
            "身份证号", "激励金额", "extra"]

_cache = {"mtime": 0.0, "rows": []}
LOG_FILE = os.path.join(DATA_DIR, "query_log.csv")

ID_RE = re.compile(r"^\d{15}$|^\d{17}[\dXx]$")


def mask_id(idno):
    """身份证掩码：前4 + ******** + 后4，用于留痕/门店核对，不暴露完整号。"""
    s = (idno or "").strip().upper()
    if len(s) >= 12:
        return s[:4] + "********" + s[-4:]
    return "****"


def log_query(mode, key, querier, month, found, count, total):
    """留痕：谁（店代码/掩码身份证/查询人）、何时、查到与否、金额、该键累计第几次。"""
    try:
        exist = os.path.exists(LOG_FILE)
        n = 0
        if exist:
            with open(LOG_FILE, newline="", encoding="utf-8-sig") as f:
                for r in csv.reader(f):
                    if len(r) >= 3 and r[2] == key:
                        n += 1
        with open(LOG_FILE, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if not exist:
                w.writerow(["时间", "查询方式", "店代码/身份证(掩码)", "查询人", "月份",
                            "是否查到", "条数", "合计金额", "该键累计第几次"])
            w.writerow([
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "门店" if mode == "store" else "个人",
                key, querier or "（未填）", month or "全部",
                "是" if found else "否", count, total, n + 1,
            ])
    except Exception:
        pass


def _norm(v):
    return ("" if v is None else str(v).strip())


def load_rows(force=False):
    """读 xlsx 三个 Sheet（按 Sheet 名，找不到则按第1/2/3个），合成为统一记录。
    xlsx 不可用时回退读 data/试驾激励.csv 等三个同名 csv。"""
    src = DATA_XLSX if (OPENPYXL_OK and os.path.exists(DATA_XLSX)) else None
    if src is None:
        src = {t: os.path.join(DATA_DIR, f"{t}.csv") for t in TYPES}
        existing = {t: p for t, p in src.items() if os.path.exists(p)}
        if not existing:
            return []
        mtime = max(os.path.getmtime(p) for p in existing.values())
    else:
        mtime = os.path.getmtime(src)
    if not force and mtime == _cache["mtime"] and _cache["rows"]:
        return _cache["rows"]

    rows = []
    for t in TYPES:
        if src is None or isinstance(src, dict):
            path = src.get(t)
            if not path or not os.path.exists(path):
                continue
            with open(path, newline="", encoding="utf-8-sig") as f:
                raw = list(csv.reader(f))
        else:
            wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
            if t in wb.sheetnames:
                ws = wb[t]
            else:
                idx = TYPES.index(t)
                ws = wb[wb.sheetnames[idx]] if idx < len(wb.sheetnames) else None
            if ws is None:
                wb.close()
                continue
            raw = [r for r in ws.iter_rows(values_only=True)]
            wb.close()
        if not raw:
            continue
        header = [_norm(h) for h in raw[0]]
        cols = SHEETS[t]
        idx = {c: header.index(c) if c in header else None for c in cols}
        for r in raw[1:]:
            if r is None or not any(c not in (None, "") for c in r):
                continue
            rec = {c: _norm(r[i]) if (i is not None and i < len(r)) else "" for c, i in idx.items()}
            try:
                rec["激励金额"] = float(rec["激励金额"])
            except (ValueError, TypeError):
                rec["激励金额"] = 0.0
            if t == "试驾激励":
                uni = {"店代码": rec["经销商编码"], "门店名称": rec["经销商名称"],
                       "姓名": rec["销售名字"], "身份证号": rec["身份证"]}
                extra = {"车系": rec["车系"], "上牌日期": rec["上牌日期"],
                         "主岗位": rec["主岗位"], "在职": rec["在职"]}
            elif t == "交付激励":
                uni = {"店代码": rec["经销商代码"], "门店名称": rec["经销商名称"],
                       "姓名": rec["姓名"], "身份证号": rec["身份证"]}
                extra = {"员工工号": rec["员工工号"], "APP订单号": rec["APP订单号"],
                         "DMS车系名称": rec["DMS车系名称"], "车型标准名称": rec["车型标准名称"]}
            else:  # 终端激励
                uni = {"店代码": rec["店编号"], "门店名称": rec["店简称"],
                       "姓名": rec["总经理名称"], "身份证号": rec["总经理身份证"]}
                extra = {"N60终端目标": rec["N60终端目标"], "N60实绩": rec["N60实绩"],
                         "总经理名称": rec["总经理名称"], "总经理电话": rec["总经理电话"]}
            rows.append({
                "激励类型": t, "大区": rec["大区"],
                "门店类型": rec.get("门店类型", "") or rec.get("店类型", ""),
                "月份": rec["月份"], "激励金额": rec["激励金额"],
                "店代码": uni["店代码"], "门店名称": uni["门店名称"],
                "姓名": uni["姓名"], "身份证号": uni["身份证号"], "extra": extra,
            })
    _cache["mtime"] = mtime
    _cache["rows"] = rows
    return rows


def query(store_code=None, idno=None, month=None):
    rows = load_rows()
    if store_code:
        code = store_code.strip().upper()
        rows = [r for r in rows if r["店代码"].strip().upper() == code]
    else:
        idv = idno.strip().upper()
        rows = [r for r in rows if r["身份证号"].strip().upper() == idv]
    if month:
        rows = [r for r in rows if r["月份"] == month]
    store_name = rows[0]["门店名称"] if rows else ""
    return store_name, rows


def row_payload(r, mode):
    """下发到前端的行：永不包含完整身份证号；门店模式给掩码便于店内核对。"""
    p = {"月份": r["月份"], "姓名": r["姓名"], "激励金额": r["激励金额"]}
    if mode == "store":
        p["身份证(掩码)"] = mask_id(r["身份证号"])
    p.update(r["extra"])
    return p


def build_groups(matched, mode):
    groups = []
    for t in TYPES:
        rs = [r for r in matched if r["激励类型"] == t]
        if not rs:
            continue
        groups.append({
            "type": t,
            "count": len(rs),
            "subtotal": round(sum(r["激励金额"] for r in rs), 2),
            "rows": [row_payload(r, mode) for r in rs],
        })
    return groups


class Handler(BaseHTTPRequestHandler):
    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, status=200):
        try:
            with open(path, "rb") as f:
                body = f.read()
        except OSError:
            self._send_json({"error": "not found"}, 404)
            return
        ctype, _ = mimetypes.guess_type(path)
        ctype = ctype or "application/octet-stream"
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send_file(os.path.join(BASE_DIR, "index.html"))
        elif self.path == "/healthz":
            self._send_json({"ok": True, "openpyxl": OPENPYXL_OK, "rows": len(load_rows())})
        else:
            # 企微可信域名校验文件（如 WW_verify_xxxx.txt）放在 public/ 下，按根路径提供
            fname = self.path.split("?", 1)[0].lstrip("/")
            pub = os.path.join(BASE_DIR, "public", fname)
            if fname and os.path.isfile(pub):
                self._send_file(pub)
            else:
                self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        if self.path != "/api/query":
            self._send_json({"error": "not found"}, 404)
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            data = json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            self._send_json({"error": "bad request"}, 400)
            return

        key = (data.get("key", "") or "").strip()
        querier = (data.get("querier", "") or "").strip()[:20]
        month = (data.get("month", "") or "").strip() or None
        if not key:
            self._send_json({"found": False, "reason": "need_key"}, 200)
            return

        if ID_RE.match(key):
            mode, idno, store_code = "person", key.upper(), None
        else:
            mode, idno, store_code = "store", None, key

        store_name, matched = query(store_code=store_code, idno=idno, month=month)
        groups = build_groups(matched, mode)
        total = round(sum(r["激励金额"] for r in matched), 2)
        log_query(mode, mask_id(idno) if idno else key.upper(), querier,
                  month, len(matched) > 0, len(matched), total)
        self._send_json({
            "found": len(matched) > 0,
            "mode": mode,
            "storeName": store_name,
            "groups": groups,
            "total": total,
            "count": len(matched),
        })

    def log_message(self, *args):
        pass


if __name__ == "__main__":
    try:
        load_rows(force=True)
    except Exception as e:
        print("warn: data load failed:", e)
    print(f"serving on :{PORT}  openpyxl={OPENPYXL_OK}")
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()

# -*- coding: utf-8 -*-
"""生成自包含本地预览 HTML（示例数据内置，无需服务器，双击即可用）。
仅用于本地演示/验证交互；正式多人查询仍用 server 版（数据在服务端）。
"""
import os, csv, json

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_XLSX = os.path.join(BASE, "data", "incentive_sample.xlsx")
DATA_CSV = os.path.join(BASE, "data", "incentive_sample.csv")
OUT = os.path.join(BASE, "local_preview.html")

COLUMNS = ["店代码", "门店名称", "姓名", "身份证后6位", "月份",
           "激励类型", "金额", "台数", "状态"]

rows = []
# 优先 xlsx
try:
    import openpyxl
    if os.path.exists(DATA_XLSX):
        wb = openpyxl.load_workbook(DATA_XLSX, read_only=True, data_only=True)
        raw = [r for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    else:
        raw = None
except Exception:
    raw = None

if not raw:
    with open(DATA_CSV, newline="", encoding="utf-8-sig") as f:
        raw = list(csv.reader(f))

header = raw[0]
idx = {str(h).strip(): i for i, h in enumerate(header)}
for r in raw[1:]:
    if r is None or not any(c not in (None, "") for c in r):
        continue
    rec = {}
    for col in COLUMNS:
        i = idx.get(col)
        v = r[i] if (i is not None and i < len(r)) else None
        rec[col] = ("" if v is None else str(v).strip())
    try:
        rec["金额"] = float(rec["金额"])
    except (ValueError, TypeError):
        rec["金额"] = 0.0
    try:
        rec["台数"] = int(float(rec["台数"])) if rec["台数"] not in ("", None) else 0
    except (ValueError, TypeError):
        rec["台数"] = 0
    rows.append(rec)

data_json = json.dumps(rows, ensure_ascii=False)

html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>激励金额自助查询（本地演示版）</title>
<style>
  :root{
    --blue:#185FA5; --blue-50:#E6F1FB; --green:#3B6D11; --text:#1F2329;
    --muted:#6B7280; --border:#D0D5DD; --bg:#F7F8FA;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"PingFang SC","Microsoft YaHei",system-ui,sans-serif;background:var(--bg);color:var(--text);padding:16px;line-height:1.6}
  .wrap{max-width:560px;margin:0 auto}
  header{background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px}
  header h1{font-size:18px;font-weight:600}
  header p{font-size:12px;color:var(--muted);margin-top:4px}
  .tabs{display:flex;gap:8px;margin-bottom:12px}
  .tab{flex:1;text-align:center;padding:10px;border:1px solid var(--border);border-radius:10px;background:#fff;cursor:pointer;font-size:14px;color:var(--muted)}
  .tab.active{background:var(--blue);color:#fff;border-color:var(--blue)}
  .card{background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px}
  label{display:block;font-size:13px;color:var(--muted);margin:10px 0 4px}
  input{width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:8px;font-size:14px;outline:none}
  input:focus{border-color:var(--blue)}
  .btn{margin-top:14px;width:100%;padding:12px;background:var(--blue);color:#fff;border:none;border-radius:10px;font-size:15px;cursor:pointer}
  .result{margin-top:14px;background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px;display:none}
  .result.show{display:block}
  .result h2{font-size:15px;margin-bottom:8px}
  .total{font-size:26px;font-weight:700;color:var(--green);margin:6px 0 12px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{text-align:left;padding:8px 6px;border-bottom:1px solid #eee}
  th{color:var(--muted);font-weight:500}
  .empty{color:var(--muted);font-size:14px;padding:10px 0}
  .warn{background:#FFF7ED;border:1px solid #FDBA74;color:#9A3412;border-radius:10px;padding:10px;font-size:13px;margin-top:12px}
  .privacy{font-size:11px;color:var(--muted);margin-top:14px;text-align:center}
  .print-btn{margin-top:14px;width:100%;padding:11px;background:#fff;color:var(--blue);border:1px solid var(--blue);border-radius:10px;font-size:14px;cursor:pointer}
  .demo{background:#FFFBEA;border:1px solid #FDE68A;color:#92400E;border-radius:10px;padding:10px;font-size:12px;margin-top:12px}
  #resMeta{font-size:12px;color:var(--muted);margin:6px 0 4px}
  @media print{
    body{padding:0;background:#fff}
    header,.tabs,.card,.privacy,.print-btn,.demo{display:none !important}
    .wrap{max-width:100%}
    .result{display:block !important;border:none;padding:0}
    .total{font-size:30px}
    #resMeta{display:block !important}
  }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>激励金额自助查询</h1>
    <p>输入校验值即可查询本人 / 本店当月激励</p>
  </header>

  <div class="tabs">
    <div class="tab active" id="tabStore" onclick="switchTab('store')">门店查询</div>
    <div class="tab" id="tabPerson" onclick="switchTab('person')">个人查询</div>
  </div>

  <div class="card">
    <div id="storeForm">
      <label>店代码</label>
      <input id="storeCode" placeholder="例如 GZ001">
      <label>月份（可选，留空查全部）</label>
      <input id="storeMonth" placeholder="例如 2026-10">
    </div>
    <div id="personForm" style="display:none">
      <label>姓名</label>
      <input id="name" placeholder="请输入本人姓名">
      <label>身份证后 6 位</label>
      <input id="idLast6" placeholder="用于校验身份" inputmode="numeric">
      <label>月份（可选，留空查全部）</label>
      <input id="personMonth" placeholder="例如 2026-10">
    </div>
    <button class="btn" onclick="doQuery()">查 询</button>
  </div>

  <div class="demo">本地演示版：数据为示例，内置供你验证交互。正式多人/公网查询由服务端版本提供（数据不落地到浏览器）。</div>

  <div class="result" id="result">
    <h2 id="resTitle"></h2>
    <div class="total" id="resTotal"></div>
    <div id="resMeta"></div>
    <div id="resBody"></div>
    <button class="print-btn" id="printBtn" style="display:none" onclick="window.print()">打印 / 导出 PDF</button>
  </div>

  <div class="privacy">演示版数据仅内置在本地文件中，不上传任何服务器。</div>
</div>

<script>
const DATA = __DATA__;
let mode = "store";
function switchTab(m){
  mode = m;
  document.getElementById("tabStore").classList.toggle("active", m==="store");
  document.getElementById("tabPerson").classList.toggle("active", m==="person");
  document.getElementById("storeForm").style.display = m==="store" ? "block":"none";
  document.getElementById("personForm").style.display = m==="person" ? "block":"none";
}
function doQuery(){
  const el = document.getElementById("result");
  el.classList.remove("show");
  let matched = [];
  if(mode==="store"){
    const sc = document.getElementById("storeCode").value.trim().toUpperCase();
    const mo = document.getElementById("storeMonth").value.trim();
    if(!sc){ alert("请输入店代码"); return; }
    matched = DATA.filter(r => r["店代码"].toUpperCase()===sc && r["姓名"]==="");
    if(mo) matched = matched.filter(r => r["月份"]===mo);
  } else {
    const name = document.getElementById("name").value.trim();
    const last6 = document.getElementById("idLast6").value.trim();
    const mo = document.getElementById("personMonth").value.trim();
    if(!name || !last6){ alert("请输入姓名和身份证后6位"); return; }
    matched = DATA.filter(r => r["姓名"]!=="" && r["姓名"]===name && r["身份证后6位"]===last6);
    if(mo) matched = matched.filter(r => r["月份"]===mo);
  }
  show(matched, mode);
}
function show(matched, m){
  const el = document.getElementById("result");
  el.classList.add("show");
  if(matched.length===0){
    document.getElementById("resTitle").textContent="未查询到结果";
    document.getElementById("resTotal").textContent="";
    document.getElementById("resBody").innerHTML='<div class="warn">未匹配到对应记录。请核对校验值是否正确；仍查不到请联系渠道赋能科核对。</div>';
    return;
  }
  const storeName = matched[0]["门店名称"] || "（店名未知）";
  const total = matched.reduce((s,r)=>s+(+r["金额"]||0),0);
  document.getElementById("resTitle").textContent = (m==="person"?"个人":"门店")+" · "+storeName;
  document.getElementById("resTotal").textContent = "合计 ¥"+Number(total).toLocaleString("zh-CN",{minimumFractionDigits:2,maximumFractionDigits:2});
  let html = '<table><thead><tr><th>月份</th><th>激励类型</th><th>金额</th><th>台数</th><th>状态</th></tr></thead><tbody>';
  matched.forEach(r=>{
    html += '<tr><td>'+r["月份"]+'</td><td>'+r["激励类型"]+'</td><td>¥'+Number(r["金额"]).toLocaleString("zh-CN")+'</td><td>'+r["台数"]+'</td><td>'+r["状态"]+'</td></tr>';
  });
  html += "</tbody></table>";
  document.getElementById("resBody").innerHTML = html;
  document.getElementById("resMeta").textContent = "查询时间：" + new Date().toLocaleString("zh-CN");
  document.getElementById("printBtn").style.display = "block";
}
</script>
</body>
</html>
"""

html = html.replace("__DATA__", data_json)
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print("wrote", OUT, "rows=", len(rows))

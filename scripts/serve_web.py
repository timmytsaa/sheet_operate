"""sheetops 網頁版（Phase A 極簡版）：上傳 xlsx ＋ 繁中指令 → 預覽變更 → 確認下載。

啟動：
  python scripts/serve_web.py            # http://localhost:8033，區網同事可用 http://<你的IP>:8033
選項（環境變數）：
  SHEETOPS_MODEL=sheetops  SHEETOPS_PORT=8033  OLLAMA_HOST=http://localhost:11434

設計要點：
- 完全重用 sheetops 管線；原始檔案永不覆寫，確認後才提供修改版下載
- 單一 GPU：全域鎖，一次處理一個請求（其他請求排隊）
- 使用紀錄 logs/usage_log.jsonl：指令/程式碼/差異統計/採納與否——v3 任務規格與 DPO 的原料
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from sheetops.ollama_client import OllamaClient
from sheetops.pipeline import process_workbook

MODEL = os.environ.get("SHEETOPS_MODEL", "sheetops")
PORT = int(os.environ.get("SHEETOPS_PORT", "8033"))
LOG_PATH = ROOT / "logs" / "usage_log.jsonl"
LOG_PATH.parent.mkdir(exist_ok=True)

app = FastAPI(title="sheetops")
client = OllamaClient(host=os.environ.get("OLLAMA_HOST", "http://localhost:11434"), model=MODEL)
GPU_LOCK = threading.Lock()
SESSIONS: dict[str, dict] = {}
SESSION_TTL = 3600


def log_event(record: dict) -> None:
    record["ts"] = time.strftime("%Y-%m-%d %H:%M:%S")
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def cleanup_sessions() -> None:
    now = time.time()
    for sid in [s for s, v in SESSIONS.items() if now - v["created"] > SESSION_TTL]:
        shutil.rmtree(SESSIONS[sid]["dir"], ignore_errors=True)
        SESSIONS.pop(sid, None)


@app.post("/api/process")
async def api_process(file: UploadFile = File(...), instruction: str = Form(...),
                      context: str = Form("")):
    cleanup_sessions()
    if not file.filename.lower().endswith(".xlsx"):
        return JSONResponse({"ok": False, "error": "僅支援 .xlsx 檔案"})
    data = await file.read()
    if len(data) > 20 * 1024 * 1024:
        return JSONResponse({"ok": False, "error": "檔案超過 20MB 上限"})

    sid = uuid.uuid4().hex[:12]
    sdir = Path(tempfile.mkdtemp(prefix=f"sheetops_web_{sid}_"))
    src = sdir / file.filename
    src.write_bytes(data)

    t0 = time.monotonic()
    with GPU_LOCK:  # 單卡：一次一個請求
        result = process_workbook(
            src, instruction, context,
            generate_fn=lambda msgs: client.chat(msgs, temperature=0.0),
            retries=1)
    seconds = round(time.monotonic() - t0, 1)

    base = {"sid": sid, "file": file.filename, "instruction": instruction,
            "context": context, "model": MODEL, "seconds": seconds,
            "attempts": result["attempts"], "code": result["code"]}
    if not result["ok"]:
        log_event({**base, "event": "process", "ok": False, "error": result["error"]})
        shutil.rmtree(sdir, ignore_errors=True)
        return JSONResponse({"ok": False, "error": result["error"], "seconds": seconds})

    final = sdir / ("已修改_" + file.filename)
    shutil.move(result["result_path"], final)
    changed = sum(v["changed"] for v in result["diff"]["sheets"].values())
    SESSIONS[sid] = {"dir": sdir, "file": final, "name": final.name,
                     "created": time.time()}
    log_event({**base, "event": "process", "ok": True, "changed_cells": changed,
               "sheets_added": result["diff"]["sheets_added"]})
    return JSONResponse({"ok": True, "sid": sid, "seconds": seconds,
                         "diff_text": result["diff_text"], "code": result["code"]})


@app.get("/api/download/{sid}")
async def api_download(sid: str):
    s = SESSIONS.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期，請重新處理"})
    log_event({"event": "decision", "sid": sid, "decision": "accepted"})
    return FileResponse(s["file"], filename=s["name"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/reject/{sid}")
async def api_reject(sid: str):
    s = SESSIONS.pop(sid, None)
    if s:
        shutil.rmtree(s["dir"], ignore_errors=True)
    log_event({"event": "decision", "sid": sid, "decision": "rejected"})
    return JSONResponse({"ok": True})


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


HTML_PAGE = """<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>試算表助理</title>
<style>
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body { font-family: "Segoe UI", "Microsoft JhengHei", system-ui, sans-serif;
         background: #f2f4f8; margin: 0; padding: 24px; color: #1a2233; }
  .card { max-width: 760px; margin: 0 auto 16px; background: #fff; border-radius: 12px;
          padding: 24px; box-shadow: 0 1px 4px rgba(20,30,60,.08); }
  h1 { font-size: 20px; margin: 0 0 4px; }
  .sub { color: #67718a; font-size: 13px; margin-bottom: 18px; }
  label { display: block; font-weight: 600; font-size: 14px; margin: 14px 0 6px; }
  input[type=file] { font-size: 14px; }
  textarea { width: 100%; border: 1px solid #ccd3e0; border-radius: 8px; padding: 10px;
             font-size: 15px; font-family: inherit; resize: vertical; }
  #instruction { min-height: 72px; } #context { min-height: 48px; }
  button { border: 0; border-radius: 8px; padding: 10px 22px; font-size: 15px;
           cursor: pointer; font-family: inherit; }
  .primary { background: #2456d6; color: #fff; } .primary:disabled { background: #9db1e0; }
  .ghost { background: #e8ecf5; color: #1a2233; margin-left: 8px; }
  #status { margin-top: 14px; font-size: 14px; color: #67718a; }
  pre { background: #f6f8fb; border: 1px solid #e3e8f2; border-radius: 8px; padding: 14px;
        font-size: 13px; white-space: pre-wrap; word-break: break-all; max-height: 340px; overflow: auto; }
  details { margin-top: 10px; } summary { cursor: pointer; color: #67718a; font-size: 13px; }
  .hidden { display: none; }
  .footer { text-align: center; color: #9aa3b8; font-size: 12px; margin-top: 8px; }
</style></head><body>
<div class="card">
  <h1>📊 試算表助理</h1>
  <div class="sub">上傳 Excel、用一句話描述要做什麼——預覽變更、確認後下載。原始檔案不會被更動。</div>
  <label>Excel 檔案（.xlsx）</label>
  <input type="file" id="file" accept=".xlsx">
  <label>操作指令</label>
  <textarea id="instruction" placeholder="例：把金額低於 5000 的列刪掉，照金額由大到小排序，底部加總計列"></textarea>
  <label>補充說明（選填：你們單位的自訂規則）</label>
  <textarea id="context" placeholder="例：含稅價 = 金額 × 1.05，四捨五入取整數"></textarea>
  <div style="margin-top:16px">
    <button class="primary" id="go" onclick="run()">執行</button>
  </div>
  <div id="status"></div>
</div>
<div class="card hidden" id="result">
  <h1>變更預覽</h1>
  <pre id="diff"></pre>
  <details><summary>檢視產生的程式碼</summary><pre id="code"></pre></details>
  <div style="margin-top:16px">
    <button class="primary" onclick="accept()">✅ 確認並下載</button>
    <button class="ghost" onclick="reject()">✖ 放棄</button>
  </div>
</div>
<div class="footer">程式碼經安全檢查後於沙盒執行・所有操作留有紀錄</div>
<script>
let sid = null;
const $ = id => document.getElementById(id);
async function run() {
  const f = $("file").files[0];
  const inst = $("instruction").value.trim();
  if (!f) { $("status").textContent = "請先選擇 .xlsx 檔案"; return; }
  if (!inst) { $("status").textContent = "請輸入操作指令"; return; }
  $("go").disabled = true;
  $("result").classList.add("hidden");
  $("status").textContent = "⏳ 模型處理中（依表格大小約 20 秒～2 分鐘，若有人排隊會更久）…";
  const fd = new FormData();
  fd.append("file", f); fd.append("instruction", inst); fd.append("context", $("context").value);
  try {
    const r = await fetch("/api/process", { method: "POST", body: fd });
    const j = await r.json();
    if (!j.ok) { $("status").textContent = "❌ 處理失敗：" + j.error; return; }
    sid = j.sid;
    $("status").textContent = `完成（${j.seconds} 秒）——請檢查下方變更，確認無誤再下載。`;
    $("diff").textContent = j.diff_text;
    $("code").textContent = j.code;
    $("result").classList.remove("hidden");
  } catch (e) { $("status").textContent = "❌ 連線錯誤：" + e; }
  finally { $("go").disabled = false; }
}
function accept() { if (sid) window.location = "/api/download/" + sid; }
async function reject() {
  if (sid) await fetch("/api/reject/" + sid, { method: "POST" });
  sid = null; $("result").classList.add("hidden");
  $("status").textContent = "已放棄此次變更（原始檔案本來就沒動）。";
}
</script></body></html>"""


if __name__ == "__main__":
    import socket

    import uvicorn
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except OSError:
        ip = "127.0.0.1"
    print(f"sheetops 網頁版啟動：http://localhost:{PORT}（區網：http://{ip}:{PORT}）")
    print(f"模型：{MODEL} @ {client.host}；使用紀錄：{LOG_PATH}")
    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="warning")

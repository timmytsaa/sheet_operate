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

from sheetops.ollama_client import OllamaClient, _load_dotenv
from sheetops.pipeline import process_workbook

_load_dotenv()          # 讓 .env 的 SHEETOPS_* 設定生效（必須在讀 os.environ 之前）
from sheetops.univer_io import apply_edit, workbook_to_snapshot

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pro_page import HTML_PRO  # noqa: E402  （Univer 版 /pro 頁面）

MODEL = os.environ.get("SHEETOPS_MODEL", "sheetops")
PORT = int(os.environ.get("SHEETOPS_PORT", "8033"))
LOG_PATH = ROOT / "logs" / "usage_log.jsonl"
LOG_PATH.parent.mkdir(exist_ok=True)

GEN_TIMEOUT = int(os.environ.get("SHEETOPS_TIMEOUT", "300"))   # 單次生成上限（秒）

# 部署用的 Ollama 位址刻意「不」讀 OLLAMA_HOST——那個是訓練/蒸餾用的雲端位址
# （.env 裡通常是 https://ollama.com），網頁版要的是本機那台。
OLLAMA_URL = os.environ.get("SHEETOPS_OLLAMA_HOST") or "http://localhost:11434"

app = FastAPI(title="sheetops")
client = OllamaClient(host=OLLAMA_URL, model=MODEL, timeout=GEN_TIMEOUT)


def preflight() -> None:
    """啟動前確認 Ollama 活著、而且要用的模型真的存在。"""
    import requests
    try:
        tags = requests.get(f"{OLLAMA_URL}/api/tags", timeout=10).json()
        names = [m.get("name", "") for m in tags.get("models", [])]
    except Exception as e:
        print(f"  ⚠ 連不上 Ollama（{OLLAMA_URL}）：{type(e).__name__}")
        print("    請確認 Ollama 已啟動；或用 SHEETOPS_OLLAMA_HOST 指定位址。")
        return
    if any(n == MODEL or n.split(":")[0] == MODEL for n in names):
        print(f"  模型「{MODEL}」已就緒 @ {OLLAMA_URL}")
    else:
        print(f"  ⚠ {OLLAMA_URL} 上找不到模型「{MODEL}」")
        print(f"    現有模型：{', '.join(names[:8]) or '(無)'}")
        print(f"    建立方式：ollama create {MODEL} -f deploy/Modelfile")


MAX_GEN_TOKENS = int(os.environ.get("SHEETOPS_MAX_TOKENS", "1400"))


def generate(msgs):
    """互動用生成：限制輸出長度（防重複生成迴圈拖到逾時）、不重試。"""
    return client.chat(msgs, temperature=0.0, max_tokens=MAX_GEN_TOKENS, retries=1)
GPU_LOCK = threading.Lock()
SESSIONS: dict[str, dict] = {}
SESSION_TTL = 3600


def log_event(record: dict) -> None:
    record["ts"] = time.strftime("%Y-%m-%d %H:%M:%S")
    record["source"] = "web"
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
            generate_fn=generate, retries=1)
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
                         "diff_text": result["diff_text"], "code": result["code"],
                         "inference": result.get("inference", ""),
                         "sheets_added": result["diff"]["sheets_added"],
                         "diff_coords": {name: v["coords"]
                                         for name, v in result["diff"]["sheets"].items()}})


@app.get("/api/download/{sid}")
def api_download(sid: str):
    s = SESSIONS.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期，請重新處理"})
    log_event({"event": "decision", "sid": sid, "decision": "accepted"})
    return FileResponse(s["file"], filename=s["name"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/reject/{sid}")
def api_reject(sid: str):
    s = SESSIONS.pop(sid, None)
    if s:
        shutil.rmtree(s["dir"], ignore_errors=True)
    log_event({"event": "decision", "sid": sid, "decision": "rejected"})
    return JSONResponse({"ok": True})


@app.get("/api/result/{sid}")
def api_result(sid: str):
    """取回結果檔載入網格預覽（不記入採納決定）。"""
    s = SESSIONS.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期"}, status_code=404)
    return FileResponse(s["file"], filename=s["name"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/accept/{sid}")
def api_accept(sid: str):
    """SpreadJS 版的網格內採納（不下載檔案也算一次採納）。"""
    if sid in SESSIONS:
        log_event({"event": "decision", "sid": sid, "decision": "accepted"})
        return JSONResponse({"ok": True})
    return JSONResponse({"ok": False, "error": "session 已過期"}, status_code=404)


@app.post("/api/edit_event")
def api_edit_event(payload: dict):
    """SpreadJS 手動編輯事件——使用者修正紀錄（未來 DPO 的原料）。"""
    log_event({"event": "user_edit", **{k: payload.get(k) for k in
               ("sid", "sheet", "row", "col", "old", "new")}})
    return JSONResponse({"ok": True})


@app.get("/api/sample")
def api_sample():
    p = ROOT / "samples" / "測試活頁簿.xlsx"
    if not p.exists():
        return JSONResponse({"ok": False, "error": "找不到範例檔"}, status_code=404)
    return FileResponse(p, filename=p.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ===== /pro（Univer）：伺服器權威工作簿 session =====
WB: dict[str, dict] = {}


def _wb_cleanup() -> None:
    now = time.time()
    for sid in [s for s, v in WB.items() if now - v["created"] > SESSION_TTL]:
        shutil.rmtree(WB[sid]["dir"], ignore_errors=True)
        WB.pop(sid, None)


def _wb_new_session(data: bytes, name: str) -> dict:
    sid = uuid.uuid4().hex[:12]
    sdir = Path(tempfile.mkdtemp(prefix=f"sheetops_wb_{sid}_"))
    current = sdir / "current.xlsx"
    current.write_bytes(data)
    WB[sid] = {"dir": sdir, "current": current, "pending": None,
               "name": name, "created": time.time()}
    return {"ok": True, "sid": sid, "name": name,
            "snapshot": workbook_to_snapshot(current)}


@app.post("/api/wb/open")
async def wb_open(file: UploadFile = File(...)):
    _wb_cleanup()
    if not file.filename.lower().endswith(".xlsx"):
        return JSONResponse({"ok": False, "error": "僅支援 .xlsx 檔案"})
    data = await file.read()
    if len(data) > 20 * 1024 * 1024:
        return JSONResponse({"ok": False, "error": "檔案超過 20MB 上限"})
    resp = _wb_new_session(data, file.filename)
    log_event({"event": "wb_open", "sid": resp["sid"], "file": file.filename, "ui": "pro"})
    return JSONResponse(resp)


@app.get("/api/wb/sample")
def wb_sample():
    p = ROOT / "samples" / "測試活頁簿.xlsx"
    if not p.exists():
        return JSONResponse({"ok": False, "error": "找不到範例檔"})
    resp = _wb_new_session(p.read_bytes(), p.name)
    log_event({"event": "wb_open", "sid": resp["sid"], "file": "(範例)", "ui": "pro"})
    return JSONResponse(resp)


@app.post("/api/wb/{sid}/edit")
def wb_edit(sid: str, payload: dict):
    s = WB.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期"}, status_code=404)
    target = s["pending"] if (payload.get("target") == "pending" and s["pending"]) else s["current"]
    ok = apply_edit(target, payload.get("sheet"), int(payload.get("row", 0)),
                    int(payload.get("col", 0)), payload.get("value"))
    log_event({"event": "user_edit", "sid": sid, "ui": "pro",
               **{k: payload.get(k) for k in ("sheet", "row", "col", "value", "target")}})
    return JSONResponse({"ok": ok})


@app.post("/api/wb/{sid}/process")
def wb_process(sid: str, payload: dict):
    s = WB.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期，請重新開啟檔案"}, status_code=404)
    if s["pending"]:
        return JSONResponse({"ok": False, "error": "請先採納或還原上一次的變更"})
    instruction = (payload.get("instruction") or "").strip()
    if not instruction:
        return JSONResponse({"ok": False, "error": "指令是空的"})
    context = payload.get("context") or ""

    t0 = time.monotonic()
    try:
        with GPU_LOCK:
            result = process_workbook(
                s["current"], instruction, context,
                generate_fn=generate, retries=1)
    except Exception as e:      # 模型逾時、連線失敗等都回 JSON，別讓前端收到 500 HTML
        seconds = round(time.monotonic() - t0, 1)
        msg = f"{type(e).__name__}: {e}"
        log_event({"event": "process", "ok": False, "sid": sid, "ui": "pro",
                   "instruction": instruction, "error": msg, "seconds": seconds})
        return JSONResponse({"ok": False, "seconds": seconds,
                             "error": f"處理失敗（{seconds} 秒）：{msg}"})
    seconds = round(time.monotonic() - t0, 1)

    base = {"sid": sid, "file": s["name"], "instruction": instruction,
            "context": context, "model": MODEL, "seconds": seconds,
            "attempts": result["attempts"], "code": result["code"], "ui": "pro"}
    if not result["ok"]:
        log_event({**base, "event": "process", "ok": False, "error": result["error"]})
        return JSONResponse({"ok": False, "error": result["error"], "seconds": seconds})

    pending = s["dir"] / "pending.xlsx"
    shutil.move(str(result["result_path"]), str(pending))
    s["pending"] = pending
    changed = sum(v["changed"] for v in result["diff"]["sheets"].values())
    log_event({**base, "event": "process", "ok": True, "changed_cells": changed,
               "sheets_added": result["diff"]["sheets_added"]})
    return JSONResponse({
        "ok": True, "sid": sid, "seconds": seconds,
        "diff_text": result["diff_text"], "code": result["code"],
        "inference": result.get("inference", ""),
        "diff_coords": {name: v["coords"] for name, v in result["diff"]["sheets"].items()},
        "snapshot": workbook_to_snapshot(pending)})


@app.post("/api/wb/{sid}/accept")
def wb_accept(sid: str):
    s = WB.get(sid)
    if not s or not s["pending"]:
        return JSONResponse({"ok": False, "error": "沒有待決定的變更"})
    Path(s["pending"]).replace(s["current"])
    s["pending"] = None
    log_event({"event": "decision", "sid": sid, "decision": "accepted", "ui": "pro"})
    return JSONResponse({"ok": True, "snapshot": workbook_to_snapshot(s["current"])})


@app.post("/api/wb/{sid}/reject")
def wb_reject(sid: str):
    s = WB.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期"}, status_code=404)
    if s["pending"]:
        Path(s["pending"]).unlink(missing_ok=True)
        s["pending"] = None
    log_event({"event": "decision", "sid": sid, "decision": "rejected", "ui": "pro"})
    return JSONResponse({"ok": True, "snapshot": workbook_to_snapshot(s["current"])})


@app.get("/api/wb/{sid}/download")
def wb_download(sid: str):
    s = WB.get(sid)
    if not s:
        return JSONResponse({"ok": False, "error": "session 已過期"}, status_code=404)
    log_event({"event": "download", "sid": sid, "ui": "pro"})
    return FileResponse(s["current"], filename="已修改_" + s["name"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


NO_CACHE = {"Cache-Control": "no-store, must-revalidate", "Pragma": "no-cache"}


@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse(HTML_PAGE, headers=NO_CACHE)


@app.get("/pro", response_class=HTMLResponse)
def pro():
    return HTMLResponse(HTML_PRO, headers=NO_CACHE)


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
  <div class="sub">上傳 Excel、用一句話描述要做什麼——預覽變更、確認後下載。原始檔案不會被更動。
    <a href="/pro" style="float:right">進階版（表格編輯）→</a></div>
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
  <div id="inferbox" style="display:none;background:#eef4ff;border:1px solid #c9d8f5;
       border-radius:8px;padding:10px 12px;margin-bottom:12px;font-size:13px">
    <b>模型的理解</b><br><span id="infer"></span></div>
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
    if (j.inference) { $("infer").textContent = j.inference; $("inferbox").style.display = "block"; }
    else { $("inferbox").style.display = "none"; }
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

    def lan_ips() -> list[str]:
        """列出所有可能的區網 IP；預設路由那張網卡排最前面（給同事的網址用它）。"""
        ips = []
        try:                                  # 走預設路由的那一張（不會真的連出去）
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ips.append(s.getsockname()[0])
            s.close()
        except OSError:
            pass
        try:
            for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
                ip = info[4][0]
                if ip not in ips and not ip.startswith(("127.", "169.254.")):
                    ips.append(ip)
        except OSError:
            pass
        return ips or ["127.0.0.1"]

    ips = lan_ips()
    print("sheetops 網頁版啟動")
    preflight()
    print(f"  本機：   http://localhost:{PORT}/pro")
    for i, ip in enumerate(ips):
        tag = "區網（給同事）：" if i == 0 else "其他網卡：    "
        print(f"  {tag} http://{ip}:{PORT}/pro")
    print(f"模型：{MODEL} @ {client.host}；使用紀錄：{LOG_PATH}")
    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="warning")

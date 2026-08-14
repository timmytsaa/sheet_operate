# -*- coding: utf-8 -*-
"""/pro 頁面（Univer 版，Apache 2.0 開源網格）。

伺服器權威架構：xlsx 只存在後端；本頁透過 /api/wb/* 取得 JSON 快照顯示、
回傳編輯事件、觸發模型處理與採納/還原。
"""

HTML_PRO = """<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="only light">
<title>試算表助理 Pro</title>
<link rel="stylesheet" href="https://unpkg.com/@univerjs/preset-sheets-core@0.25.1/lib/index.css">
<script src="https://unpkg.com/react@18.3.1/umd/react.production.min.js"></script>
<script src="https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js"></script>
<script src="https://unpkg.com/rxjs@7.8.1/dist/bundles/rxjs.umd.min.js"></script>
<script src="https://unpkg.com/@univerjs/presets@0.25.1/lib/umd/index.js"></script>
<script src="https://unpkg.com/@univerjs/preset-sheets-core@0.25.1/lib/umd/index.js"></script>
<script src="https://unpkg.com/@univerjs/preset-sheets-core@0.25.1/lib/umd/locales/zh-TW.js"></script>
<style>
  /* 樣式全部鎖定在自家 id/class，避免污染 Univer 內部元件 */
  :root { color-scheme: only light; }
  html, body { margin: 0; height: 100%; background: #fff; }
  body { font-family: "Segoe UI", "Microsoft JhengHei", system-ui, sans-serif;
         display: flex; flex-direction: column; color: #1a2233; }
  #topbar { height: 50px; display: flex; align-items: center; gap: 12px; padding: 0 16px;
            background: #1f2d50; color: #fff; flex: none; box-sizing: border-box; }
  #topbar b { font-size: 16px; } #topbar a { color: #aabdf0; font-size: 13px; text-decoration: none; }
  #topbar .spacer { flex: 1; }
  .btn { border: 0; border-radius: 7px; padding: 7px 16px; font-size: 14px; cursor: pointer;
         font-family: inherit; box-sizing: border-box; }
  .b-blue { background: #2456d6; color: #fff; } .b-blue:disabled { background: #90a5d8; }
  .b-grey { background: #e8ecf5; color: #1a2233; }
  .b-dark { background: #33406b; color: #fff; }
  #layout { flex: 1; display: flex; min-height: 0; height: calc(100vh - 50px); }
  #univer { flex: 1 1 auto; min-width: 0; height: 100%; position: relative; }
  #panel { width: 350px; flex: none; border-left: 1px solid #dde3ef; padding: 14px;
           overflow-y: auto; background: #f7f9fd; box-sizing: border-box; }
  #panel label { display: block; font-weight: 600; font-size: 13px; margin: 10px 0 4px; }
  #panel textarea { width: 100%; border: 1px solid #ccd3e0; border-radius: 7px; padding: 8px;
                    font-size: 14px; font-family: inherit; resize: vertical; box-sizing: border-box; }
  #instruction { min-height: 74px; } #context { min-height: 44px; }
  #status { font-size: 13px; color: #5a6580; margin-top: 10px; min-height: 18px; }
  #diffbox { display: none; margin-top: 10px; }
  #diff { background: #fff; border: 1px solid #e3e8f2; border-radius: 7px; padding: 10px;
          font-size: 12.5px; white-space: pre-wrap; word-break: break-all;
          max-height: 250px; overflow: auto; }
  details { margin-top: 8px; } summary { cursor: pointer; color: #5a6580; font-size: 12.5px; }
  #code { background: #fff; border: 1px solid #e3e8f2; border-radius: 7px; padding: 10px;
          font-size: 12px; white-space: pre-wrap; max-height: 200px; overflow: auto; }
  .hint { font-size: 12px; color: #9aa3b8; margin-top: 12px; }
</style></head><body>
<div id="topbar">
  <b>📊 試算表助理 Pro</b>
  <button class="btn b-dark" onclick="document.getElementById('file').click()">開啟檔案</button>
  <button class="btn b-dark" onclick="loadSample()">載入範例</button>
  <input type="file" id="file" accept=".xlsx" style="display:none">
  <span class="spacer"></span>
  <button class="btn b-grey" onclick="download()">下載</button>
  <a href="/">← 簡易版</a>
</div>
<div id="layout">
  <div id="univer"></div>
  <div id="panel">
    <label>操作指令</label>
    <textarea id="instruction" placeholder="例：刪除金額低於 7500 的列，按金額由大到小排序，底部加總計列"></textarea>
    <label>補充說明（選填）</label>
    <textarea id="context" placeholder="例：本公司折扣一律 85 折，小數無條件捨去"></textarea>
    <div style="margin-top:12px">
      <button class="btn b-blue" id="go" onclick="run()">執行</button>
    </div>
    <div id="status">先開啟檔案或載入範例。</div>
    <div id="diffbox">
      <label>變更預覽（黃底＝變更的儲存格）</label>
      <div id="diff"></div>
      <details><summary>檢視產生的程式碼</summary><pre id="code"></pre></details>
      <div style="margin-top:10px">
        <button class="btn b-blue" onclick="accept()">✅ 採納</button>
        <button class="btn b-grey" onclick="reject()">✖ 還原</button>
      </div>
    </div>
    <div class="hint">採納後可繼續下指令，或直接在表格內修改（手動修正會同步回檔案並記錄）。</div>
  </div>
</div>
<script>
const $ = id => document.getElementById(id);
let inst = null, sid = null, pendingOpen = false, loading = false;
let sheetIdToName = {};

// Univer 會跟隨系統深色偏好；本工具固定淺色，否則試算表變黑底看不見。
// 渲染分多階段完成，故重複套用並持續監看 dark class。
function forceLight(api) {
  const kill = () => {
    try { if (api && api.toggleDarkMode) api.toggleDarkMode(false); } catch (e) {}
    document.documentElement.classList.remove("dark");
    document.body.classList.remove("dark");
    document.querySelectorAll(".dark").forEach(el => el.classList.remove("dark"));
  };
  kill();
  [100, 300, 800, 1500].forEach(ms => setTimeout(kill, ms));
  if (!window._darkObserver) {
    window._darkObserver = new MutationObserver(muts => {
      for (const m of muts) {
        if (m.target.classList && m.target.classList.contains("dark")) {
          m.target.classList.remove("dark");
        }
      }
    });
    window._darkObserver.observe(document.documentElement,
      { attributes: true, attributeFilter: ["class"], subtree: true });
  }
}

function mount(snapshot, coordsBySheet) {
  loading = true;
  try {
    if (inst) { try { inst.univer.dispose(); } catch (e) {} }
    const P = window.UniverPresets;
    const C = window.UniverPresetSheetsCore;
    const LT = window.UniverCore.LocaleType;   // 0.25.x：LocaleType 在 UniverCore
    const theme = (window.UniverThemes && window.UniverThemes.defaultTheme) ||
                  (window.UniverDesign && window.UniverDesign.defaultTheme);
    const loc = LT.ZH_TW || LT.ZH_CN;
    const locData = window.UniverPresetSheetsCoreZhTW || window.UniverPresetSheetsCoreZhCN || {};
    inst = P.createUniver({
      locale: loc,
      locales: { [loc]: locData },
      theme: theme,
      darkMode: false,
      presets: [C.UniverSheetsCorePreset({ container: "univer" })],
    });
    const api = inst.univerAPI;
    // Univer 會跟隨系統的深色偏好；本工具固定淺色，否則試算表變黑底看不清楚
    if (api.toggleDarkMode) { try { api.toggleDarkMode(false); } catch (e) {} }
    (api.createWorkbook || api.createUniverSheet).call(api, snapshot);
    sheetIdToName = {};
    for (const [sid_, sh] of Object.entries(snapshot.sheets || {})) sheetIdToName[sid_] = sh.name;
    bindEdits(api);
    forceLight(api);
    setTimeout(() => window.dispatchEvent(new Event("resize")), 150);
    if (coordsBySheet) setTimeout(() => highlight(coordsBySheet), 300);
  } catch (e) {
    $("status").textContent = "❌ 網格初始化失敗：" + e;
    console.error(e);
  }
  setTimeout(() => { loading = false; }, 500);
}

function highlight(coordsBySheet) {
  try {
    const wbk = inst.univerAPI.getActiveWorkbook();
    for (const [name, coords] of Object.entries(coordsBySheet)) {
      const sh = wbk.getSheetByName(name);
      if (!sh) continue;
      coords.forEach(([r, c]) => { try { sh.getRange(r, c).setBackgroundColor("#FFF3B0"); } catch (e) {} });
    }
  } catch (e) { console.error("highlight", e); }
}

function bindEdits(api) {
  if (!api.onCommandExecuted) return;
  api.onCommandExecuted(cmd => {
    if (loading || !sid) return;
    if (!cmd || !cmd.id || cmd.id.indexOf("set-range-values") < 0) return;
    try {
      const p = cmd.params || {};
      const sheetName = sheetIdToName[p.subUnitId] || p.subUnitId;
      const cv = p.cellValue || {};
      for (const [r, colObj] of Object.entries(cv)) {
        for (const [c, entry] of Object.entries(colObj || {})) {
          const v = entry && entry.v !== undefined ? entry.v : null;
          fetch(`/api/wb/${sid}/edit`, { method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ sheet: sheetName, row: +r, col: +c, value: v,
                                   target: pendingOpen ? "pending" : "current" }) });
        }
      }
    } catch (e) { console.error("edit-log", e); }
  });
}

async function openSession(resp) {
  const j = await resp.json();
  if (!j.ok) { $("status").textContent = "❌ " + j.error; return; }
  sid = j.sid; pendingOpen = false; $("diffbox").style.display = "none";
  mount(j.snapshot);
  $("status").textContent = "已載入：" + j.name + "——可以下指令或直接編輯。";
}

$("file").addEventListener("change", async e => {
  const f = e.target.files[0];
  if (!f) return;
  const fd = new FormData(); fd.append("file", f);
  $("status").textContent = "上傳中…";
  openSession(await fetch("/api/wb/open", { method: "POST", body: fd }));
});

async function loadSample() {
  $("status").textContent = "載入範例中…";
  openSession(await fetch("/api/wb/sample"));
}

async function run() {
  if (!sid) { $("status").textContent = "請先開啟檔案或載入範例"; return; }
  if (pendingOpen) { $("status").textContent = "請先採納或還原上一次的變更"; return; }
  const instn = $("instruction").value.trim();
  if (!instn) { $("status").textContent = "請輸入操作指令"; return; }
  $("go").disabled = true;
  $("status").textContent = "⏳ 模型處理中（約 20 秒～2 分鐘）…";
  try {
    const r = await fetch(`/api/wb/${sid}/process`, { method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ instruction: instn, context: $("context").value }) });
    const j = await r.json();
    if (!j.ok) { $("status").textContent = "❌ " + j.error; return; }
    pendingOpen = true;
    mount(j.snapshot, j.diff_coords);
    $("diff").textContent = j.diff_text;
    $("code").textContent = j.code;
    $("diffbox").style.display = "block";
    $("status").textContent = `完成（${j.seconds} 秒）——檢查黃底變更後採納或還原。`;
  } catch (e) { $("status").textContent = "❌ 連線錯誤：" + e; }
  finally { $("go").disabled = false; }
}

async function accept() {
  if (!sid || !pendingOpen) return;
  const j = await (await fetch(`/api/wb/${sid}/accept`, { method: "POST" })).json();
  pendingOpen = false; $("diffbox").style.display = "none";
  if (j.ok) mount(j.snapshot);
  $("status").textContent = "已採納，可繼續下一個指令或手動微調。";
}

async function reject() {
  if (!sid || !pendingOpen) return;
  const j = await (await fetch(`/api/wb/${sid}/reject`, { method: "POST" })).json();
  pendingOpen = false; $("diffbox").style.display = "none";
  if (j.ok) mount(j.snapshot);
  $("status").textContent = "已還原到執行前的狀態。";
}

function download() {
  if (!sid) return;
  if (pendingOpen) { $("status").textContent = "有未決定的變更——請先採納或還原再下載"; return; }
  window.location = `/api/wb/${sid}/download`;
}
</script></body></html>"""

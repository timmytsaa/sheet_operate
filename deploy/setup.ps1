<#
.SYNOPSIS
  sheetops 一鍵部署：建虛擬環境、裝套件、備妥 .env，然後啟動網頁服務（區網 IP 可連）。

.EXAMPLE
  .\deploy\setup.ps1                     # 首次部署或日常啟動（自動判斷）
  .\deploy\setup.ps1 -Port 8080          # 換埠號
  .\deploy\setup.ps1 -Reinstall          # 重建虛擬環境
  .\deploy\setup.ps1 -Firewall           # 順便開防火牆（需以系統管理員身分執行）
  .\deploy\setup.ps1 -NoStart            # 只安裝不啟動
#>
param(
    [int]    $Port      = 0,
    [string] $Model     = "",
    [switch] $Reinstall,
    [switch] $Firewall,
    [switch] $NoStart
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$VenvDir = Join-Path $Root ".venv"
$VenvPy  = Join-Path $VenvDir "Scripts\python.exe"
$EnvFile = Join-Path $Root ".env"

function Say($msg, $color = "White") { Write-Host $msg -ForegroundColor $color }
function Step($n, $msg) { Write-Host "`n[$n] $msg" -ForegroundColor Cyan }

Say "===== sheetops 部署 =====" Cyan
Say "專案位置：$Root"

# ---------- 1. Python ----------
Step 1 "檢查 Python"
$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) { $py = Get-Command py -ErrorAction SilentlyContinue }
if (-not $py) {
    Say "找不到 Python。請先安裝 Python 3.10 以上：https://www.python.org/downloads/" Red
    Say "安裝時記得勾選 Add python.exe to PATH。" Red
    exit 1
}
$pyVer = & $py.Source -c "import sys; print('.'.join(map(str, sys.version_info[:2])))"
Say "  使用 $($py.Source)（$pyVer）"

# ---------- 2. 虛擬環境 ----------
Step 2 "虛擬環境"
if ($Reinstall -and (Test-Path $VenvDir)) {
    Say "  -Reinstall：移除既有的 .venv"
    Remove-Item -Recurse -Force $VenvDir
}
if (-not (Test-Path $VenvPy)) {
    Say "  建立 .venv（第一次會花一兩分鐘）…"
    & $py.Source -m venv $VenvDir
    if (-not (Test-Path $VenvPy)) { Say "建立虛擬環境失敗" Red; exit 1 }
} else {
    Say "  已存在，沿用（要重建請加 -Reinstall）"
}

# ---------- 3. 套件 ----------
Step 3 "安裝套件"
$req = Join-Path $PSScriptRoot "requirements-serve.txt"
$pipLog = & $VenvPy -m pip install --upgrade pip 2>&1
$pipLog = & $VenvPy -m pip install -r $req 2>&1
if ($LASTEXITCODE -ne 0) {
    Say "套件安裝失敗：" Red
    $pipLog | Select-Object -Last 15 | ForEach-Object { Say "    $_" Red }
    exit 1
}
Say "  完成：fastapi / uvicorn / openpyxl / requests"

# ---------- 4. .env ----------
Step 4 "設定檔 .env"
if (-not (Test-Path $EnvFile)) {
    Copy-Item (Join-Path $PSScriptRoot "env.example") $EnvFile
    Say "  已從範本建立 .env"
} else {
    Say "  已存在，保留原設定"
}
if ($Port -gt 0) {
    (Get-Content $EnvFile -Encoding UTF8) -replace '^SHEETOPS_PORT=.*', "SHEETOPS_PORT=$Port" |
        Set-Content $EnvFile -Encoding UTF8
    Say "  埠號設為 $Port"
}
if ($Model -ne "") {
    (Get-Content $EnvFile -Encoding UTF8) -replace '^SHEETOPS_MODEL=.*', "SHEETOPS_MODEL=$Model" |
        Set-Content $EnvFile -Encoding UTF8
    Say "  模型設為 $Model"
}
# 讀回設定值（給後面顯示網址與檢查模型用）
$cfg = @{}
foreach ($line in Get-Content $EnvFile -Encoding UTF8) {
    if ($line -match '^\s*([A-Z_]+)\s*=\s*(.*)$') { $cfg[$matches[1]] = $matches[2].Trim() }
}
$usePort  = if ($cfg.SHEETOPS_PORT)  { $cfg.SHEETOPS_PORT }  else { "8033" }
$useModel = if ($cfg.SHEETOPS_MODEL) { $cfg.SHEETOPS_MODEL } else { "sheetops" }

# ---------- 5. Ollama 與模型 ----------
Step 5 "檢查 Ollama 與模型"
$ollama = Get-Command ollama -ErrorAction SilentlyContinue
if (-not $ollama) {
    Say "  ⚠ 找不到 ollama 指令。請先安裝 Ollama：https://ollama.com/download" Yellow
} else {
    $models = (& ollama list) 2>$null
    if ($models -match [regex]::Escape($useModel)) {
        Say "  模型「$useModel」已就緒"
    } else {
        Say "  ⚠ Ollama 裡沒有「$useModel」模型。" Yellow
        Say "    請把 sheetops-q8_0.gguf 放到 deploy\ 後執行：" Yellow
        Say "      ollama create $useModel -f deploy\Modelfile" Yellow
    }
}

# ---------- 6. 防火牆（選用，需系統管理員） ----------
Step 6 "區網存取（防火牆）"
$ruleName = "sheetops $usePort"
$hasRule = $false
try { $hasRule = [bool](Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue) } catch {}
if ($hasRule) {
    Say "  防火牆規則已存在（$ruleName）"
} elseif ($Firewall) {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        netsh advfirewall firewall add rule name="$ruleName" dir=in action=allow `
              protocol=TCP localport=$usePort profile=domain,private | Out-Null
        Say "  已新增防火牆規則：$ruleName" Green
    } else {
        Say "  ⚠ -Firewall 需要系統管理員權限。請以系統管理員身分重開 PowerShell 再執行。" Yellow
    }
} else {
    Say "  尚未開放。同事若連不上，請以系統管理員身分執行：" Yellow
    Say "    netsh advfirewall firewall add rule name=`"$ruleName`" dir=in action=allow protocol=TCP localport=$usePort profile=domain,private" Yellow
}

# ---------- 7. 啟動 ----------
if ($NoStart) {
    Step 7 "完成（-NoStart，未啟動服務）"
    Say "  日後啟動：.\deploy\start.bat 或 .\deploy\setup.ps1"
    exit 0
}

Step 7 "啟動服務"

# 埠號被佔用時 uvicorn 會秒退，視窗跟著關掉——先擋下來講清楚
$busy = Get-NetTCPConnection -LocalPort $usePort -State Listen -ErrorAction SilentlyContinue
if ($busy) {
    $pid0 = ($busy | Select-Object -First 1).OwningProcess
    $pname = (Get-Process -Id $pid0 -ErrorAction SilentlyContinue).ProcessName
    Say "  ⚠ 連接埠 $usePort 已被佔用（PID $pid0 $pname）——服務可能已經在跑。" Yellow
    Say "    直接開 http://localhost:$usePort/pro 看看；" Yellow
    Say "    若要重開，先停掉它：  taskkill /PID $pid0 /F" Yellow
    Say "    或改用其他埠號：      .\deploy\setup.ps1 -Port 8080" Yellow
    exit 1
}

Say "  視窗保持開著；要停止按 Ctrl+C" Yellow
Say ""
& $VenvPy (Join-Path $Root "scripts\serve_web.py")
$code = $LASTEXITCODE
if ($code -ne 0) { Say "`n服務結束（離開碼 $code）" Red }
exit $code

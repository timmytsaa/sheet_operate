"""在子行程中執行模型產生的 Python 程式碼。

契約（與 prompts.SYSTEM_PROMPT 一致）：
- 程式碼可讀取變數 INPUT_PATH / OUTPUT_PATH
- 程式碼必須把處理後的工作簿存到 OUTPUT_PATH

v1 沙盒等級：獨立子行程 + 逾時 + 獨立工作目錄。teacher / 自產程式碼風險低；
之後在 Colab 上跑 RL rollout 時，建議再加一層（nsjail 或容器）。
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path

_DRIVER_PRELUDE = """# -*- coding: utf-8 -*-
import os
INPUT_PATH = os.environ["SHEETOPS_INPUT"]
OUTPUT_PATH = os.environ["SHEETOPS_OUTPUT"]
# === 以下為待執行程式碼 ===
"""

_CODE_BLOCK_RE = re.compile(r"```(?:python|py)?\s*\n(.*?)```", re.DOTALL)


def extract_code(text: str) -> str | None:
    """從模型回覆中取出最後一個 python 程式碼區塊；沒有區塊時，若整段看似程式碼則直接回傳。"""
    blocks = _CODE_BLOCK_RE.findall(text)
    if blocks:
        return blocks[-1].strip()
    if "wb.save" in text or "openpyxl" in text:
        return text.strip()
    return None


@dataclass
class ExecResult:
    ok: bool
    output_exists: bool
    stdout: str = ""
    stderr: str = ""
    seconds: float = 0.0
    reason: str = ""

    def feedback(self, limit: int = 1200) -> str:
        parts = []
        if self.reason:
            parts.append(f"[執行狀態] {self.reason}")
        if self.stdout.strip():
            parts.append("[stdout]\n" + self.stdout.strip()[-limit:])
        if self.stderr.strip():
            parts.append("[stderr]\n" + self.stderr.strip()[-limit:])
        return "\n".join(parts) if parts else "[執行狀態] 完成，無輸出"


def run_code(code: str, input_path: str | Path, output_path: str | Path,
             timeout: int = 30, safety: bool = True) -> ExecResult:
    if safety:
        from .safety import check_code_safety
        ok, reason = check_code_safety(code)
        if not ok:
            return ExecResult(False, False, "", reason, 0.0,
                              f"程式碼未通過安全檢查：{reason}")

    input_path = str(Path(input_path).resolve())
    output_path = str(Path(output_path).resolve())

    with tempfile.TemporaryDirectory(prefix="sheetops_run_") as work:
        driver = Path(work) / "driver.py"
        driver.write_text(_DRIVER_PRELUDE + code + "\n", encoding="utf-8")

        env = os.environ.copy()
        env["SHEETOPS_INPUT"] = input_path
        env["SHEETOPS_OUTPUT"] = output_path
        env["PYTHONIOENCODING"] = "utf-8"

        t0 = time.monotonic()
        try:
            proc = subprocess.run(
                [sys.executable, "-X", "utf8", str(driver)],
                cwd=work, env=env, capture_output=True,
                encoding="utf-8", errors="replace", timeout=timeout,
            )
            seconds = time.monotonic() - t0
            out_exists = Path(output_path).exists()
            if proc.returncode != 0:
                return ExecResult(False, out_exists, proc.stdout, proc.stderr, seconds,
                                  f"程式碼執行失敗 (exit {proc.returncode})")
            if not out_exists:
                return ExecResult(False, False, proc.stdout, proc.stderr, seconds,
                                  "程式碼執行完成，但沒有存檔到 OUTPUT_PATH")
            return ExecResult(True, True, proc.stdout, proc.stderr, seconds, "執行成功")
        except subprocess.TimeoutExpired as e:
            seconds = time.monotonic() - t0
            return ExecResult(False, Path(output_path).exists(),
                              (e.stdout or b"").decode("utf-8", "replace") if isinstance(e.stdout, bytes) else (e.stdout or ""),
                              (e.stderr or b"").decode("utf-8", "replace") if isinstance(e.stderr, bytes) else (e.stderr or ""),
                              seconds, f"逾時（超過 {timeout} 秒）")

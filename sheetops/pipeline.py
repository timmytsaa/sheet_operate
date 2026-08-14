"""共用處理管線：一份 xlsx ＋ 一句指令 → 沙盒執行 → 差異報告。

CLI 與 Web 服務共用。generate_fn 由呼叫端注入（Ollama / llama.cpp / 任意 LLM），
本模組不關心模型來自哪裡。
"""
from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Callable

from .diff import diff_workbooks, render_diff
from .encoder import encode_workbook
from .executor import extract_code, run_code
from .prompts import SYSTEM_PROMPT, build_user_prompt


def process_workbook(src: str | Path, instruction: str, context: str = "",
                     generate_fn: Callable[[list[dict]], str] = None,
                     retries: int = 1, exec_timeout: int = 60) -> dict:
    """回傳：
    {ok, code, attempts, result_path(暫存), diff(dict), diff_text, error}
    ok=False 時 result_path 為 None，error 描述原因。
    呼叫端負責在用完後清理 result_path 所在的暫存資料夾。
    """
    src = Path(src).resolve()
    sheet_text = encode_workbook(src)
    messages = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": build_user_prompt(instruction, sheet_text, context)}]

    work = Path(tempfile.mkdtemp(prefix="sheetops_pipe_"))
    result_path = work / "result.xlsx"
    code = None

    for attempt in range(retries + 1):
        reply = generate_fn(messages)
        code = extract_code(reply)
        if not code:
            return {"ok": False, "code": None, "attempts": attempt + 1,
                    "result_path": None, "error": "模型沒有輸出程式碼區塊"}

        exec_result = run_code(code, src, result_path, timeout=exec_timeout)
        if exec_result.ok:
            d = diff_workbooks(src, result_path)
            return {"ok": True, "code": code, "attempts": attempt + 1,
                    "result_path": result_path, "diff": d,
                    "diff_text": render_diff(d), "error": ""}

        if attempt < retries:
            messages.append({"role": "assistant", "content": reply})
            messages.append({"role": "user",
                             "content": "程式碼執行失敗，請修正後重新輸出完整程式碼。\n"
                                        + exec_result.feedback()})

    return {"ok": False, "code": code, "attempts": retries + 1,
            "result_path": None, "error": exec_result.feedback()}

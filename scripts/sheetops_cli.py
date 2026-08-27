"""sheetops 本地部署 CLI：吃 xlsx＋繁中指令 → 本地模型產碼 → 沙盒執行 → 預覽變更 → 確認後寫入。

安全契約：
- 原始檔案永不直接覆寫；預設輸出「<原名>_已修改.xlsx」
- --in-place 時先備份「<原名>.xlsx.bak」再取代
- 模型程式碼經 AST 白名單檢查後才在子行程沙盒執行
- 變更摘要先給你看，按下確認才落地（--yes 可跳過，供批次使用）

用法：
  python scripts/sheetops_cli.py 報表.xlsx "把金額低於 5000 的列刪掉，再照金額由大到小排序"
  python scripts/sheetops_cli.py 報表.xlsx "新增含稅價欄" --context "含稅價 = 金額 × 1.05，四捨五入取整"
  python scripts/sheetops_cli.py 報表.xlsx "..." --in-place --yes
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
# Colab/IPython 的 sys.stdout 是 ipykernel 的 OutStream，沒有 reconfigure——
# 這些腳本會被 notebook import，不防守就會在 import 當下就炸掉。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.diff import diff_workbooks, render_diff
from sheetops.encoder import encode_workbook
from sheetops.executor import extract_code, run_code
from sheetops.ollama_client import OllamaClient
from sheetops.prompts import SYSTEM_PROMPT, build_user_prompt

LOG_PATH = Path(__file__).resolve().parents[1] / "logs" / "usage_log.jsonl"


def log_event(record: dict) -> None:
    """與網頁版同格式的使用紀錄（v3 任務規格與 DPO 的原料）。"""
    import json
    import time
    record["ts"] = time.strftime("%Y-%m-%d %H:%M:%S")
    record["source"] = "cli"
    LOG_PATH.parent.mkdir(exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def main() -> int:
    ap = argparse.ArgumentParser(description="sheetops：口語指令操作 Excel")
    ap.add_argument("file", help="目標 .xlsx 檔案")
    ap.add_argument("instruction", help="繁體中文操作指令")
    ap.add_argument("--context", default="", help="補充說明（公司自訂規則）")
    ap.add_argument("--model", default="sheetops", help="本地 Ollama 模型名")
    ap.add_argument("--host", default="http://localhost:11434")
    ap.add_argument("--gguf", default=None,
                    help="直接載入 GGUF 推論、不經 Ollama（需 pip install llama-cpp-python）")
    ap.add_argument("--out", default=None, help="輸出路徑（預設 <原名>_已修改.xlsx）")
    ap.add_argument("--in-place", action="store_true", help="直接取代原檔（自動留 .bak 備份）")
    ap.add_argument("--yes", action="store_true", help="跳過確認（批次模式）")
    ap.add_argument("--retries", type=int, default=1, help="執行失敗時帶回饋重試次數")
    ap.add_argument("--show-code", action="store_true", help="顯示模型產生的程式碼")
    ap.add_argument("--code-file", default=None, help=argparse.SUPPRESS)  # 測試用：跳過模型
    args = ap.parse_args()

    src = Path(args.file).resolve()
    if not src.exists():
        print(f"找不到檔案：{src}")
        return 1

    sheet_text = encode_workbook(src)
    user_prompt = build_user_prompt(args.instruction, sheet_text, args.context)
    messages = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt}]

    # 推論後端：預設 Ollama；--gguf 時行程內直連 llama.cpp
    if args.gguf:
        try:
            from llama_cpp import Llama
        except ImportError:
            print("--gguf 模式需要：pip install llama-cpp-python")
            return 1
        llm = Llama(model_path=args.gguf, n_ctx=8192, n_gpu_layers=-1, verbose=False)

        def generate(msgs):
            out = llm.create_chat_completion(messages=msgs, temperature=0.0, max_tokens=1024)
            return out["choices"][0]["message"]["content"]
    else:
        client = OllamaClient(host=args.host, model=args.model)

        def generate(msgs):
            return client.chat(msgs, temperature=0.0)

    work = Path(tempfile.mkdtemp(prefix="sheetops_cli_"))
    result_path = work / "result.xlsx"

    exec_result = None
    for attempt in range(args.retries + 1):
        if args.code_file:
            code = Path(args.code_file).read_text(encoding="utf-8")
        else:
            label = "產生操作程式碼" if attempt == 0 else f"重試（第 {attempt} 次）"
            print(f"[{Path(args.gguf).name if args.gguf else args.model}] {label}…")
            reply = generate(messages)
            code = extract_code(reply)
            if not code:
                print("模型沒有輸出程式碼區塊，中止。")
                return 1
            messages.append({"role": "assistant", "content": reply})
        if args.show_code:
            print("----- 程式碼 -----")
            print(code)
            print("------------------")

        exec_result = run_code(code, src, result_path, timeout=60)
        if exec_result.ok:
            break
        print(f"執行失敗：{exec_result.reason}")
        if args.code_file or attempt >= args.retries:
            print(exec_result.feedback())
            log_event({"event": "process", "ok": False, "file": src.name,
                       "instruction": args.instruction, "context": args.context,
                       "model": Path(args.gguf).name if args.gguf else args.model,
                       "code": code, "error": exec_result.reason})
            return 1
        messages.append({"role": "user",
                         "content": "程式碼執行失敗，請修正後重新輸出完整程式碼。\n"
                                    + exec_result.feedback()})

    d = diff_workbooks(src, result_path)
    print("\n===== 變更預覽 =====")
    print(render_diff(d))
    print("====================\n")

    base_log = {"event": "process", "ok": True, "file": src.name,
                "instruction": args.instruction, "context": args.context,
                "model": Path(args.gguf).name if args.gguf else args.model,
                "code": code,
                "changed_cells": sum(v["changed"] for v in d["sheets"].values()),
                "sheets_added": d["sheets_added"]}

    if not args.yes:
        ans = input("套用變更？[y/N] ").strip().lower()
        if ans not in ("y", "yes"):
            log_event({**base_log, "decision": "rejected"})
            print("已取消，原始檔案未變動。")
            return 0
    log_event({**base_log, "decision": "accepted"})

    if args.in_place:
        backup = src.with_suffix(src.suffix + ".bak")
        shutil.copy2(src, backup)
        shutil.copy2(result_path, src)
        print(f"已套用到原檔：{src}（備份：{backup}）")
    else:
        out = Path(args.out) if args.out else src.with_name(src.stem + "_已修改" + src.suffix)
        shutil.copy2(result_path, out)
        print(f"已輸出：{out}（原始檔案未變動）")
    shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())

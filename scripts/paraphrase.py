"""指令改寫：用 Ollama 大模型把模板指令改寫得更自然、多樣（保留數值與名稱）。

改寫後覆寫 task.json 的 instruction；原句保留在 instruction_seed。
安全檢查：改寫後必須保留原句中的所有數字與「」內的名稱，否則維持原句。

用法：
  python scripts/paraphrase.py --tasks data/tasks/train [--limit 50]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.ollama_client import OllamaClient

PROMPT = """請把下面這句試算表操作指令改寫成另一種自然的繁體中文說法。
要求：
1. 語義完全相同，不可增減任何操作要求
2. 所有數字、工作表名稱、欄位名稱（「」內的文字）必須原樣保留
3. 只輸出改寫後的指令，不要任何解釋

原指令：__INSTRUCTION__"""


def _tokens(text: str) -> set[str]:
    nums = set(re.findall(r"\d+(?:\.\d+)?", text))
    names = set(re.findall(r"「([^」]+)」", text))
    return nums | names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="data/tasks/train")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--model", default=None)
    args = ap.parse_args()

    client = OllamaClient(model=args.model)
    ok, msg = client.available()
    if not ok:
        print(msg)
        sys.exit(1)

    task_files = sorted(Path(args.tasks).glob("*/task.json"))
    if args.limit:
        task_files = task_files[:args.limit]

    n_ok = n_keep = 0
    n_done = 0
    for tf in task_files:
        spec = json.loads(tf.read_text(encoding="utf-8"))
        if spec.get("instruction_seed") and spec["instruction"] != spec["instruction_seed"]:
            n_done += 1
            continue  # 已改寫過，跳過（支援增量續跑）
        seed = spec.get("instruction_seed") or spec["instruction"]
        try:
            new = client.chat(
                [{"role": "user", "content": PROMPT.replace("__INSTRUCTION__", seed)}],
                temperature=0.9).strip().strip("「」\"'")
        except RuntimeError as e:
            print(f"  API 錯誤：{e}")
            break

        if new and _tokens(seed) <= _tokens(new) and 5 < len(new) < len(seed) * 3:
            spec["instruction"] = new
            spec["instruction_seed"] = seed
            tf.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
            n_ok += 1
        else:
            n_keep += 1

    print(f"改寫完成 {n_ok} 筆；維持原句 {n_keep} 筆；先前已改寫跳過 {n_done} 筆")


if __name__ == "__main__":
    main()

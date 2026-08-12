"""Teacher 蒸餾：用 Ollama 大模型解任務，經 Gym 驗證後保留通過的軌跡（rejection sampling）。

需要環境變數 OLLAMA_API_KEY（見 sheetops/ollama_client.py）。
支援中斷續跑：已在輸出檔中的任務會自動略過。

用法：
  python scripts/teacher_solve.py --tasks data/tasks/train --out data/sft/teacher_sft.jsonl --k 4
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.encoder import encode_workbook
from sheetops.env import solve_once
from sheetops.executor import extract_code
from sheetops.ollama_client import OllamaClient
from sheetops.prompts import SYSTEM_PROMPT, build_user_prompt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="data/tasks/train")
    ap.add_argument("--out", default="data/sft/teacher_sft.jsonl")
    ap.add_argument("--k", type=int, default=4, help="每題最多採樣次數")
    ap.add_argument("--temperature", type=float, default=0.8)
    ap.add_argument("--model", default=None, help="覆寫 OLLAMA_MODEL")
    ap.add_argument("--limit", type=int, default=0, help="只處理前 N 題（0 = 全部）")
    ap.add_argument("--shard", default=None,
                    help="k/n：只處理第 k 份任務（0-based），供多 teacher 平行分工")
    args = ap.parse_args()

    client = OllamaClient(model=args.model)
    ok, msg = client.available()
    if not ok:
        print(msg)
        sys.exit(1)
    print(f"teacher 模型：{client.model} @ {client.host}")

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    done_ids = set()
    if out_path.exists():
        with out_path.open(encoding="utf-8") as f:
            for line in f:
                try:
                    done_ids.add(json.loads(line)["id"])
                except (json.JSONDecodeError, KeyError):
                    pass

    task_dirs = sorted(p.parent for p in Path(args.tasks).glob("*/task.json"))
    if args.shard:
        k, n = map(int, args.shard.split("/"))
        task_dirs = [d for i, d in enumerate(task_dirs) if i % n == k]
    if args.limit and args.limit < len(task_dirs):
        # 等距抽樣，讓小批試跑也能涵蓋各任務家族
        stride = len(task_dirs) // args.limit
        task_dirs = task_dirs[::stride][:args.limit]

    n_solved = n_failed = 0
    with out_path.open("a", encoding="utf-8") as f:
        for task_dir in task_dirs:
            spec = json.loads((task_dir / "task.json").read_text(encoding="utf-8"))
            if spec["id"] in done_ids:
                continue

            sheet_text = encode_workbook(task_dir / "start.xlsx")
            user_prompt = build_user_prompt(spec["instruction"], sheet_text,
                                            spec.get("context", ""))
            messages = [{"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt}]

            solved = False
            for attempt in range(args.k):
                temp = 0.2 if attempt == 0 else args.temperature
                try:
                    reply = client.chat(messages, temperature=temp)
                except RuntimeError as e:
                    print(f"  API 錯誤（{spec['id']}）：{e}")
                    break
                code = extract_code(reply)
                if not code:
                    continue
                report = solve_once(task_dir, code)
                if report["full_match"]:
                    record = {
                        "id": spec["id"], "family": spec["family"],
                        "source": f"teacher:{client.model}", "attempt": attempt,
                        "messages": messages + [
                            {"role": "assistant", "content": "```python\n" + code + "\n```"}],
                    }
                    f.write(json.dumps(record, ensure_ascii=False) + "\n")
                    f.flush()
                    solved = True
                    break

            if solved:
                n_solved += 1
                print(f"  [V] {spec['id']}")
            else:
                n_failed += 1
                print(f"  [X] {spec['id']}（{args.k} 次採樣皆未通過）")

    print(f"\n完成：通過 {n_solved}、未通過 {n_failed} → {out_path}")
    print("未通過的題目可提高 --k 或換更強的 teacher 模型重跑（會自動略過已完成者）。")


if __name__ == "__main__":
    main()

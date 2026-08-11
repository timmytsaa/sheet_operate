"""把任務目錄轉成 SFT 訓練資料（chat 格式 jsonl）。

預設使用每個任務的參考解法（種子資料）；teacher 蒸餾產生的軌跡
由 teacher_solve.py 直接輸出同格式，之後合併即可。

用法：
  python scripts/build_sft_dataset.py --tasks data/tasks/train --out data/sft/seed_sft.jsonl [--verify]
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
from sheetops.prompts import SYSTEM_PROMPT, build_user_prompt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="data/tasks/train")
    ap.add_argument("--out", default="data/sft/seed_sft.jsonl")
    ap.add_argument("--verify", action="store_true", help="輸出前重新驗證參考解法（較慢）")
    args = ap.parse_args()

    task_dirs = sorted(p.parent for p in Path(args.tasks).glob("*/task.json"))
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    n_ok, n_skip = 0, 0
    with out_path.open("w", encoding="utf-8") as f:
        for task_dir in task_dirs:
            spec = json.loads((task_dir / "task.json").read_text(encoding="utf-8"))
            if args.verify:
                report = solve_once(task_dir, spec["ref_solution"])
                if not report["full_match"]:
                    print(f"跳過（參考解法未通過驗證）: {spec['id']}")
                    n_skip += 1
                    continue
            sheet_text = encode_workbook(task_dir / "start.xlsx")
            record = {
                "id": spec["id"],
                "family": spec["family"],
                "source": "seed",
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": build_user_prompt(
                        spec["instruction"], sheet_text, spec.get("context", ""))},
                    {"role": "assistant",
                     "content": "```python\n" + spec["ref_solution"].strip() + "\n```"},
                ],
            }
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
            n_ok += 1

    print(f"完成：{n_ok} 筆 → {out_path}" + (f"（跳過 {n_skip} 筆）" if n_skip else ""))


if __name__ == "__main__":
    main()

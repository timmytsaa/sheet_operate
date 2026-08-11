"""端到端自我測試：

每個任務家族產生數個任務 → 將「參考解法」丟進 Gym 執行 → 驗證 reward 必須為滿分。
這條路通了，代表 編碼器 / 執行器 / 驗證器 / 任務產生器 彼此一致，
之後 teacher 蒸餾與 RL 都建立在同一條驗證路徑上。

用法： python scripts/selftest.py [--n 3] [--keep]
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from random import Random

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.encoder import encode_workbook
from sheetops.env import solve_once
from sheetops.taskgen import FAMILIES, generate_task


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=3, help="每個家族的測試任務數")
    ap.add_argument("--keep", action="store_true", help="保留產生的測試任務目錄")
    args = ap.parse_args()

    out_root = ROOT / "data" / "selftest"
    if out_root.exists():
        shutil.rmtree(out_root)

    results = []
    first_dir = None
    for fam in FAMILIES:
        for i in range(args.n):
            task_id = f"{fam}-{i:02d}"
            rng = Random(f"selftest:{fam}:{i}")
            task_dir = generate_task(fam, rng, task_id, out_root)
            if first_dir is None:
                first_dir = task_dir

            import json
            spec = json.loads((task_dir / "task.json").read_text(encoding="utf-8"))
            report = solve_once(task_dir, spec["ref_solution"])
            results.append((task_id, spec["meta"].get("variant", "-"), report))

    # ---- 報表 ----
    print(f"{'任務':<24}{'變體':<12}{'執行':<6}{'分數':<8}全對")
    print("-" * 60)
    n_pass = 0
    for task_id, variant, rep in results:
        ok = "OK" if rep["exec_ok"] else "FAIL"
        full = "V" if rep["full_match"] else "X"
        n_pass += rep["full_match"]
        print(f"{task_id:<24}{variant:<12}{ok:<6}{rep['score']:<8.3f}{full}")
        if not rep["full_match"]:
            for m in rep["mismatches"][:3]:
                print(f"    !! {m}")
            if not rep["exec_ok"]:
                print("    !! " + rep["exec_feedback"].replace("\n", "\n    !! ")[:800])

    print("-" * 60)
    print(f"通過 {n_pass}/{len(results)}")

    if first_dir is not None:
        print("\n===== 編碼器輸出範例（第一個任務的起始表） =====")
        text = encode_workbook(first_dir / "start.xlsx")
        lines = text.splitlines()
        print("\n".join(lines[:28]))
        if len(lines) > 28:
            print(f"…（共 {len(lines)} 行）")

    if not args.keep:
        shutil.rmtree(out_root, ignore_errors=True)

    sys.exit(0 if n_pass == len(results) else 1)


if __name__ == "__main__":
    main()

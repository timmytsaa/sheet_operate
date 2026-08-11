"""批次產生合成任務。

用法：
  python scripts/gen_tasks.py --out data/tasks/train --n 30 --seed 20260811
  python scripts/gen_tasks.py --out data/tasks/eval  --n 5  --seed 900001

每個任務使用獨立 rng（seed:family:index 派生），新增家族或改變 n 不會影響既有任務內容。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from random import Random

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.taskgen import FAMILIES, generate_task


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="data/tasks/train")
    ap.add_argument("--n", type=int, default=10, help="每個家族的任務數")
    ap.add_argument("--seed", type=int, default=20260811)
    ap.add_argument("--families", default="all", help="逗號分隔的家族名稱，或 all")
    ap.add_argument("--start-index", type=int, default=0,
                    help="任務編號起點（增量擴產用，避免覆蓋既有任務）")
    ap.add_argument("--schema", default=None,
                    help="強制使用指定 schema（如 hr）；預設隨機輪換訓練用 schema")
    args = ap.parse_args()

    if args.schema:
        from sheetops.taskgen import base as _base
        assert args.schema in _base.SCHEMAS, f"未知 schema: {args.schema}"
        _base.FORCE_SCHEMA = args.schema

    fams = list(FAMILIES) if args.families == "all" else args.families.split(",")
    out_root = Path(args.out)
    out_root.mkdir(parents=True, exist_ok=True)

    manifest = []
    for fam in fams:
        for i in range(args.start_index, args.start_index + args.n):
            task_id = f"{fam}-{i:04d}"
            rng = Random(f"{args.seed}:{fam}:{i}")
            task_dir = generate_task(fam, rng, task_id, out_root)
            spec = json.loads((task_dir / "task.json").read_text(encoding="utf-8"))
            manifest.append({"id": task_id, "family": fam,
                             "variant": spec["meta"].get("variant", ""),
                             "instruction": spec["instruction"]})

    # manifest 以「掃描整個輸出目錄」重建，增量生成單一家族時不會覆蓋掉其他家族的紀錄
    all_entries = []
    for tf in sorted(out_root.glob("*/task.json")):
        try:
            spec = json.loads(tf.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            continue  # 其他程序可能正在寫入，略過
        all_entries.append({"id": spec["id"], "family": spec["family"],
                            "variant": spec["meta"].get("variant", ""),
                            "instruction": spec["instruction"]})
    with (out_root / "manifest.jsonl").open("w", encoding="utf-8") as f:
        for m in all_entries:
            f.write(json.dumps(m, ensure_ascii=False) + "\n")

    print(f"完成：本次產生 {len(manifest)} 個任務，目錄共 {len(all_entries)} 個 → {out_root}")
    by_fam: dict[str, int] = {}
    for m in manifest:
        by_fam[m["family"]] = by_fam.get(m["family"], 0) + 1
    for fam, cnt in by_fam.items():
        print(f"  {fam:<18} {cnt}")


if __name__ == "__main__":
    main()

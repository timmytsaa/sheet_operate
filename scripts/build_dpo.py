"""把蒸餾產物組成 DPO 偏好對（chosen / rejected）。

負例有兩種來源，價值不同
------------------------
A. teacher_*_rejected.jsonl —— 沒通過 Gym（答案錯或程式碼掛掉）
   便宜、量大，教「不要寫壞掉的程式碼」。

B. 通過 Gym 但被品質關卡剔除的 —— **最值錢**
   它們逐格比對全對，卻硬編欄位索引、靜默吞例外、或換組資料就壞。
   這是「看起來對但方法錯」，正是這個專案在對付的核心病症：
   真實 AVTC 檔上 column=14 猜對過一次，換一份檔就靜默輸出空表。
   只用 A 類負例，模型學到的是「不要出錯」；加上 B 類才學到「即使會對也不要這樣寫」。

B 類靠「原始 teacher 檔 減去 最終資料」還原（以 id + 正規化程式碼比對）。

配對規則
--------
- 同一題的 chosen 取第一筆（品質關卡都過的），與該題每一筆負例配對
- 每題最多 --max-pairs 對，避免少數題目權重過高
- chosen 與 rejected 程式碼實質相同者跳過（沒有偏好訊號）
- prompt（system + user）統一改寫成當前的 SYSTEM_PROMPT——各批資料是不同時期
  蒸餾的，訓練與推論必須看到同一份提示

輸出為 TRL DPOTrainer 的對話格式：
  {"prompt": [system, user], "chosen": [assistant], "rejected": [assistant],
   "id": ..., "family": ..., "reason": ...}

用法：
  python scripts/build_dpo.py --chosen "data/sft/v6_colres_gen.jsonl,data/sft/v7_diff_gen.jsonl" \\
      --raw "data/sft/teacher_v6_*.jsonl,data/sft/teacher_v7_*.jsonl" --out data/sft/dpo_v6v7.jsonl
"""
from __future__ import annotations

import argparse
import glob as _glob
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.prompts import SYSTEM_PROMPT          # noqa: E402

sys.path.insert(0, str(ROOT / "scripts"))
from merge_teacher import normalize_code            # noqa: E402


def load(patterns: str, want_rejected: bool | None = None) -> list[dict]:
    out = []
    for pat in patterns.split(","):
        for fp in sorted(_glob.glob(pat.strip())):
            is_rej = "_rejected" in Path(fp).name
            if want_rejected is not None and is_rej != want_rejected:
                continue
            with open(fp, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        out.append(json.loads(line))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--chosen", required=True, help="最終正樣本檔（逗號分隔）")
    ap.add_argument("--raw", required=True, help="原始 teacher 檔樣式（逗號分隔，含 _rejected）")
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-pairs", type=int, default=4, help="每題最多幾對")
    args = ap.parse_args()

    chosen_recs = load(args.chosen)
    raw_pass = load(args.raw, want_rejected=False)     # 通過 Gym 的全部
    raw_rej = load(args.raw, want_rejected=True)       # 沒通過 Gym 的
    print(f"讀入：最終正樣本 {len(chosen_recs)}、Gym 通過 {len(raw_pass)}、Gym 未通過 {len(raw_rej)}")

    # 每題的 chosen（取第一筆）與其正規化碼集合
    chosen_by: dict[str, dict] = {}
    kept_codes: dict[str, set[str]] = defaultdict(set)
    for r in chosen_recs:
        kept_codes[r["id"]].add(normalize_code(r["messages"][-1]["content"]))
        chosen_by.setdefault(r["id"], r)

    # B 類：通過 Gym 但不在最終資料裡 = 被方法檢查／泛化檢查剔除
    negatives: list[tuple[dict, str]] = []
    for r in raw_rej:
        negatives.append((r, "A_未通過Gym"))
    for r in raw_pass:
        if normalize_code(r["messages"][-1]["content"]) not in kept_codes.get(r["id"], set()):
            negatives.append((r, "B_通過Gym但方法錯"))

    pairs, per_task = [], Counter()
    skipped_same = skipped_nochosen = 0
    for neg, kind in negatives:
        tid = neg["id"]
        ch = chosen_by.get(tid)
        if ch is None:
            skipped_nochosen += 1
            continue
        if per_task[tid] >= args.max_pairs:
            continue
        ch_code = normalize_code(ch["messages"][-1]["content"])
        ng_code = normalize_code(neg["messages"][-1]["content"])
        if ch_code == ng_code:
            skipped_same += 1
            continue
        user_msg = ch["messages"][1]                    # 用 chosen 的 user prompt 為準
        pairs.append({
            "id": tid, "family": ch.get("family", "?"), "neg_kind": kind,
            "reason": neg.get("reason", "方法檢查或泛化檢查剔除"),
            "prompt": [{"role": "system", "content": SYSTEM_PROMPT}, user_msg],
            "chosen": [ch["messages"][-1]],
            "rejected": [neg["messages"][-1]],
        })
        per_task[tid] += 1

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        for p in pairs:
            f.write(json.dumps(p, ensure_ascii=False) + "\n")

    kind_c = Counter(p["neg_kind"] for p in pairs)
    fam_c = Counter(p["family"] for p in pairs)
    print(f"\n產出 {len(pairs)} 對 → {out_path}")
    print(f"  涵蓋題數 {len(per_task)}；每題平均 {len(pairs)/max(len(per_task),1):.2f} 對")
    print(f"  跳過：chosen 與 rejected 實質相同 {skipped_same}、該題沒有 chosen {skipped_nochosen}")
    print("\n負例來源：")
    for k, n in kind_c.most_common():
        print(f"  {k:<22}{n:>5}  ({n/len(pairs):.0%})")
    print("\n各家族：")
    for k, n in sorted(fam_c.items()):
        print(f"  {k:<22}{n:>5}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

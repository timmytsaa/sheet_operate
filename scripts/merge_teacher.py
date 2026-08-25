"""合併多個 teacher 的蒸餾結果，並依「程式碼實質內容」去重。

為什麼要多 teacher：teacher_solve.py 每題只保留第一個通過的樣本就 break，
所以 --k 是重試上限、不是多樣性來源。要讓學生看到「同一原則的多種實作」
（headers.index() / dict comprehension / enumerate 迴圈…），只能靠不同 teacher
各自跑完整題庫，再在這裡合併。

去重規則：同一題若兩個 teacher 產出的程式碼正規化後相同（去註解、去空行、
統一縮排與引號），只留第一份——完全一樣的樣本沒有多樣性價值，只會加重權重。

方法檢查（--reject-literal-index，預設開啟）
------------------------------------------
Gym 驗證的是「輸出」，不是「方法」。teacher 可能用肉眼看 encoder 輸出、判斷
「H 欄，索引 7」然後硬編 row[7]——剛好猜對就通過驗證，卻示範了正好要根除的習慣
（真實 BOM 檔上就是這樣爆的：column=14 # M/S 欄是第14欄）。

判準：合規解法一定先把索引存進變數（row[st_i]、row[pos[h]]），所以字面量
row[N≥1] 就是違規訊號；row[0] 當空列守衛是合法的（參考解法也這樣寫）。
實測 210 筆 v6 樣本抓出 18 筆違規，四個家族的參考解法零誤殺。

用法：
  python scripts/merge_teacher.py --glob "data/sft/teacher_v6_*.jsonl" --out data/sft/v6_colres.jsonl
"""
from __future__ import annotations

import argparse
import glob
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


# 字面量欄位索引＝「用肉眼數欄位」的痕跡；row[0] 例外（空列守衛）
LITERAL_INDEX = re.compile(
    r"row\[\s*[1-9]\d*\s*\]|values\[\s*[1-9]\d*\s*\]|"
    r"\.cell\((?![^)]*value\s*=)[^)]*column\s*=\s*(?!1\b)\d+[^)]*\)\.value")


def normalize_code(text: str) -> str:
    """正規化程式碼，用於判斷「實質相同」：去 markdown 圍籬、註解、空行，統一引號。"""
    body = text
    m = re.search(r"```(?:python)?\n(.*?)```", text, re.S)
    if m:
        body = m.group(1)
    lines = []
    for line in body.splitlines():
        line = re.sub(r"(?<!['\"])#.*$", "", line).rstrip()   # 去行尾註解
        if not line.strip():
            continue
        lines.append(re.sub(r"\s+", " ", line.strip()).replace('"', "'"))
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--glob", required=True, help='輸入檔樣式，例如 "data/sft/teacher_v6_*.jsonl"')
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-per-task", type=int, default=3,
                    help="同一題最多保留幾個不同實作（避免少數題目權重過高）")
    ap.add_argument("--keep-literal-index", action="store_true",
                    help="關閉方法檢查，連硬編欄位索引的樣本也收（不建議）")
    args = ap.parse_args()

    files = sorted(glob.glob(args.glob))
    if not files:
        print(f"找不到符合的檔案：{args.glob}")
        return 1

    by_task: dict[str, list[dict]] = defaultdict(list)
    seen: dict[str, set[str]] = defaultdict(set)
    n_read = n_dup = 0
    per_source: Counter = Counter()
    rejected: list[tuple[str, str, str]] = []
    rej_source: Counter = Counter()

    for fp in files:
        with open(fp, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                rec = json.loads(line)
                n_read += 1
                tid = rec["id"]
                raw = rec["messages"][-1]["content"]
                if not args.keep_literal_index:
                    hit = LITERAL_INDEX.search(raw)
                    if hit:
                        # 通過 Gym 但方法錯——硬編剛好猜對，正是要根除的習慣
                        rejected.append((tid, rec.get("source", "?"), hit.group(0)))
                        rej_source[rec.get("source", "?")] += 1
                        continue
                code = normalize_code(raw)
                if code in seen[tid]:
                    n_dup += 1
                    continue
                seen[tid].add(code)
                by_task[tid].append(rec)
                per_source[rec.get("source", "?")] += 1

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    n_written = 0
    fam_count: Counter = Counter()
    with out_path.open("w", encoding="utf-8") as f:
        for tid in sorted(by_task):
            for rec in by_task[tid][: args.max_per_task]:
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")
                fam_count[rec["family"]] += 1
                n_written += 1

    print(f"輸入 {len(files)} 檔、{n_read} 筆；方法檢查剔除 {len(rejected)} 筆；"
          f"實質重複捨棄 {n_dup} 筆；寫出 {n_written} 筆 → {out_path}")
    print(f"涵蓋題數 {len(by_task)}；平均每題 {n_written / max(len(by_task), 1):.2f} 種實作")

    if rejected:
        print(f"\n方法檢查剔除（通過 Gym 但硬編欄位索引）：{len(rejected)} 筆")
        for src, n in rej_source.most_common():
            print(f"  {src:<40}{n}")
        for tid, src, frag in rejected[:5]:
            print(f"    {tid:<24}{src.split(':')[-1][:16]:<18}{frag}")
        if len(rejected) > 5:
            print(f"    …另外 {len(rejected) - 5} 筆")
    print("\n各 teacher 貢獻（去重後）：")
    for src, n in per_source.most_common():
        print(f"  {src:<40}{n}")
    print("\n各家族筆數：")
    for fam, n in fam_count.most_common():
        print(f"  {fam:<24}{n}")

    # 只有一種實作的題目＝多樣性沒吃到，提醒一下
    single = [t for t in by_task if len(by_task[t]) == 1]
    if single:
        print(f"\n注意：{len(single)}/{len(by_task)} 題只有一種實作"
              f"（teacher 們寫出了實質相同的程式碼，或只有一個 teacher 解出來）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

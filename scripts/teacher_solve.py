"""Teacher 蒸餾：用 Ollama 大模型解任務，經 Gym 驗證後保留通過的軌跡（rejection sampling）。

需要環境變數 OLLAMA_API_KEY（見 sheetops/ollama_client.py）。
支援中斷續跑：已在輸出檔中的任務會自動略過。

負例保存（預設開啟，--no-rejected 可關）
------------------------------------
rejection sampling 原本把沒通過的採樣直接丟掉，但那些不是垃圾：同一個 prompt、
同一個任務，一個通過 Gym、一個沒有——這是**驗證過**的偏好對，正是 DPO 要的
(chosen, rejected)，不必人工標註。

實測 kimi 單一家、單一批（v7 100 題）就丟掉約 49 筆負例；三家跑完 v6+v7 約 265 筆。
那些推論已經花掉了，不存下來純粹是浪費。

寫到 <out 去掉副檔名>_rejected.jsonl，每筆記錄失敗原因（執行失敗的 stderr 尾段，
或分數與第一個不符處）——這同時是「多回合修復訓練」的原料：
(失敗程式碼, 錯誤訊息) → 通過的那一版。

用法：
  python scripts/teacher_solve.py --tasks data/tasks/train --out data/sft/teacher_sft.jsonl --k 4
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
# Colab/IPython 的 sys.stdout 是 ipykernel 的 OutStream，沒有 reconfigure——
# 這些腳本會被 notebook import，不防守就會在 import 當下就炸掉。
if hasattr(sys.stdout, "reconfigure"):
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
    ap.add_argument("--no-rejected", action="store_true",
                    help="不保存未通過的採樣（預設會寫到 <out>_rejected.jsonl 當 DPO 原料）")
    ap.add_argument("--temperature", type=float, default=0.8)
    ap.add_argument("--model", default=None, help="覆寫 OLLAMA_MODEL")
    ap.add_argument("--limit", type=int, default=0, help="只處理前 N 題（0 = 全部）")
    ap.add_argument("--skip-from", default=None,
                    help="glob（如 'data/sft/v5_*.jsonl'）：這些檔案裡已完成的題目一併略過，"
                         "供跨 teacher 救援時不重複解題")
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

    if args.skip_from:
        import glob as _g
        for fp in _g.glob(args.skip_from):
            with open(fp, encoding="utf-8") as f:
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

    n_solved = n_failed = n_rejected = 0
    rate_limited = False
    rej_path = out_path.with_name(out_path.stem + "_rejected.jsonl")
    rej_file = None if args.no_rejected else rej_path.open("a", encoding="utf-8")

    def save_rejected(spec, messages, attempt, code, why, score):
        """沒通過的採樣：DPO 的 rejected 半邊，也是多回合修復訓練的原料。"""
        nonlocal n_rejected
        if rej_file is None:
            return
        rec = {
            "id": spec["id"], "family": spec["family"],
            "source": f"teacher:{client.model}", "attempt": attempt,
            "reason": why, "score": score,
            "messages": messages + [
                {"role": "assistant",
                 "content": ("```python\n" + code + "\n```") if code else "(未產生程式碼區塊)"}],
        }
        rej_file.write(json.dumps(rec, ensure_ascii=False) + "\n")
        rej_file.flush()
        n_rejected += 1

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
                    # 額度用盡要整輪中止：繼續跑只會把每一題都標成「失敗」，
                    # 而且 API 錯誤不經過 save_rejected，負例也收不到。
                    # （實測過一次：41 題被誤標失敗、負例只存下 7 筆。）
                    if "429" in str(e) or "usage limit" in str(e):
                        print(f"\n  ⚠ API 額度用盡（{spec['id']}）：{e}")
                        print("  中止本輪。額度恢復後直接重跑即可——已完成的題目會自動略過。")
                        rate_limited = True
                        break
                    print(f"  API 錯誤（{spec['id']}）：{e}")
                    break
                code = extract_code(reply)
                if not code:
                    save_rejected(spec, messages, attempt, "", "沒有 ```python 程式碼區塊", 0.0)
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
                # 失敗原因：執行掛掉就記 traceback 最後幾行（切行不切字，修復訓練才讀得懂），
                # 答案不符就記前兩個不符處
                if not report.get("exec_ok", True):
                    tb = (report.get("exec_feedback") or "").strip().splitlines()
                    why = "執行失敗：" + " / ".join(x.strip() for x in tb[-3:] if x.strip())
                else:
                    why = "答案不符：" + "；".join((report.get("mismatches") or [])[:2])
                save_rejected(spec, messages, attempt, code, why, float(report.get("score", 0.0)))

            if solved:
                n_solved += 1
                print(f"  [V] {spec['id']}")
            elif rate_limited:
                break                       # 額度用盡：這題不算失敗，留給下次續跑
            else:
                n_failed += 1
                print(f"  [X] {spec['id']}（{args.k} 次採樣皆未通過）")

    if rej_file is not None:
        rej_file.close()

    head = "中止（API 額度用盡，尚未處理的題目留給下次續跑）" if rate_limited else "完成"
    print(f"\n{head}：通過 {n_solved}、未通過 {n_failed} → {out_path}")
    if not args.no_rejected:
        print(f"      負例 {n_rejected} 筆 → {rej_path}（DPO 的 rejected 半邊，以 id 配對）")
    print("未通過的題目可提高 --k 或換更強的 teacher 模型重跑（會自動略過已完成者）。")


if __name__ == "__main__":
    main()

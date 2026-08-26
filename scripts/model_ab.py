"""模型對照評測：同一批 eval、同一個 Gym 驗證器，比較兩個（以上）本機模型。

用途：決定「要不要換 base model」。關鍵比較不是「9B vs 4B」，而是
**裸的候選模型 vs 已訓練的現行模型**——因為換模型的代價是作廢整套已驗證的權重
（SFT + GRPO），只有當裸模型的起點就明顯更高，重跑才有回報。

判準（跑之前先講好，避免看到數字才找理由）：
  候選裸模型 明顯贏過 已訓練模型 → 起點高很多，換值得
  打平或輸                      → 訓練才是主要貢獻者，換是拿已驗證的賭未驗證的

thinking 模型注意事項：
  - 生成慢 2~3 倍，max_tokens 要調高（think 區塊也算 token）
  - ollama_client 的 _LEAK_TAGS 已會清掉洩漏的 <think>/<tool_call> 標記
  - extract_code 取 ```python 區塊，think 區塊在前面不影響

用法：
  python scripts/model_ab.py --models sheetops,ornith --tasks data/tasks/eval_v6,data/tasks/eval_v7
  python scripts/model_ab.py --models sheetops,ornith --tasks ... --limit 8 --max-tokens 4000
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.encoder import encode_workbook          # noqa: E402
from sheetops.executor import extract_code, run_code  # noqa: E402
from sheetops.ollama_client import OllamaClient       # noqa: E402
from sheetops.prompts import SYSTEM_PROMPT, build_user_prompt  # noqa: E402
from sheetops.verifier import verify                  # noqa: E402


def collect(task_roots: list[str], limit_per_root: int) -> list[Path]:
    out: list[Path] = []
    for root in task_roots:
        dirs = sorted(p.parent for p in Path(root).glob("*/task.json"))
        if limit_per_root and limit_per_root < len(dirs):
            stride = len(dirs) // limit_per_root         # 等距抽樣，涵蓋各家族
            dirs = dirs[::stride][:limit_per_root]
        out += dirs
    return out


def run_model(model: str, host: str, task_dirs: list[Path], max_tokens: int,
              temperature: float, timeout: int, work: Path) -> dict:
    client = OllamaClient(model=model, host=host, timeout=timeout)
    per_family: dict[str, list[int]] = defaultdict(list)
    rows, t0 = [], time.time()
    for i, d in enumerate(task_dirs, 1):
        spec = json.loads((d / "task.json").read_text(encoding="utf-8"))
        fam = spec["family"]
        msgs = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": build_user_prompt(
                    spec["instruction"], encode_workbook(d / "start.xlsx"),
                    spec.get("context", ""))}]
        status, score = "", 0.0
        try:
            reply = client.chat(msgs, temperature=temperature, max_tokens=max_tokens)
        except RuntimeError as e:
            status = "API錯"
            if "429" in str(e) or "usage limit" in str(e):
                print(f"  ⚠ 額度用盡，中止：{e}")
                break
        else:
            code = extract_code(reply)
            if not code:
                status = "無程式碼"
            else:
                out = work / f"{model.replace('/', '_').replace(':', '_')}_{spec['id']}.xlsx"
                res = run_code(code, d / "start.xlsx", out, timeout=40)
                if not res.ok:
                    status = "執行失敗"
                else:
                    rep = verify(out, d / "goal.xlsx", spec["check"])
                    score = float(rep["score"])
                    status = "全對" if rep["full_match"] else "不符"
        ok = 1 if status == "全對" else 0
        per_family[fam].append(ok)
        rows.append((spec["id"], fam, status, score))
        print(f"  [{i:>3}/{len(task_dirs)}] {spec['id']:<26}{status:<8}{score:.3f}")
    return {"model": model, "rows": rows, "per_family": per_family,
            "elapsed": time.time() - t0}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--models", required=True, help="逗號分隔的本機 Ollama 模型名")
    ap.add_argument("--tasks", required=True, help="逗號分隔的任務目錄")
    ap.add_argument("--host", default="http://localhost:11434")
    ap.add_argument("--limit", type=int, default=0, help="每個任務目錄最多幾題（0=全部）")
    ap.add_argument("--max-tokens", type=int, default=4000,
                    help="thinking 模型要調高——think 區塊也算 token")
    ap.add_argument("--temperature", type=float, default=0.2)
    ap.add_argument("--timeout", type=int, default=600)
    ap.add_argument("--out", default=None, help="把逐題結果寫成 json")
    args = ap.parse_args()

    task_dirs = collect(args.tasks.split(","), args.limit)
    models = [m.strip() for m in args.models.split(",") if m.strip()]
    print(f"評測 {len(task_dirs)} 題 × {len(models)} 個模型"
          f"（max_tokens={args.max_tokens}, temperature={args.temperature}）\n")

    work = ROOT / "data" / "ab_work"
    work.mkdir(parents=True, exist_ok=True)
    results = []
    for m in models:
        print(f"===== {m} =====")
        results.append(run_model(m, args.host, task_dirs, args.max_tokens,
                                 args.temperature, args.timeout, work))
        print()

    # ---- 對照表 ----
    fams = sorted({f for r in results for f in r["per_family"]})
    w = max(22, max((len(f) for f in fams), default=22) + 2)
    print("=" * (w + 14 * len(results)))
    print(f"{'家族':<{w}}" + "".join(f"{r['model'][:12]:<14}" for r in results))
    print("-" * (w + 14 * len(results)))
    for f in fams:
        line = f"{f:<{w}}"
        for r in results:
            v = r["per_family"].get(f, [])
            line += f"{sum(v)}/{len(v):<12}" if v else f"{'-':<14}"
        print(line)
    print("-" * (w + 14 * len(results)))
    line = f"{'合計':<{w}}"
    for r in results:
        allv = [x for v in r["per_family"].values() for x in v]
        line += f"{sum(allv)}/{len(allv)} ({sum(allv)/max(len(allv),1):.0%})".ljust(14)
    print(line)
    line = f"{'耗時':<{w}}"
    for r in results:
        line += f"{r['elapsed']/60:.1f} 分".ljust(14)
    print(line)
    print("=" * (w + 14 * len(results)))

    if args.out:
        Path(args.out).write_text(json.dumps(
            [{"model": r["model"], "elapsed": r["elapsed"], "rows": r["rows"]} for r in results],
            ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n逐題結果 → {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

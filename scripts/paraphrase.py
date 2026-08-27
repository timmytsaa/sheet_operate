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
# Colab/IPython 的 sys.stdout 是 ipykernel 的 OutStream，沒有 reconfigure——
# 這些腳本會被 notebook import，不防守就會在 import 當下就炸掉。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from sheetops.ollama_client import OllamaClient

PROMPT = """請把下面這句試算表操作指令改寫成另一種自然的繁體中文說法。
要求：
1. 語義完全相同，不可增減任何操作要求
2. 所有數字、工作表名稱、欄位名稱（「」內的文字）必須原樣保留
3. 只輸出改寫後的指令，不要任何解釋

原指令：__INSTRUCTION__"""

# 口語化改寫：真實使用紀錄的指令中位數是 35 字，模板指令是 54 字。
# 差距不全是廢話——訓練題必須指定輸出規格才能驗證——但語氣與冗詞可以壓。
# 最好的風格範例就是使用者自己打過的句子，所以拿 usage_log 當 few-shot。
COLLOQUIAL = """請把下面這句試算表操作指令改寫成「真實上班族會打的樣子」：更口語、更短。

真實使用者實際打過的句子長這樣（風格參考，不要照抄內容）：
__EXAMPLES__

要求：
1. 語義完全相同，不可增減任何操作要求
2. 所有數字、工作表名稱、欄位名稱（「」內的文字）必須原樣保留——這些是驗證用的，少一個就不合格
3. 刪掉「請」「務必」「注意」這類贅詞，以及重複說明同一件事的句子
4. 不要解釋做法、不要提示怎麼解——使用者只會說要什麼，不會說怎麼做
5. 只輸出改寫後的指令，不要任何解釋

原指令：__INSTRUCTION__"""

USAGE_LOG = Path(__file__).resolve().parents[1] / "logs" / "usage_log.jsonl"


def real_examples(limit: int = 8) -> str:
    """從真實使用紀錄挑最短的幾句當風格範例；沒有紀錄就用內建的。"""
    seen: set[str] = set()
    if USAGE_LOG.exists():
        for line in USAGE_LOG.read_text(encoding="utf-8").splitlines():
            try:
                ins = (json.loads(line).get("instruction") or "").strip()
            except json.JSONDecodeError:
                continue
            if 8 <= len(ins) <= 60:
                seen.add(ins)
    picked = sorted(seen, key=len)[:limit]
    if not picked:
        picked = ["訂單依金額由大到小排序", "新增「含稅價」欄位放在最右邊",
                  "幫我挑出金額最高的前 3 筆做成新表"]
    return "\n".join(f"- {x}" for x in picked)


def _tokens(text: str) -> set[str]:
    nums = set(re.findall(r"\d+(?:\.\d+)?", text))
    names = set(re.findall(r"「([^」]+)」", text))
    return nums | names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="data/tasks/train")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--model", default=None)
    ap.add_argument("--style", choices=["natural", "colloquial"], default="natural",
                    help="colloquial：以真實使用紀錄為範例，改寫得更短更口語")
    args = ap.parse_args()

    client = OllamaClient(model=args.model)
    ok, msg = client.available()
    if not ok:
        print(msg)
        sys.exit(1)

    template = PROMPT
    if args.style == "colloquial":
        ex = real_examples()
        template = COLLOQUIAL.replace("__EXAMPLES__", ex)
        print(f"口語化模式，風格範例 {len(ex.splitlines())} 句（取自 logs/usage_log.jsonl）")

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
                [{"role": "user", "content": template.replace("__INSTRUCTION__", seed)}],
                temperature=0.9).strip().strip("「」\"'")
        except RuntimeError as e:
            print(f"  API 錯誤：{e}")
            break

        max_len = len(seed) if args.style == "colloquial" else len(seed) * 3
        if new and _tokens(seed) <= _tokens(new) and 5 < len(new) <= max_len:
            spec["instruction"] = new
            spec["instruction_seed"] = seed
            tf.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
            n_ok += 1
        else:
            n_keep += 1

    print(f"改寫完成 {n_ok} 筆；維持原句 {n_keep} 筆；先前已改寫跳過 {n_done} 筆")
    lens = []
    for tf in task_files:
        sp = json.loads(tf.read_text(encoding="utf-8"))
        lens.append((len(sp.get("instruction_seed") or sp["instruction"]), len(sp["instruction"])))
    if lens:
        before = sorted(x for x, _ in lens); after = sorted(y for _, y in lens)
        m = len(lens) // 2
        print(f"指令長度中位數：{before[m]} → {after[m]} 字（真實使用者 35 字）")


if __name__ == "__main__":
    main()

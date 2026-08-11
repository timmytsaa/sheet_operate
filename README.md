# sheet_operate — 試算表操作模型訓練

訓練一個能依繁體中文指令操作 .xlsx 工作簿的模型。方法整合三篇論文：

| 論文 | 借鏡 |
|---|---|
| [Spreadsheet-RL](https://arxiv.org/abs/2605.22642) | 主軸：可執行沙盒（Gym）+ 可驗證 reward + SFT→RL 訓練 |
| [SpreadsheetLLM](https://arxiv.org/abs/2407.09025) | 表格緊湊編碼（座標保留、格式附註、頭尾抽樣） |
| [SheetMind](https://arxiv.org/abs/2506.12339) | 反面教材：純 prompting 天花板低，需靠訓練突破 |

**技術路線**：學生模型用 Qwen3-4B-Instruct-2507 → 合成資料 + teacher 蒸餾（Ollama `qwen3.5:397b`）→ Colab Pro+（96GB）上 SFT → GRPO。
（2026-08-11 定案：蒸餾資料為純程式碼目標，與 Instruct 形態匹配、GRPO rollout 也快得多；
論文原用 Thinking 版，若之後組合任務卡在推理瓶頸再評估切換並補思考軌跡資料。）

## 架構

```
sheetops/
  encoder.py        .xlsx → 緊湊文字表示（模型的觀察）
  executor.py       子行程沙盒執行模型產生的 Python 程式碼
  verifier.py       結果 vs 目標工作簿比對 → reward（值 + 格式）
  env.py            Spreadsheet Gym：reset/step 多回合環境
  prompts.py        SFT / teacher / rollout 共用提示詞
  zh_data.py        繁中（台灣）商業假資料詞庫
  taskgen/          10 個合成任務家族（起始表+指令+目標表+參考解法）
  ollama_client.py  Ollama 雲端/本地 API 客戶端
scripts/
  selftest.py           端到端自測：參考解法必須全數在 Gym 拿滿分
  gen_tasks.py          批次產生任務（train/eval 用不同 seed 隔離）
  build_sft_dataset.py  任務 → SFT chat jsonl（種子資料）
  teacher_solve.py      teacher 蒸餾 + rejection sampling（需 OLLAMA_API_KEY）
  paraphrase.py         指令自然化改寫（需 OLLAMA_API_KEY）
data/
  tasks/train/  300 題（10 家族 × 30）   tasks/eval/  50 題（獨立 seed）
  sft/seed_sft.jsonl  300 筆種子 SFT 資料
```

任務契約：模型讀 `INPUT_PATH` 工作簿，執行操作後存到 `OUTPUT_PATH`；
reward = 與目標工作簿的匹配率（值逐格比對 + 指定格式檢查），全對才算 solved。

## 任務家族（v1）

filter_rows（條件篩選）、sort_rows（排序）、groupby_summary（分類彙總）、
compute_column（計算欄位）、total_row（總計列）、format_style（粗體/底色/數值格式/紅字）、
clean_data（去重/去空白/補空值）、join_lookup（跨表查價回填）、
split_concat（日期拆欄/姓名合併）、top_n（前 N 筆到新表）、
**composite（組合式多步：篩選→排序→總計、篩選→彙總、清理→計算等，一條指令 2~3 個依序操作）**。

每題由「同一段邏輯」同時產生目標表與參考解法，保證可解且驗證器判分一致（`selftest.py` 把關）。

## 快速開始

```bash
pip install -r requirements.txt
python scripts/selftest.py            # 應顯示 通過 30/30
python scripts/gen_tasks.py --out data/tasks/train --n 30 --seed 20260811
python scripts/build_sft_dataset.py   # → data/sft/seed_sft.jsonl
```

Teacher 蒸餾：先編輯專案根目錄的 `.env`，在 `OLLAMA_API_KEY=` 後貼上金鑰（此檔已列入 .gitignore），然後：

```bash
python scripts/teacher_solve.py --tasks data/tasks/train --k 4   # 可中斷續跑
python scripts/paraphrase.py --tasks data/tasks/train            # 指令多樣化
```

## 訓練 Roadmap

- [x] **Phase 1 基礎設施**：encoder / gym / verifier / taskgen / teacher 接口（本 repo）
- [ ] **Phase 2 資料**：teacher 蒸餾軌跡（目標 ≥2k 筆）、指令改寫、難題組合（多步驟任務）
- [ ] **Phase 3 SFT**：Colab Pro+ 96GB，Qwen3-4B-Instruct-2507 全參數 SFT（bf16 + 8-bit optimizer），
      checkpoint 存 Google Drive；eval 集 pass@1 追蹤
- [ ] **Phase 4 GRPO**：TRL/Unsloth GRPOTrainer + 本 Gym reward（score，全對加成）；
      沙盒程式碼在 Colab 端執行，建議加一層隔離
- [ ] **Phase 5 部署**：merge → GGUF → 本地 Ollama 服務
- [ ] **Phase 6（擴充）影像入口**：掃描件/照片/PDF 報表 → [PaddleOCR-VL-1.6](https://huggingface.co/PaddlePaddle/PaddleOCR-VL-1.6)
      （0.9B 文件解析模型）辨識表格結構 → 轉 .xlsx → 交給操作模型執行指令。
      與訓練主線完全解耦，僅作為 pipeline 前端元件（不是 base model）。

## 設計備註

- 驗證以「值」為準；模型若寫入 Excel 公式，驗證器會嘗試用選用套件 `formulas` 求值後比對
  （未安裝則該格判不符）。指令與 system prompt 已引導模型直接寫入計算值。
- 執行沙盒 v1 = 子行程 + 逾時 + 獨立工作目錄；RL rollout 大規模跑模型產碼前應再加強隔離。
- eval 集（seed 900001）只用於評測，不得進訓練資料。

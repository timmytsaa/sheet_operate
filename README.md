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
**composite（組合式多步：一條指令 2~3 個依序操作）**、
**context_rule（關鍵規則只在【補充說明】——訓練模型遵循使用者自訂規則）**、
**large_table（120~260 列大表，觀察被截斷，程式碼必須泛化到看不見的列）**。

表格外皮由 5 套 schema 輪換（訂單/報銷明細/庫存清單/工時紀錄/銷售紀錄）；
第 6 套 hr（出勤紀錄）**只進 OOD 評測集**（`data/tasks/eval_ood`），訓練從未見過。
模型產生的程式碼執行前先過 `sheetops/safety.py` AST 白名單檢查（封鎖 os/shutil/網路/eval 等）。

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
- [ ] **Phase 3 SFT**：Colab Pro+，Qwen3-4B-Instruct-2507 + LoRA r=64（抗遺忘），
      混 10~15% 通用中文指令資料（TaiwanChat），只對 assistant 回覆算 loss，
      Drive checkpoint 斷點續訓 → `notebooks/colab_sft_lora.ipynb`
      （流程：裸模型基準線 eval → SFT → 訓後 eval 對比 → 通用能力抽查）
- [x] **Phase 4 GRPO**：`notebooks/colab_grpo.ipynb`——從 sft_v2 接續、reward = Gym 執行驗證。
      **成果：v1 100%（104 題）、OOD 98.5%、v2 難度階梯 93.8%**（僅 column_ops 62.5% 待補）
- [ ] **Phase 5 部署**（進行中）：
      1. Colab 跑 `notebooks/colab_export_gguf.ipynb`（merge adapter → GGUF Q8_0 → Drive）
      2. 下載 `sheetops-q8_0.gguf` 到 `deploy/`，執行 `ollama create sheetops -f deploy/Modelfile`
      3. CLI：`python scripts/sheetops_cli.py 報表.xlsx "指令"`（預覽變更→確認→輸出副本；
         `--in-place` 自動備份、`--context` 帶自訂規則、`--gguf` 可跳過 Ollama 直連 llama.cpp）
      4. **網頁版（給同事用）**：`python scripts/serve_web.py` → http://localhost:8033
         （區網分享 http://<你的IP>:8033；上傳→指令→預覽→確認下載；原檔永不覆寫；
         單卡全域鎖排隊；`logs/usage_log.jsonl` 記錄指令/程式碼/採納與否 = v3 與 DPO 原料）
- [ ] **Phase 6（擴充）影像入口**：掃描件/照片/PDF 報表 → [PaddleOCR-VL-1.6](https://huggingface.co/PaddlePaddle/PaddleOCR-VL-1.6)
      （0.9B 文件解析模型）辨識表格結構 → 轉 .xlsx → 交給操作模型執行指令。
      與訓練主線完全解耦，僅作為 pipeline 前端元件（不是 base model）。

## 設計備註

- 驗證以「值」為準；模型若寫入 Excel 公式，驗證器會嘗試用選用套件 `formulas` 求值後比對
  （未安裝則該格判不符）。指令與 system prompt 已引導模型直接寫入計算值。
- 執行沙盒 v1 = 子行程 + 逾時 + 獨立工作目錄；RL rollout 大規模跑模型產碼前應再加強隔離。
- eval 集（seed 900001）只用於評測，不得進訓練資料。

"""sheetops — 試算表操作模型的訓練基礎設施。

模組總覽：
- encoder    : 將 .xlsx 工作簿編碼為緊湊文字表示（餵給 LLM 的觀察）
- executor   : 在子行程沙盒中執行模型產生的 Python 程式碼
- verifier   : 比對執行結果與目標工作簿，計算 reward
- env        : 多回合互動環境（reset / step），SFT 資料驗證與 RL rollout 共用
- taskgen    : 合成任務產生器（起始表 + 繁中指令 + 目標表 + 參考解法）
- ollama_client : Ollama 雲端 API 客戶端（teacher 蒸餾 / 指令改寫用）
"""

__version__ = "0.1.0"

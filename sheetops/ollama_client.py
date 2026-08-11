"""Ollama API 客戶端（雲端或本地）。

設定來源（優先序）：建構子參數 > 環境變數 > 專案根目錄 .env 檔
  OLLAMA_API_KEY  雲端金鑰（host 為 ollama.com 時必填）
  OLLAMA_HOST     預設 https://ollama.com；本地改 http://localhost:11434
  OLLAMA_MODEL    預設 qwen3-coder:480b（teacher 用大模型）
"""
from __future__ import annotations

import os
import time
from pathlib import Path

import requests

DEFAULT_HOST = "https://ollama.com"
DEFAULT_MODEL = "qwen3.5:397b"


def _load_dotenv() -> None:
    """讀取專案根目錄的 .env（已設定的環境變數不覆寫）。無外部依賴。"""
    env_file = Path(__file__).resolve().parents[1] / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        k, v = k.strip(), v.strip().strip('"').strip("'")
        if k and v and k not in os.environ:
            os.environ[k] = v


class OllamaClient:
    def __init__(self, host: str | None = None, api_key: str | None = None,
                 model: str | None = None, timeout: int = 600):
        _load_dotenv()
        self.host = (host or os.environ.get("OLLAMA_HOST") or DEFAULT_HOST).rstrip("/")
        self.api_key = api_key or os.environ.get("OLLAMA_API_KEY") or ""
        self.model = model or os.environ.get("OLLAMA_MODEL") or DEFAULT_MODEL
        self.timeout = timeout

    def available(self) -> tuple[bool, str]:
        if "ollama.com" in self.host and not self.api_key:
            return False, ("尚未設定 OLLAMA_API_KEY。\n"
                           "請開啟專案根目錄的 .env 檔，在 OLLAMA_API_KEY= 後面貼上金鑰。\n"
                           "（或把 OLLAMA_HOST 指向本地 http://localhost:11434）")
        return True, ""

    def _headers(self) -> dict:
        h = {"Content-Type": "application/json"}
        if self.api_key:
            h["Authorization"] = f"Bearer {self.api_key}"
        return h

    def chat(self, messages: list[dict], temperature: float = 0.7,
             max_tokens: int | None = None, retries: int = 3) -> str:
        body: dict = {"model": self.model, "messages": messages, "stream": False,
                      "options": {"temperature": temperature}}
        if max_tokens:
            body["options"]["num_predict"] = max_tokens

        last_err: Exception | None = None
        for attempt in range(retries):
            try:
                resp = requests.post(f"{self.host}/api/chat", json=body,
                                     headers=self._headers(), timeout=self.timeout)
            except Exception as e:          # 網路層錯誤 → 重試
                last_err = e
            else:
                if resp.status_code == 200:
                    return resp.json()["message"]["content"]
                if 400 <= resp.status_code < 500:   # 用戶端錯誤（模型退役、額度等）→ 不重試
                    raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")
                last_err = RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")
            if attempt < retries - 1:
                time.sleep(2 ** attempt * 2)
        raise RuntimeError(f"Ollama API 呼叫失敗（重試 {retries} 次）: {last_err}")

    def list_models(self) -> list[str]:
        resp = requests.get(f"{self.host}/api/tags", headers=self._headers(), timeout=30)
        resp.raise_for_status()
        return [m["name"] for m in resp.json().get("models", [])]

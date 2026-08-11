"""Spreadsheet Gym：多回合互動環境。

任務目錄格式（taskgen 產生）：
    task_dir/
      start.xlsx   起始工作簿
      goal.xlsx    目標工作簿
      task.json    {id, family, instruction, check, ref_solution, meta}

使用方式：
    env = SpreadsheetEnv(task_dir)
    obs = env.reset()                    # {"instruction", "sheet_text", "turn"}
    obs, score, done, info = env.step(code)

reward = verifier 的 score（0~1），done 於 full_match 或回合數用盡。
"""
from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

from .encoder import encode_workbook
from .executor import ExecResult, run_code
from .verifier import verify


class SpreadsheetEnv:
    def __init__(self, task_dir: str | Path, max_turns: int = 3,
                 timeout: int = 30, max_rows_obs: int = 40, keep_workdir: bool = False):
        self.task_dir = Path(task_dir)
        self.max_turns = max_turns
        self.timeout = timeout
        self.max_rows_obs = max_rows_obs
        self.keep_workdir = keep_workdir

        self.task = json.loads((self.task_dir / "task.json").read_text(encoding="utf-8"))
        self.start_path = self.task_dir / "start.xlsx"
        self.goal_path = self.task_dir / "goal.xlsx"

        self._work: Path | None = None
        self.turn = 0

    # ------------------------------------------------------------------
    def reset(self) -> dict:
        self._cleanup()
        self._work = Path(tempfile.mkdtemp(prefix=f"sheetops_env_{self.task['id']}_"))
        shutil.copy(self.start_path, self.state_path)
        self.turn = 0
        return {
            "instruction": self.task["instruction"],
            "sheet_text": encode_workbook(self.state_path, self.max_rows_obs),
            "turn": 0,
        }

    @property
    def state_path(self) -> Path:
        assert self._work is not None, "先呼叫 reset()"
        return self._work / "state.xlsx"

    # ------------------------------------------------------------------
    def step(self, code: str):
        assert self._work is not None, "先呼叫 reset()"
        self.turn += 1
        next_path = self._work / f"turn{self.turn}.xlsx"

        exec_result: ExecResult = run_code(code, self.state_path, next_path, self.timeout)
        if exec_result.ok:
            shutil.copy(next_path, self.state_path)

        report = verify(self.state_path, self.goal_path, self.task.get("check"))
        done = bool(report["full_match"]) or self.turn >= self.max_turns

        obs = {
            "instruction": self.task["instruction"],
            "sheet_text": encode_workbook(self.state_path, self.max_rows_obs),
            "exec_feedback": exec_result.feedback(),
            "turn": self.turn,
        }
        info = {"exec": exec_result, "verify": report}
        if done and not self.keep_workdir:
            self._cleanup()
        return obs, report["score"], done, info

    # ------------------------------------------------------------------
    def _cleanup(self):
        if self._work is not None and not self.keep_workdir:
            shutil.rmtree(self._work, ignore_errors=True)
        self._work = None

    def close(self):
        self._cleanup()


def solve_once(task_dir: str | Path, code: str, **env_kw) -> dict:
    """單回合工具函式：跑一段程式碼並回傳驗證報告（rejection sampling 用）。"""
    env = SpreadsheetEnv(task_dir, **env_kw)
    env.reset()
    _obs, score, _done, info = env.step(code)
    env.close()
    report = info["verify"]
    report["exec_ok"] = info["exec"].ok
    report["exec_feedback"] = info["exec"].feedback()
    return report

from random import Random
from pathlib import Path

from .base import TaskSpec, write_task
from .families import FAMILIES


def generate_task(family: str, rng: Random, task_id: str, out_root: str | Path) -> Path:
    spec, start_wb, goal_wb = FAMILIES[family](rng, task_id)
    return write_task(out_root, spec, start_wb, goal_wb)

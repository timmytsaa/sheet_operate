"""模型產生程式碼的靜態安全檢查（AST 白名單）。

定位：防呆與縱深防禦，不是完整的安全邊界——模型幻覺一行 shutil.rmtree
或亂開網路連線時直接擋下，但無法防禦刻意的沙盒逃逸（那要靠容器層）。

規則：
- import 只允許白名單模組（openpyxl 及純運算標準庫）
- 封鎖危險內建呼叫（eval / exec / open / __import__ …）
- 封鎖雙底線屬性存取（__globals__ / __subclasses__ 這類逃逸跳板）
- 語法錯誤不在此攔（放行讓執行器回報自然的錯誤訊息，供模型修正）
"""
from __future__ import annotations

import ast

ALLOWED_IMPORTS = {
    "openpyxl", "datetime", "math", "re", "collections", "itertools",
    "functools", "statistics", "string", "decimal", "copy", "json",
    "unicodedata", "bisect", "heapq", "operator",
}

# exit()/quit() 不封鎖：子行程內無害（頂多提前結束，缺輸出會被自然回報）
BLOCKED_CALLS = {
    "eval", "exec", "compile", "__import__", "open", "input",
    "breakpoint", "globals", "locals", "vars",
}


def check_code_safety(code: str) -> tuple[bool, str]:
    """回傳 (ok, reason)。語法錯誤視為通過（交給執行器產生自然回饋）。"""
    try:
        tree = ast.parse(code)
    except SyntaxError:
        return True, ""

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                root = alias.name.split(".")[0]
                if root not in ALLOWED_IMPORTS:
                    return False, f"不允許 import {alias.name}（白名單外模組）"
        elif isinstance(node, ast.ImportFrom):
            root = (node.module or "").split(".")[0]
            if root not in ALLOWED_IMPORTS:
                return False, f"不允許 from {node.module} import ...（白名單外模組）"
        elif isinstance(node, ast.Call):
            fn = node.func
            if isinstance(fn, ast.Name) and fn.id in BLOCKED_CALLS:
                return False, f"不允許呼叫 {fn.id}()"
        elif isinstance(node, ast.Attribute):
            if node.attr.startswith("__") and node.attr.endswith("__"):
                return False, f"不允許存取雙底線屬性 {node.attr}"
        elif isinstance(node, ast.Name):
            if node.id in ("__builtins__", "__loader__", "__spec__"):
                return False, f"不允許存取 {node.id}"
    return True, ""

"""從「模型產生的程式碼」推導出它實際做了什麼——不採信模型自己的說法。

為什麼不用模型的推斷註解
------------------------
規則 6 訓練模型寫 `# 推斷：資料表=X｜欄位=Y｜…`，那是**跟程式碼同時生成的一段文字**，
不是對程式碼的觀察。它可以憑空編造：真實案例裡模型宣稱「鍵=Find Number」，
而那份工作簿根本沒有這個欄位，程式碼裡也沒出現過這個字串。

這支模組改用 AST 解析程式碼本身，回報：開了哪些工作表、讀第幾列當表頭、
查了哪些欄名、輸出到哪張新表。這些是**事實**，不是宣稱。

最有價值的是 contradictions()：把推斷宣稱的欄名拿去跟程式碼比對，
「說了但程式碼裡沒有」就是模型在編造——那本身就是最強的警訊。
"""
from __future__ import annotations

import ast
import re


def _str_args(node: ast.AST) -> list[str]:
    return [a.value for a in getattr(node, "args", [])
            if isinstance(a, ast.Constant) and isinstance(a.value, str)]


def analyze(code: str) -> dict:
    """回傳 {sheets, created, header_rows, columns, ok, error}。解析失敗時 ok=False。"""
    out = {"sheets": [], "created": [], "header_rows": [], "columns": [],
           "ok": True, "error": ""}
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        return {**out, "ok": False, "error": f"程式碼無法解析：{e.msg}"}

    sheets, created, rows, cols = [], [], [], []

    for n in ast.walk(tree):
        # wb["工作表"]／wb.get_sheet_by_name("…")
        if isinstance(n, ast.Subscript) and isinstance(n.slice, ast.Constant) \
                and isinstance(n.slice.value, str):
            base = n.value
            name = getattr(base, "id", "") or getattr(base, "attr", "")
            if "wb" in name.lower() or "book" in name.lower():
                sheets.append(n.slice.value)
            else:
                cols.append(n.slice.value)      # pos["欄名"] 這種查表
        if isinstance(n, ast.Call):
            fn = n.func
            attr = getattr(fn, "attr", "")
            if attr == "create_sheet":
                created += _str_args(n)
                for kw in n.keywords:
                    if kw.arg in ("title",) and isinstance(kw.value, ast.Constant):
                        created.append(kw.value.value)
            elif attr == "index":                # headers.index("欄名")
                cols += _str_args(n)
            elif attr in ("get", "setdefault"):  # pos.get("欄名")
                cols += _str_args(n)
            elif attr == "iter_rows":
                for kw in n.keywords:
                    if kw.arg == "min_row" and isinstance(kw.value, ast.Constant):
                        rows.append(("資料起始列", kw.value.value))
            elif attr == "cell":
                for kw in n.keywords:
                    if kw.arg == "row" and isinstance(kw.value, ast.Constant):
                        rows.append(("讀取列", kw.value.value))
        # ws[1] / ws[2] —— 整列取值，通常就是表頭
        if isinstance(n, ast.Subscript) and isinstance(n.slice, ast.Constant) \
                and isinstance(n.slice.value, int):
            name = getattr(n.value, "id", "")
            if name.startswith("ws") or name.endswith("ws") or "sheet" in name.lower():
                rows.append(("整列讀取", n.slice.value))
        # 與字串比較：if row[i] == "M"
        if isinstance(n, ast.Compare):
            for c in n.comparators:
                if isinstance(c, ast.Constant) and isinstance(c.value, str) and c.value:
                    cols.append(c.value)

    def uniq(seq):
        seen, out_ = set(), []
        for x in seq:
            if x not in seen:
                seen.add(x)
                out_.append(x)
        return out_

    out["sheets"] = uniq(s for s in sheets if s not in created)
    out["created"] = uniq(created)
    out["header_rows"] = uniq(rows)
    out["columns"] = uniq(c for c in cols if len(c) <= 40)
    return out


_CLAIM = re.compile(r"[｜|]\s*(?:欄位|鍵)\s*=\s*([^｜|]+)")


# 「用肉眼數欄位」的痕跡。這裡是單一來源——scripts/merge_teacher.py 的方法檢查、
# GRPO 的 reward 懲罰、以及下面的解釋都用同一份判準，避免三處各寫一套而漂移。
# row[0] 例外（空列守衛，參考解法也這樣寫）。
LITERAL_INDEX = re.compile(
    r"row\[\s*[1-9]\d*\s*\]|values\[\s*[1-9]\d*\s*\]|"
    r"\.cell\((?![^)]*value\s*=)[^)]*column\s*=\s*(?!1\b)\d+[^)]*\)\.value")

SWALLOW = re.compile(r"except[^\n]*:\s*\n\s*(?:pass|continue)\s*(?:\n|$)")


def _strip_comments(code: str) -> str:
    """比對前要去掉註解——推斷本身就寫在註解裡，不去掉的話任何宣稱都「找得到」。"""
    return "\n".join(l.split("#", 1)[0] for l in (code or "").splitlines())


def contradictions(inference: str, code: str) -> list[str]:
    """推斷宣稱的欄名／鍵，在程式碼裡完全找不到 → 模型在編造。

    真實案例：推斷寫「鍵=Find Number」，但那份工作簿沒有這個欄位，
    程式碼裡也沒有這個字串——使用者若只看推斷，會以為它抓對了。
    """
    if not inference or not code:
        return []
    body = _strip_comments(code)
    bad = []
    for seg in _CLAIM.findall(inference):
        for name in re.split(r"[、,／/]", seg):
            name = re.sub(r"（[^）]*）", "", name).strip()
            # 太短或明顯是描述詞就跳過，避免誤報
            if len(name) < 2 or name in ("無", "None", "不適用"):
                continue
            if name not in body:
                bad.append(name)
    return sorted(set(bad))


def render(code: str, inference: str = "") -> str:
    """組成給使用者看的「程式碼實際做了什麼」。"""
    a = analyze(code)
    if not a["ok"]:
        return a["error"]
    lines = []
    if a["sheets"]:
        lines.append("讀取工作表：" + "、".join(a["sheets"]))
    hr = [f"第 {v} 列（{k}）" for k, v in a["header_rows"]]
    if hr:
        lines.append("碰到的列：" + "、".join(hr[:4]))
    if a["columns"]:
        lines.append("用到的欄名／值：" + "、".join(a["columns"][:8]))
    if a["created"]:
        lines.append("新增工作表：" + "、".join(a["created"]))
    body = _strip_comments(code)
    hit = LITERAL_INDEX.search(body)
    if hit:
        lines.append(f"⚠ 用了硬編欄位索引 {hit.group(0)}——沒有依欄名查找，換一份檔就可能取錯欄")
    if SWALLOW.search(body):
        lines.append("⚠ 有 except: pass——出錯的資料會被靜默略過，總數可能不對")
    bad = contradictions(inference, code)
    if bad:
        lines.append("⚠ 推斷提到「" + "、".join(bad) + "」，但程式碼裡沒有——模型可能講錯了")
    return "\n".join(lines) if lines else "（無法從程式碼判讀）"

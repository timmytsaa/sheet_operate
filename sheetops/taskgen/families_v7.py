"""v7 任務家族：兩表差異比對（sheet_diff）。

來源：七個探針實測部署中的模型（真實 BOM 資料改造）。基本的三分類 diff
（新增/刪除/變更）邏輯它已經會了——9 個差異全中；壞的全在「鍵怎麼處理」和
「輸出要帶什麼」：

  探針                    結果            機制
  鍵有空白（指令不提醒）   29 列 vs 正解 1   new_ui_items.add(值) 完全不正規化
  鍵有空白（指令有提醒）   正確             明講就會 strip
  重複鍵 M/S            輸出毀損         f"{料號}{M/S}" 再 split('-')，料號本身含連字號
  空鍵列                多一列           UI Item Number 為空的列被當成「只在舊版有」
  多欄比對              漏 1/3           抽了 Qty 卻沒寫比較分支
  帶出其他欄位           全是 None        建索引時只存了比較用的那一欄
  欄名不同＋表頭位置不同    正確 ✅          （不需訓練）
  顏色標記              正確 ✅          （不需訓練）

最危險的是第一個：輸出 29 列、格式正常、料號正常，看起來就是一份差異報告，
只是 28 筆是假的。靜默錯誤比當機危險得多，而 Gym 只驗輸出正是在這裡失效。

五個變體
--------
diff_dirty_key   鍵有空白/大小寫/型別差異，且指令不提醒 → 必須自己正規化
diff_nullkey     表中有小計列、註記列（鍵欄空白）→ 不可算成差異
diff_dupkey      同鍵多列（主/替代料）→ 用 tuple 複合鍵，不可字串黏合再 split
diff_carry_cols  差異表要帶出指令指定的其他欄位，不是只有鍵和比較值
diff_multicol    指令列出 N 個欄位就要有 N 個比較分支

設計原則
--------
1. 沿用 v6 的 _assert_discriminates()：天真作法（不正規化／dict 覆蓋／只比一欄）
   必須得到不同答案，否則這題沒有鑑別度，直接重抽。
2. 「資料消歧、不是指令消歧」（v5 原則）：指令不提醒要 strip，但同一料號在兩表
   的品名/類別完全相同，看得出來是同一個料——答案仍然唯一。
3. 指令措辭明確列舉要比的欄位，不用「除了 X」這種有歧義的說法。
4. 四套 schema 輪換（BOM／報價／盤點／名冊），避免學成 BOM 專用。
"""
from __future__ import annotations

from random import Random

from .. import zh_data
from .base import TaskSpec, new_wb, write_table

# ----------------------------------------------------------------------
# Schema：欄位「角色」一致（key / name / num / cat / unit），外皮不同
# ----------------------------------------------------------------------

SCHEMAS = [
    {"tag": "bom", "old": "舊版BOM", "new": "新版BOM", "thing": "料號",
     "key": "料號", "name": "品名", "num": "數量", "cat": "類別", "unit": "單位",
     "kfmt": lambda r: f"{r.choice('345')}{r.randint(10,99)}-{r.randint(1000,9999):05d}",
     "cats": ["塑膠件", "金屬件", "五金", "線材", "包材"],
     "units": ["EA", "PCS", "SET"],
     "names": ["上蓋", "下蓋", "中框", "導光柱", "散熱片", "彈簧", "螺絲 M3x6",
               "泡棉墊", "銅柱", "排線", "天線", "絕緣片", "麥拉片", "腳墊",
               "防塵蓋", "卡榫", "遮光罩", "固定夾", "束帶", "標籤", "軸承", "銘板"]},
    {"tag": "price", "old": "上期報價", "new": "本期報價", "thing": "品號",
     "key": "品號", "name": "品名", "num": "單價", "cat": "幣別", "unit": "計價單位",
     "kfmt": lambda r: f"{r.choice('ABCDE')}-{r.randint(1000, 9999)}",
     "cats": ["TWD", "USD", "CNY"], "units": ["個", "箱", "公斤"],
     "names": zh_data.INVENTORY_ITEMS[:22]},
    {"tag": "stock", "old": "系統帳", "new": "實盤", "thing": "料號",
     "key": "料號", "name": "品名", "num": "數量", "cat": "儲位", "unit": "單位",
     "kfmt": lambda r: f"MAT-{r.randint(10000, 99999)}",
     "cats": ["A01", "A02", "B11", "B12", "C21"], "units": ["EA", "箱", "包"],
     "names": zh_data.INVENTORY_ITEMS[:22]},
    {"tag": "roster", "old": "上月名冊", "new": "本月名冊", "thing": "員工編號",
     "key": "員工編號", "name": "姓名", "num": "時數", "cat": "部門", "unit": "職稱",
     "kfmt": lambda r: f"EMP-{r.randint(1000, 9999)}",
     "cats": zh_data.DEPARTMENTS, "units": ["工程師", "專員", "組長", "副理"],
     "names": None},   # 用 zh_data 產生人名
]


def _names(s: dict, rng: Random, n: int) -> list[str]:
    """zh_data 的品項是 (名稱, 分類, 低價, 高價) tuple——只取名稱。"""
    if s["names"]:
        pool = [x[0] if isinstance(x, tuple) else x for x in s["names"]]
        rng.shuffle(pool)
        base = pool[:]
        i = 1
        while len(pool) < n:                       # 詞庫不夠就加序號，維持唯一
            pool += [f"{x}-{i}" for x in base]
            i += 1
        return pool[:n]
    return [rng.choice(zh_data.SURNAMES) + rng.choice(zh_data.GIVEN_NAMES) for _ in range(n)]


def _base_rows(s: dict, rng: Random, n: int) -> tuple[list[str], list[list]]:
    """產生一份基準資料；鍵保證唯一（重複鍵的情境由 diff_dupkey 另外造）。"""
    headers = [s["key"], s["name"], s["num"], s["cat"], s["unit"]]
    keys, seen = [], set()
    while len(keys) < n:
        k = s["kfmt"](rng)
        if k not in seen:
            seen.add(k)
            keys.append(k)
    names = _names(s, rng, n)
    rows = [[keys[i], names[i], rng.randint(1, 20), rng.choice(s["cats"]),
             rng.choice(s["units"])] for i in range(n)]
    return headers, rows


def _tier(rng: Random) -> str:
    """指令揭露層級。

    真實使用紀錄（logs/usage_log.jsonl，85 筆）的指令長度中位數是 35 字，
    而 v6/v7 初版寫出來是 92~95 字——因為我把「陷阱怎麼處理」直接寫進指令，
    等於送答案。這裡分三層：

      full  完整說明陷阱（「注意鍵不是唯一的，要各自比對」）
      mid   只暗示有異常，不說怎麼處理（「兩張表的料號格式不完全一致」）
      terse 完全不提陷阱——真實使用者的樣子

    砍掉的只有「陷阱提示」，輸出欄位與排序一律保留（驗證需要唯一答案）。
    答案仍然唯一，因為重複鍵/空白/小計列在資料裡看得出來——
    這是 v5 的「資料消歧、不是指令消歧」原則。
    """
    r = rng.random()
    return "full" if r < 0.34 else ("mid" if r < 0.67 else "terse")


def _compose(rng: Random, tiers: dict[str, list[str]]) -> str:
    return rng.choice(tiers[_tier(rng)])


def _assert_discriminates(expected: list, naive: list, msg: str, min_n: int = 1) -> None:
    """天真作法必須得到不同答案，否則這題沒有鑑別度。"""
    assert len(expected) >= min_n, f"{msg}（正解列數太少：{len(expected)}）"
    e = [tuple(map(str, r)) for r in expected]
    nv = [tuple(map(str, r)) for r in naive]
    assert e != nv, msg


def _pick_schema(rng: Random) -> dict:
    return rng.choice(SCHEMAS)


def _finish(task_id: str, family: str, s: dict, instruction: str, out_cols: list,
            expected: list, sheets: list[tuple[str, list, list]], ref: str,
            meta: dict, out_name: str):
    """兩份工作簿（起始／目標）與 check 的共同組裝。"""
    def build(with_result: bool):
        wb = new_wb()
        for name, hdr, rows in sheets:
            write_table(wb.create_sheet(name), hdr, rows)
        if with_result:
            ws = wb.create_sheet(out_name)
            ws.append(out_cols)
            for r in expected:
                ws.append(list(r))
        return wb

    known = [name for name, _, _ in sheets]
    check = {"target_sheets": known,
             "new_sheet": {"known_sheets": known, "headers": out_cols,
                           "rows": [list(r) for r in expected]}}
    spec = TaskSpec(task_id, family, instruction, check, ref,
                    meta={**meta, "schema": s["tag"], "n_expected": len(expected)})
    return spec, build(False), build(True)


# 參考解法共用開頭：讀兩張表 + 建欄名對照（沿用 v6 規則 7 的技術）
_HEAD = '''import openpyxl
wb = openpyxl.load_workbook(INPUT_PATH)

def table(name):
    """回傳 (欄名→索引, 資料列)；欄位一律按表頭名稱定位。"""
    ws = wb[name]
    headers = [cell.value for cell in ws[1]]
    pos = {{h: i for i, h in enumerate(headers)}}
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    return pos, rows

old_pos, old_rows = table("{old}")
new_pos, new_rows = table("{new}")
'''

_NORM = '''
def norm(v):
    """鍵正規化：去前後空白、統一大小寫、數字與文字視為相同。"""
    if v is None:
        return None
    s = str(v).strip()
    return s.upper() if s else None
'''


# ======================================================================
# 1. diff_dirty_key —— 鍵不乾淨，指令不提醒
# ======================================================================

def gen_diff_dirty_key(rng: Random, task_id: str):
    for _ in range(40):
        try:
            s = _pick_schema(rng)
            n = rng.randint(16, 26)
            headers, base = _base_rows(s, rng, n)

            old_rows = [list(r) for r in base]
            new_rows = [list(r) for r in base]
            # 真的被刪掉的（＝正解）
            n_del = rng.randint(2, 4)
            del_idx = sorted(rng.sample(range(len(new_rows)), n_del), reverse=True)
            for i in del_idx:
                new_rows.pop(i)
            # 髒鍵污染：只污染其中一張表，另一張保持乾淨（輸出才有唯一寫法）
            dirty_new = rng.random() < 0.5
            target = new_rows if dirty_new else old_rows
            n_dirty = max(4, len(target) // 3)
            for i in rng.sample(range(len(target)), n_dirty):
                k = str(target[i][0])
                style = rng.randint(0, 2)
                target[i][0] = f" {k}" if style == 0 else (f"{k} " if style == 1 else k.lower())

            out_cols = [s["key"], s["name"]]
            new_keys = {str(r[0]).strip().upper() for r in new_rows}
            expected = [[r[0], r[1]] for r in old_rows
                        if str(r[0]).strip().upper() not in new_keys]
            naive_keys = {r[0] for r in new_rows}
            naive = [[r[0], r[1]] for r in old_rows if r[0] not in naive_keys]
            _assert_discriminates(expected, naive,
                                  "diff_dirty_key：不正規化也會對", min_n=2)
            assert len(naive) > len(expected) + 2, "diff_dirty_key：假差異太少，鑑別度不足"
            break
        except AssertionError:
            continue
    else:
        raise RuntimeError("diff_dirty_key：40 次都建不出有鑑別度的資料")

    instruction = _compose(rng, {
        "full": [
            f"比對「{s['old']}」和「{s['new']}」兩張表，用「{s['key']}」當鍵。"
            f"注意兩表的「{s['key']}」有些帶了前後空白或大小寫不同，其實是同一筆，"
            f"比對前要先正規化。把只在「{s['old']}」有的挑出來做成新工作表，"
            f"欄位為「{s['key']}、{s['name']}」。",
        ],
        "mid": [
            f"比對「{s['old']}」和「{s['new']}」（鍵是「{s['key']}」，兩表的寫法不完全一致）。"
            f"列出只在「{s['old']}」有的，新工作表欄位「{s['key']}、{s['name']}」。",

            f"「{s['old']}」和「{s['new']}」以「{s['key']}」對照——資料是人工整理過的。"
            f"把被刪掉的做成新工作表，欄位「{s['key']}、{s['name']}」。",
        ],
        "terse": [
            f"「{s['new']}」比「{s['old']}」少了哪些？做成新工作表，"
            f"欄位「{s['key']}、{s['name']}」。",

            f"比對「{s['old']}」和「{s['new']}」，只在舊的有的挑出來，"
            f"新工作表放「{s['key']}、{s['name']}」。",

            f"「{s['old']}」有、「{s['new']}」沒有的做成新表，欄位「{s['key']}、{s['name']}」。",
        ],
    })

    ref = _HEAD.format(old=s["old"], new=s["new"]) + _NORM + f'''
# 兩表的鍵有前後空白／大小寫差異，直接比對會產生大量假差異——先正規化
nk = new_pos["{s['key']}"]
new_keys = {{norm(r[nk]) for r in new_rows if norm(r[nk]) is not None}}

ok, on = old_pos["{s['key']}"], old_pos["{s['name']}"]
out = wb.create_sheet("差異清單")
out.append({out_cols!r})
for r in old_rows:
    k = norm(r[ok])
    if k is None or k in new_keys:
        continue
    out.append([r[ok], r[on]])
wb.save(OUTPUT_PATH)
'''
    return _finish(task_id, "diff_dirty_key", s, instruction, out_cols, expected,
                   [(s["old"], headers, old_rows), (s["new"], headers, new_rows)],
                   ref, {"variant": "dirty_key", "n_dirty": n_dirty}, "差異清單")


# ======================================================================
# 2. diff_nullkey —— 小計列／註記列（鍵欄空白）不可算成差異
# ======================================================================

def gen_diff_nullkey(rng: Random, task_id: str):
    for _ in range(40):
        try:
            s = _pick_schema(rng)
            n = rng.randint(14, 22)
            headers, base = _base_rows(s, rng, n)

            old_rows = [list(r) for r in base]
            new_rows = [list(r) for r in base]
            # 真的變更（＝正解）
            chg = sorted(rng.sample(range(n), rng.randint(2, 4)))
            for i in chg:
                new_rows[i][2] = new_rows[i][2] + rng.randint(1, 6)
            expected = [[base[i][0], base[i][1], base[i][2], new_rows[i][2]] for i in chg]

            # 兩表都有小計列與註記列，鍵欄空白、數值不同（不處理就會被當成變更）
            def decorate(rows, tag):
                total = sum(r[2] for r in rows)
                out = [list(r) for r in rows]
                out.append([None, "小計", total, None, None])
                out.append([None, f"（{tag}，資料來源系統匯出）", None, None, None])
                return out

            old_rows = decorate(old_rows, s["old"])
            new_rows = decorate(new_rows, s["new"])
            assert old_rows[-2][2] != new_rows[-2][2], "小計相同就考不到空鍵處理"

            out_cols = [s["key"], s["name"], f"舊{s['num']}", f"新{s['num']}"]
            # 天真作法：不跳過空鍵 → 小計列被當成一筆變更
            naive = expected + [[None, "小計", old_rows[-2][2], new_rows[-2][2]]]
            _assert_discriminates(expected, naive, "diff_nullkey：空鍵列沒造成差異", min_n=2)
            break
        except AssertionError:
            continue
    else:
        raise RuntimeError("diff_nullkey：40 次都建不出有鑑別度的資料")

    cols4 = f"{s['key']}、{s['name']}、舊{s['num']}、新{s['num']}"
    instruction = _compose(rng, {
        "full": [
            f"比對「{s['old']}」和「{s['new']}」兩張表的「{s['num']}」，以「{s['key']}」為鍵。"
            f"表格最後有小計列和說明列，那些不是資料、不要算進去。"
            f"把「{s['num']}」有變動的做成新工作表，欄位為「{cols4}」。",
        ],
        "mid": [
            f"「{s['old']}」與「{s['new']}」以「{s['key']}」對照，"
            f"列出「{s['num']}」不一樣的項目（這兩張表是從系統匯出的，格式請自己判斷），"
            f"新工作表欄位：{cols4}。",

            f"請做一張「{s['num']}」異動表（新工作表），比較「{s['old']}」和「{s['new']}」，"
            f"鍵是「{s['key']}」，欄位「{cols4}」，只列真正有變動的資料。",
        ],
        "terse": [
            f"「{s['old']}」和「{s['new']}」的「{s['num']}」有變的挑出來做成新表，欄位「{cols4}」。",

            f"比一下「{s['old']}」跟「{s['new']}」的「{s['num']}」差在哪，"
            f"做成新工作表，欄位「{cols4}」。",

            f"「{s['num']}」異動表做在新工作表，比較「{s['old']}」和「{s['new']}」，欄位「{cols4}」。",
        ],
    })

    ref = _HEAD.format(old=s["old"], new=s["new"]) + f'''
ok, on, oq = old_pos["{s['key']}"], old_pos["{s['name']}"], old_pos["{s['num']}"]
nk, nq = new_pos["{s['key']}"], new_pos["{s['num']}"]

# 小計列／註記列的鍵欄是空的——先過濾掉，否則會被當成一筆差異
new_qty = {{}}
for r in new_rows:
    if r[nk] is None or str(r[nk]).strip() == "":
        continue
    new_qty[str(r[nk]).strip()] = r[nq]

out = wb.create_sheet("數量異動")
out.append({out_cols!r})
for r in old_rows:
    if r[ok] is None or str(r[ok]).strip() == "":
        continue
    k = str(r[ok]).strip()
    if k not in new_qty:
        continue
    if new_qty[k] != r[oq]:
        out.append([r[ok], r[on], r[oq], new_qty[k]])
wb.save(OUTPUT_PATH)
'''
    return _finish(task_id, "diff_nullkey", s, instruction, out_cols, expected,
                   [(s["old"], headers, old_rows), (s["new"], headers, new_rows)],
                   ref, {"variant": "nullkey"}, "數量異動")


# ======================================================================
# 3. diff_dupkey —— 同鍵多列（主/替代料），必須用 tuple 複合鍵
# ======================================================================

_MS_PAIRS = [("M", "S"), ("主", "替"), ("主料", "替代料"), ("正式", "備選")]


def gen_diff_dupkey(rng: Random, task_id: str):
    for _ in range(40):
        try:
            s = _pick_schema(rng)
            main_tag, alt_tag = rng.choice(_MS_PAIRS)
            tag_col = rng.choice(["M/S", "主替", "料件屬性"])
            n = rng.randint(12, 18)
            base_h, base = _base_rows(s, rng, n)
            headers = [s["key"], s["name"], s["num"], tag_col, s["cat"]]

            n_pair = rng.randint(3, 5)
            pair_at = set(rng.sample(range(n), n_pair))
            rows = []
            for i, b in enumerate(base):
                if i in pair_at:
                    rows.append([b[0], b[1], b[2], main_tag, b[3]])
                    rows.append([b[0], b[1], b[2], alt_tag, b[3]])
                else:
                    rows.append([b[0], b[1], b[2], None, b[3]])

            old_rows = [list(r) for r in rows]
            new_rows = [list(r) for r in rows]
            # 只改「主料」那一列 → dict 後寫覆蓋（取到替代料）必然漏掉
            main_idx = [i for i, r in enumerate(new_rows) if r[3] == main_tag]
            chg = sorted(rng.sample(main_idx, min(len(main_idx), rng.randint(2, 3))))
            for i in chg:
                new_rows[i][2] += rng.randint(1, 5)
            expected = [[rows[i][0], rows[i][3], rows[i][2], new_rows[i][2]] for i in chg]

            # 天真作法：dict[鍵] = 數量，後寫覆蓋 → 只看得到替代料（沒變）→ 0 筆
            naive_old, naive_new = {}, {}
            for r in old_rows:
                naive_old[r[0]] = r[2]
            for r in new_rows:
                naive_new[r[0]] = r[2]
            naive = [[k, None, naive_old[k], naive_new[k]]
                     for k in naive_old if naive_new.get(k) != naive_old[k]]
            _assert_discriminates(expected, naive,
                                  "diff_dupkey：dict 覆蓋也會對", min_n=2)
            assert "-" in str(rows[0][0]), "diff_dupkey：鍵必須含分隔符，才能考到 split 拆解的坑"
            break
        except AssertionError:
            continue
    else:
        raise RuntimeError("diff_dupkey：40 次都建不出有鑑別度的資料")

    out_cols = [s["key"], tag_col, f"舊{s['num']}", f"新{s['num']}"]
    cols4 = f"{s['key']}、{tag_col}、舊{s['num']}、新{s['num']}"
    instruction = _compose(rng, {
        "full": [
            f"「{s['old']}」和「{s['new']}」兩張表裡，同一個「{s['key']}」可能有多列——"
            f"「{tag_col}」欄標「{main_tag}」的是主要項目、標「{alt_tag}」的是替代項目，"
            f"兩者要分開比對，不可以合併成一筆。請把「{s['num']}」有變動的做成新工作表，"
            f"欄位為「{cols4}」。",

            f"以「{s['key']}」加上「{tag_col}」兩欄合起來當鍵，比對「{s['old']}」和「{s['new']}」，"
            f"列出「{s['num']}」不同的項目，做成新工作表，欄位「{cols4}」。",
        ],
        "mid": [
            f"比對「{s['old']}」與「{s['new']}」的「{s['num']}」——注意「{s['key']}」不是唯一的，"
            f"要看清楚才不會比錯。新工作表欄位：{cols4}。",

            f"「{s['old']}」和「{s['new']}」比「{s['num']}」，"
            f"鍵請自己判斷（單看「{s['key']}」會有多筆），新工作表欄位「{cols4}」。",
        ],
        "terse": [
            f"比對「{s['old']}」和「{s['new']}」的「{s['num']}」，"
            f"有變的做成新工作表，欄位「{cols4}」。",

            f"「{s['num']}」有異動的挑出來做成新表，比較「{s['old']}」跟「{s['new']}」，欄位「{cols4}」。",

            f"「{s['old']}」與「{s['new']}」的「{s['num']}」差異做在新工作表，欄位「{cols4}」。",
        ],
    })

    ref = _HEAD.format(old=s["old"], new=s["new"]) + f'''
def index_by_pair(pos, rows):
    """鍵不唯一——用 tuple ({s['key']}, {tag_col}) 當複合鍵。
    不可以把兩欄黏成字串再 split 拆回來：{s['key']} 本身就含分隔符。"""
    k, t, q = pos["{s['key']}"], pos["{tag_col}"], pos["{s['num']}"]
    out = {{}}
    for r in rows:
        if r[k] is None:
            continue
        out[(str(r[k]).strip(), r[t])] = r[q]
    return out

old_qty = index_by_pair(old_pos, old_rows)
new_qty = index_by_pair(new_pos, new_rows)

out = wb.create_sheet("數量異動")
out.append({out_cols!r})
for key, tag in old_qty:
    if (key, tag) not in new_qty:
        continue
    if new_qty[(key, tag)] != old_qty[(key, tag)]:
        out.append([key, tag, old_qty[(key, tag)], new_qty[(key, tag)]])
wb.save(OUTPUT_PATH)
'''
    return _finish(task_id, "diff_dupkey", s, instruction, out_cols, expected,
                   [(s["old"], headers, old_rows), (s["new"], headers, new_rows)],
                   ref, {"variant": "dupkey", "tag_col": tag_col,
                         "main_tag": main_tag, "alt_tag": alt_tag}, "數量異動")


# ======================================================================
# 4. diff_carry_cols —— 差異表要帶出指令指定的其他欄位
# ======================================================================

def gen_diff_carry_cols(rng: Random, task_id: str):
    for _ in range(40):
        try:
            s = _pick_schema(rng)
            n = rng.randint(16, 24)
            headers, base = _base_rows(s, rng, n)
            old_rows = [list(r) for r in base]
            new_rows = [list(r) for r in base]

            n_add = rng.randint(2, 3)
            n_del = rng.randint(2, 3)
            for i in sorted(rng.sample(range(len(new_rows)), n_del), reverse=True):
                new_rows.pop(i)
            extra_h, extra = _base_rows(s, rng, n_add)
            existing = {r[0] for r in old_rows}
            assert not (existing & {r[0] for r in extra}), "新增料的鍵撞到既有鍵"
            new_rows += [list(r) for r in extra]

            new_keys = {r[0] for r in new_rows}
            old_keys = {r[0] for r in old_rows}
            # 「欄位與原表相同、最前面加一欄」比逐一點名五個欄位自然得多，指令也短得多
            out_cols = ["差異類型"] + headers
            expected = ([["新增"] + list(r) for r in new_rows if r[0] not in old_keys]
                        + [["刪除"] + list(r) for r in old_rows if r[0] not in new_keys])
            # 天真作法：只帶鍵，其他欄位留空
            naive = [[e[0], e[1]] + [None] * (len(headers) - 1) for e in expected]
            _assert_discriminates(expected, naive,
                                  "diff_carry_cols：其他欄位剛好都是空的", min_n=4)
            break
        except AssertionError:
            continue
    else:
        raise RuntimeError("diff_carry_cols：40 次都建不出有鑑別度的資料")

    instruction = _compose(rng, {
        "full": [
            f"比對「{s['old']}」和「{s['new']}」兩張表（鍵是「{s['key']}」），"
            f"做一張新工作表列出新增與刪除的項目，欄位與原表相同、最前面加一欄「差異類型」；"
            f"差異類型填「新增」（只在「{s['new']}」有）或「刪除」（只在「{s['old']}」有），"
            f"先列新增再列刪除。其餘欄位要從原表整列帶出來，不可以留空。",
        ],
        "mid": [
            f"請以「{s['key']}」對照「{s['old']}」與「{s['new']}」，"
            f"新增與刪除的做成新工作表，欄位與原表相同、前面加「差異類型」，"
            f"新增排前面，資料照原表帶。",
        ],
        "terse": [
            f"「{s['old']}」和「{s['new']}」比一下，新增和刪除的做成新工作表，"
            f"欄位與原表相同、最前面加一欄「差異類型」，新增排前面。",

            f"兩張表的增刪做成新表，欄位照原表、第一欄放「差異類型」，先新增後刪除。",

            f"做一張增刪對照表（新工作表），欄位與原表相同、前面多一欄「差異類型」，新增在前。",
        ],
    })

    ref = _HEAD.format(old=s["old"], new=s["new"]) + f'''
COLS = {headers!r}          # 欄位與原表相同
# 索引要存「整列」，不能只存鍵——否則其他欄位會全部變成空值
old_by = {{}}
for r in old_rows:
    if r[old_pos["{s['key']}"]] is not None:
        old_by[r[old_pos["{s['key']}"]]] = r
new_by = {{}}
for r in new_rows:
    if r[new_pos["{s['key']}"]] is not None:
        new_by[r[new_pos["{s['key']}"]]] = r

out = wb.create_sheet("差異清單")
out.append(["差異類型"] + COLS)
for k, r in new_by.items():
    if k not in old_by:
        out.append(["新增"] + [r[new_pos[c]] for c in COLS])
for k, r in old_by.items():
    if k not in new_by:
        out.append(["刪除"] + [r[old_pos[c]] for c in COLS])
wb.save(OUTPUT_PATH)
'''
    return _finish(task_id, "diff_carry_cols", s, instruction, out_cols, expected,
                   [(s["old"], headers, old_rows), (s["new"], headers, new_rows)],
                   ref, {"variant": "carry_cols"}, "差異清單")


# ======================================================================
# 5. diff_multicol —— 指令列出 N 個欄位就要有 N 個比較分支
# ======================================================================

def gen_diff_multicol(rng: Random, task_id: str):
    for _ in range(40):
        try:
            s = _pick_schema(rng)
            n = rng.randint(16, 24)
            headers, base = _base_rows(s, rng, n)
            old_rows = [list(r) for r in base]
            new_rows = [list(r) for r in base]

            # 比對「鍵以外的所有欄位」——真實使用者不會逐一點名，只會說「哪些欄位有變」
            targets = [(s["num"], 2), (s["cat"], 3), (s["unit"], 4)]
            picked = rng.sample(range(n), 3 * rng.randint(1, 2))
            expected, used = [], set()
            for j, (cname, ci) in enumerate(targets):
                for i in picked[j::3]:
                    if i in used:
                        continue
                    used.add(i)
                    old_v = new_rows[i][ci]
                    if ci == 2:
                        new_v = old_v + rng.randint(1, 6)
                    else:
                        pool = [x for x in (s["cats"] if ci == 3 else s["units"]) if x != old_v]
                        new_v = rng.choice(pool)
                    new_rows[i][ci] = new_v
                    expected.append([base[i][0], cname, old_v, new_v])
            # 輸出依「原表列序，同列依原表欄位順序」排
            col_order = {c: headers.index(c) for c, _ in targets}
            key_row = {base[i][0]: i for i in range(n)}
            expected.sort(key=lambda e: (key_row[e[0]], col_order[e[1]]))

            for cname, _ in targets:
                assert any(e[1] == cname for e in expected), f"「{cname}」沒有任何變更"
            # 天真作法：只比第一個欄位（多半是數量）
            naive = [e for e in expected if e[1] == targets[0][0]]
            _assert_discriminates(expected, naive,
                                  "diff_multicol：只比一欄也會對", min_n=3)
            break
        except AssertionError:
            continue
    else:
        raise RuntimeError("diff_multicol：40 次都建不出有鑑別度的資料")

    out_cols = [s["key"], "變更欄位", "舊值", "新值"]
    cols4 = f"{s['key']}、變更欄位、舊值、新值"
    instruction = _compose(rng, {
        "full": [
            f"比對「{s['old']}」和「{s['new']}」兩張表，鍵是「{s['key']}」。"
            f"「{s['key']}」以外的欄位全部都要比，逐欄列出變更做成新工作表，"
            f"欄位為「{cols4}」，一列只記一個欄位的變更；"
            f"依原表的列順序排，同一列有多個欄位變更時依欄位在原表的順序。",
        ],
        "mid": [
            f"請比對「{s['old']}」與「{s['new']}」（以「{s['key']}」為鍵），"
            f"其他欄位每個有變的都記一列，新工作表欄位「{cols4}」，"
            f"依原表列序、同列依原表欄位順序。",
        ],
        "terse": [
            f"「{s['old']}」和「{s['new']}」哪些欄位有變？逐欄列出來做成新表，"
            f"欄位「{cols4}」，照原表順序。",

            f"比對兩張表，變更明細做在新工作表，欄位「{cols4}」，依原表列序排。",

            f"「{s['old']}」跟「{s['new']}」的變更逐欄列出，新工作表欄位「{cols4}」。",
        ],
    })

    ref = _HEAD.format(old=s["old"], new=s["new"]) + f'''
# 指令點名三個欄位，就要有三個比較分支——用迴圈跑完清單，不要漏掉任何一欄
KEY = "{s['key']}"
# 指令說「{s['key']} 以外的欄位全部都要比」——用表頭算出清單，不要手寫欄名
COMPARE = [h for h in old_pos if h != KEY]

new_by = {{}}
for r in new_rows:
    if r[new_pos[KEY]] is not None:
        new_by[r[new_pos[KEY]]] = r

out = wb.create_sheet("變更明細")
out.append({out_cols!r})
for r in old_rows:
    k = r[old_pos[KEY]]
    if k is None or k not in new_by:
        continue
    nr = new_by[k]
    for col in COMPARE:
        ov, nv = r[old_pos[col]], nr[new_pos[col]]
        if ov != nv:
            out.append([k, col, ov, nv])
wb.save(OUTPUT_PATH)
'''
    return _finish(task_id, "diff_multicol", s, instruction, out_cols, expected,
                   [(s["old"], headers, old_rows), (s["new"], headers, new_rows)],
                   ref, {"variant": "multicol"}, "變更明細")


V7_FAMILIES = {
    "diff_dirty_key": gen_diff_dirty_key,
    "diff_nullkey": gen_diff_nullkey,
    "diff_dupkey": gen_diff_dupkey,
    "diff_carry_cols": gen_diff_carry_cols,
    "diff_multicol": gen_diff_multicol,
}

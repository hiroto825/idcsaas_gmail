"""有望技術DBの試験収集結果（4マス分のJSON）をExcel/CSVにまとめる。

usage: python3 scripts/build_tech_db.py <raw_dir> <output_dir>
"""
import csv
import json
import sys
from collections import Counter
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

QUADRANTS = {
    "jp_univ": "日本×大学・公的研究機関",
    "jp_corp": "日本×企業",
    "global_univ": "グローバル×大学・公的研究機関",
    "global_corp": "グローバル×企業",
}
CUTOFF = "2023-10-30"
COLUMNS = [
    ("No", 6), ("マス", 22), ("ソース", 18), ("ソース種別", 14), ("組織", 26), ("組織種別", 12), ("国", 8),
    ("日付", 12), ("日付種別", 9), ("分野", 14), ("タイトル", 50), ("概要", 60), ("成熟度の根拠", 45),
    ("成熟度確度", 9), ("法人化", 11), ("判定", 10), ("メモ", 30), ("URL", 50),
]
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def norm(url):
    return (url or "").strip().rstrip("/").lower()


def judge(it):
    """3条件で機械的に判定する。成熟度は収集時の確度を使う。"""
    if (it.get("date") or "") < CUTOFF[: len(it.get("date") or "")]:
        return "除外:3年超"
    if it.get("incorporated") == "法人化済み":
        return "要確認:法人化"
    if not it.get("date"):
        return "要確認:日付"
    return "対象"


def load(raw_dir):
    items, counts, seen = [], [], set()
    for q in QUADRANTS:
        f = Path(raw_dir) / f"{q}.json"
        if not f.exists():
            continue
        d = json.loads(f.read_text(encoding="utf-8"))
        for c in d.get("source_counts", []):
            counts.append({"quadrant": q, **c})
        for it in d.get("items", []):
            key = norm(it.get("url")) or it.get("title", "")
            if key in seen:
                continue
            seen.add(key)
            it["quadrant"] = q
            it["judgement"] = judge(it)
            items.append(it)
    return items, counts


def style_header(ws):
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    ws.freeze_panes = "A2"


def row_of(n, it):
    return [
        n, QUADRANTS[it["quadrant"]], it.get("source", ""), it.get("source_type", ""), it.get("org", ""),
        it.get("org_type", ""), it.get("country", ""), it.get("date") or "", it.get("date_type", ""),
        it.get("field", ""), it.get("title", ""), it.get("summary_ja", ""), it.get("maturity_evidence", ""),
        it.get("maturity_confidence", ""), it.get("incorporated", ""), it["judgement"], it.get("notes", ""),
        it.get("url", ""),
    ]


def write_list(ws, rows):
    ws.append([c for c, _ in COLUMNS])
    for n, it in enumerate(rows, 1):
        ws.append(row_of(n, it))
        if it.get("url"):
            cell = ws.cell(row=n + 1, column=len(COLUMNS))
            cell.hyperlink = it["url"]
            cell.font = Font(color="0563C1", underline="single")
    for idx, (_, w) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    for r in ws.iter_rows(min_row=2):
        for cell in r[10:13]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    style_header(ws)
    ws.auto_filter.ref = ws.dimensions


def write_counts(ws, counts, items):
    collected = Counter((i["quadrant"], i.get("source", "")) for i in items)
    ws.append(["マス", "ソース", "3年件数", "根拠", "算定方法", "試験収集件数", "URL"])
    for c in counts:
        ws.append([QUADRANTS[c["quadrant"]], c.get("source", ""), c.get("count_3y"), c.get("count_basis", ""),
                   c.get("method", ""), collected.get((c["quadrant"], c.get("source", "")), 0), c.get("url", "")])
    for col, w in zip("ABCDEFG", (22, 28, 10, 8, 70, 12, 45)):
        ws.column_dimensions[col].width = w
    for r in ws.iter_rows(min_row=2):
        r[4].alignment = Alignment(wrap_text=True, vertical="top")
    style_header(ws)


def write_summary(ws, items):
    judgements = ["対象", "要確認:法人化", "要確認:日付", "除外:3年超"]
    ws.append(["マス"] + judgements + ["計"])
    for q, label in QUADRANTS.items():
        c = Counter(i["judgement"] for i in items if i["quadrant"] == q)
        ws.append([label] + [c[j] for j in judgements] + [sum(c.values())])
    c = Counter(i["judgement"] for i in items)
    ws.append(["合計"] + [c[j] for j in judgements] + [len(items)])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 30
    for col in "BCDEF":
        ws.column_dimensions[col].width = 14
    style_header(ws)


def main(raw_dir, output_dir):
    items, counts = load(raw_dir)
    items.sort(key=lambda i: (list(QUADRANTS).index(i["quadrant"]), i.get("source", ""), i.get("date") or ""))

    wb = Workbook()
    ws = wb.active
    ws.title = "集計"
    write_summary(ws, items)
    write_list(wb.create_sheet("技術リスト"), [i for i in items if i["judgement"] != "除外:3年超"])
    write_counts(wb.create_sheet("ソース別件数"), counts, items)
    write_list(wb.create_sheet("除外"), [i for i in items if i["judgement"] == "除外:3年超"])

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    stamp = date.today().strftime("%Y-%m")
    wb.save(out / f"tech_db_pilot_{stamp}.xlsx")
    with open(out / f"tech_db_pilot_{stamp}.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([c for c, _ in COLUMNS])
        for n, it in enumerate([i for i in items if i["judgement"] != "除外:3年超"], 1):
            w.writerow(row_of(n, it))

    print(f"items={len(items)}", dict(Counter(i["judgement"] for i in items)))
    for q in QUADRANTS:
        print(q, dict(Counter(i.get("source", "") for i in items if i["quadrant"] == q)))


if __name__ == "__main__":
    main(*sys.argv[1:3])

"""
핫딜 수집 결과(JSON/JSONL)를 보기 편한 엑셀(.xlsx)로 변환합니다.

사용 예:
    python export_excel.py deals.json                 # → deals.xlsx
    python export_excel.py deals.json -o 오늘핫딜.xlsx

hotdeal_scraper.py에서 -o deals.xlsx로 저장해도 이 모듈을 사용합니다.

시트 구성:
- 핫딜: 전체 목록. 최신 글이 위. 제목을 누르면 게시글로 이동. 필터·첫 줄 고정 적용.
- 요약: 사이트별, 분류별, 판매처별 건수.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

KST = timezone(timedelta(hours=9))

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SITE_FILL = {"퀘이사존": PatternFill("solid", fgColor="FFF2CC"), "아카라이브": PatternFill("solid", fgColor="DDEBF7")}
NEW_FILL = PatternFill("solid", fgColor="E2EFDA")
DIM_FONT = Font(color="A6A6A6")
DIM_LINK_FONT = Font(color="A6A6A6", underline="single")
LINK_FONT = Font(color="0563C1", underline="single")
THIN = Border(bottom=Side(style="thin", color="D9D9D9"))

# (헤더, 열 너비, 헤더 메모)
COLUMNS: list[tuple[str, float, str | None]] = [
    ("사이트", 11, None),
    ("신규", 7, "이전 결과(before.json)에 없던 글입니다."),
    ("상태", 7, "종료된 딜만 '종료'로 표시됩니다."),
    ("분류", 13, None),
    ("제목", 62, "누르면 게시글이 열립니다."),
    ("가격(원)", 12, "원화 가격만 숫자로 변환했습니다. 외화는 '가격 표기' 열을 보세요."),
    ("가격 표기", 12, None),
    ("판매처", 12, None),
    ("배송", 10, None),
    ("댓글", 7, None),
    ("조회", 8, "퀘이사존 값은 페이지 로드 전 값이라 실제와 다를 수 있습니다."),
    ("추천", 7, None),
    ("작성 시각", 17, "한국 시각. 퀘이사존은 'N시간 전' 표기를 수집 시각 기준으로 환산한 추정값입니다."),
    ("작성 표기", 11, "사이트에 표시된 원래 값"),
    ("글번호", 11, None),
]


def krw_price(text: str) -> int | None:
    """'￦1,611,700', '12,620원' → 정수. 외화($, ¥ 등)나 숫자가 없으면 None."""
    if not text or re.search(r"[$¥€£]|달러|엔|위안", text):
        return None
    if "￦" not in text and "₩" not in text and "원" not in text:
        return None
    m = re.search(r"\d[\d,]*", text)
    return int(m.group().replace(",", "")) if m else None


def delivery_label(text: str) -> str:
    t = re.sub(r"^배송비\s*", "", text or "").strip()
    return "무료" if t in ("0", "무배", "무료") else t


def posted_time(posted: str, collected_at: str) -> datetime | None:
    """ISO 시각(UTC 포함) 또는 'N분/시간/일 전' 표기를 한국 시각(naive)으로 변환."""
    if not posted:
        return None
    try:
        dt = datetime.fromisoformat(posted.replace("Z", "+00:00"))
        if dt.tzinfo:
            dt = dt.astimezone(KST)
        return dt.replace(tzinfo=None, microsecond=0)
    except ValueError:
        pass
    m = re.fullmatch(r"\s*(\d+)\s*(초|분|시간|일)\s*전\s*", posted)
    if m and collected_at:
        unit = {"초": "seconds", "분": "minutes", "시간": "hours", "일": "days"}[m.group(2)]
        base = datetime.fromisoformat(collected_at)
        return (base - timedelta(**{unit: int(m.group(1))})).replace(second=0, microsecond=0)
    return None


def clean_title(title: str, comments: Any) -> str:
    """이전 버전 수집 결과의 제목 끝 댓글 수 표기('... [5]')를 제거."""
    m = re.search(r"\s*\[(\d+)\]\s*$", title or "")
    if m and isinstance(comments, int) and int(m.group(1)) == comments:
        return title[:m.start()].rstrip()
    return title or ""


def load_items(path: Path) -> list[dict[str, Any]]:
    text = path.read_text(encoding="utf-8-sig")
    if path.suffix.lower() == ".jsonl":
        return [json.loads(line) for line in text.splitlines() if line.strip()]
    return json.loads(text)


def write_xlsx(items: list[dict[str, Any]], out: Path) -> Path:
    rows = []
    for it in items:
        rows.append({
            "it": it,
            "posted_dt": posted_time(str(it.get("posted") or ""), str(it.get("collected_at") or "")),
        })
    rows.sort(key=lambda r: r["posted_dt"] or datetime.min, reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "핫딜"

    for col, (name, width, note) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=1, column=col, value=name)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        if note:
            c.comment = Comment(note, "hotdeal")
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22

    for r_idx, row in enumerate(rows, start=2):
        it = row["it"]
        done = it.get("status") == "종료" or "블라인드" in str(it.get("title", ""))
        values = [
            it.get("site", ""),
            "신규" if it.get("is_new") else "",
            it.get("status", ""),
            it.get("category", ""),
            clean_title(str(it.get("title") or ""), it.get("comments")),
            krw_price(str(it.get("price") or "")),
            it.get("price", ""),
            it.get("store", ""),
            delivery_label(str(it.get("delivery") or "")),
            it.get("comments"),
            it.get("views"),
            it.get("votes"),
            row["posted_dt"],
            it.get("posted", ""),
            int(it["post_id"]) if str(it.get("post_id", "")).isdigit() else it.get("post_id", ""),
        ]
        for col, v in enumerate(values, start=1):
            c = ws.cell(row=r_idx, column=col, value=v)
            c.border = THIN
            c.alignment = Alignment(vertical="center")
            if done:
                c.font = DIM_FONT
        ws.cell(row=r_idx, column=1).fill = SITE_FILL.get(it.get("site", ""), PatternFill())
        if it.get("is_new"):
            new_cell = ws.cell(row=r_idx, column=2)
            new_cell.fill = NEW_FILL
            new_cell.font = Font(bold=True, color="375623")

        title = ws.cell(row=r_idx, column=5)
        if it.get("url"):
            title.hyperlink = it["url"]
            title.font = DIM_LINK_FONT if done else LINK_FONT
        ws.cell(row=r_idx, column=6).number_format = "#,##0"
        ws.cell(row=r_idx, column=13).number_format = "yyyy-mm-dd hh:mm"
        for col in (2, 3, 10, 11, 12):
            ws.cell(row=r_idx, column=col).alignment = Alignment(horizontal="center", vertical="center")

    last_col = get_column_letter(len(COLUMNS))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{last_col}{max(1, len(rows)) + 1}"

    _write_summary(wb.create_sheet("요약"), items)

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def _write_summary(ws, items: list[dict[str, Any]]) -> None:
    bold = Font(bold=True)
    collected = sorted({str(it.get("collected_at", "")) for it in items if it.get("collected_at")})
    ws["A1"] = "핫딜 수집 요약"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "수집 시각"
    ws["B2"] = f"{collected[0]} ~ {collected[-1]}" if len(collected) > 1 else (collected[0] if collected else "")
    ws["A3"] = "전체 건수"
    ws["B3"] = len(items)
    ws["A4"] = "새 글 (이전 결과에 없던 글)"
    ws["B4"] = sum(1 for it in items if it.get("is_new"))
    ws["A5"] = "종료된 딜"
    ws["B5"] = sum(1 for it in items if it.get("status") == "종료")
    for cell in ("A2", "A3", "A4", "A5"):
        ws[cell].font = bold

    row = 7

    def table(title: str, counter: Counter, top: int | None = None) -> None:
        nonlocal row
        ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=12)
        row += 1
        for col, name in enumerate(("항목", "건수"), start=1):
            c = ws.cell(row=row, column=col, value=name)
            c.fill, c.font = HEADER_FILL, HEADER_FONT
        row += 1
        for name, n in counter.most_common(top):
            ws.cell(row=row, column=1, value=name or "(없음)")
            ws.cell(row=row, column=2, value=n)
            row += 1
        row += 1

    table("사이트별", Counter(it.get("site", "") for it in items))
    for site in sorted({it.get("site", "") for it in items}):
        table(f"{site} 분류별", Counter(it.get("category", "") for it in items if it.get("site") == site))
    table("판매처별 (상위 15)", Counter(it.get("store", "") for it in items), top=15)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 44


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="핫딜 수집 결과를 엑셀로 변환")
    ap.add_argument("input", help="hotdeal_scraper.py가 저장한 .json 또는 .jsonl 파일")
    ap.add_argument("-o", "--output", help="출력 .xlsx 파일 (기본: 입력 파일명.xlsx)")
    args = ap.parse_args(argv)

    src = Path(args.input)
    out = Path(args.output) if args.output else src.with_suffix(".xlsx")
    items = load_items(src)
    write_xlsx(items, out)
    print(f"{len(items)}건 → {out.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

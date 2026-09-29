# -*- coding: utf-8 -*-
"""갱신보험료 납입 엑셀 파일 인쇄용 서식 정리 스크립트.

지정한 폴더 안의 .xlsx 파일을 모두 열어서
  - 1행(제목행): 밝은 색 배경 + 굵은 글씨 + 가운데 정렬 + 흰 테두리 느낌의 구분선
  - 나머지 셀: 기본 가운데 정렬(가로/세로)
  - 금액으로 보이는 숫자: 오른쪽 정렬 유지
  - 얇은 회색 테두리, 열 너비 자동, 틀 고정, 인쇄 설정(가로 1페이지 맞춤 + 제목행 반복)
을 적용하고 원본을 덮어씁니다. (--backup 으로 원본 백업 가능)

값이나 수식은 바꾸지 않습니다. 서식만 손댑니다.

사용법:
    python format_renewal_xlsx.py "C:\\Users\\dingg\\OneDrive\\바탕 화면\\갱신보험료 납입"
    python format_renewal_xlsx.py . --backup
    python format_renewal_xlsx.py . --comma          # General 서식 금액에 1,000 단위 콤마 추가
    python format_renewal_xlsx.py . --orientation portrait
"""

from __future__ import annotations

import argparse
import datetime as _dt
import re
import shutil
import sys
import unicodedata
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties
except ImportError:  # pragma: no cover
    sys.exit("openpyxl 이 필요합니다.  먼저 실행:  pip install openpyxl")


# ---------------------------------------------------------------- 서식 상수
HEADER_FILL = PatternFill("solid", fgColor="FFD9E8F7")   # 밝은 하늘색
HEADER_FONT = Font(name="맑은 고딕", size=10, bold=True, color="FF1F3864")
BODY_FONT = Font(name="맑은 고딕", size=10, color="FF262626")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=False, indent=1)

_LINE = Side(style="thin", color="FFBFCEE0")
BORDER = Border(left=_LINE, right=_LINE, top=_LINE, bottom=_LINE)

HEADER_ROW_HEIGHT = 30
BODY_ROW_HEIGHT = 20
MIN_COL_WIDTH = 8
MAX_COL_WIDTH = 42

# 이 값 이상인 숫자는 "금액"으로 보고 오른쪽 정렬. 회차·연령 같은 작은 수는 가운데.
AMOUNT_THRESHOLD = 1000

# 문자열로 들어간 금액 (예: "1,234,000", "USD 5,000", "$1,200.50", "3,000원")
_TEXT_AMOUNT = re.compile(
    r"^[\s]*(?:USD|KRW|HKD|SGD|JPY|EUR|\$|₩|￦|€|¥)?[\s]*"
    r"\d{1,3}(?:,\d{3})+(?:\.\d+)?[\s]*(?:원|달러|USD|KRW)?[\s]*$"
)

# 날짜/시간 서식 판별용
_DATEISH = re.compile(r"(?<!\\)[ymdhs]", re.IGNORECASE)

# 제목에 이 말이 들어간 열은 "금액 열"로 보고 열 전체를 오른쪽 정렬
_MONEY_HEADER = (
    "금액", "보험료", "납입액", "원금", "합계", "잔액", "대출", "환산",
    "premium", "amount", "levy", "loan", "balance", "total", "fee",
    "usd", "krw", "hkd", "sgd", "jpy", "eur",
)

# 제목에 이 말이 들어간 열은 숫자여도 식별번호로 보고 가운데 정렬
_ID_HEADER = (
    "번호", "증권", "코드", "연락처", "회차", "차년", "연령", "나이", "건수",
    "no.", "no", "policy", "code", "id", "phone", "count", "age", "year",
)


def _is_dateish(cell) -> bool:
    if isinstance(cell.value, (_dt.datetime, _dt.date, _dt.time)):
        return True
    fmt = (cell.number_format or "").lower()
    if fmt in ("general", ""):
        return False
    stripped = re.sub(r'"[^"]*"', "", fmt)          # 따옴표 안 리터럴 제거
    return bool(_DATEISH.search(stripped))


def _is_amount(cell) -> bool:
    """오른쪽 정렬 대상인지 판정."""
    v = cell.value
    if isinstance(v, bool) or v is None:
        return False
    if _is_dateish(cell):
        return False
    if isinstance(v, (int, float)):
        fmt = (cell.number_format or "").lower()
        if "#,##" in fmt or "0.00" in fmt or any(s in fmt for s in ("₩", "$", "usd", "krw", "€", "¥")):
            return True
        if "%" in fmt:
            return False
        return abs(v) >= AMOUNT_THRESHOLD
    if isinstance(v, str):
        return bool(_TEXT_AMOUNT.match(v))
    return False


def _display_width(value) -> int:
    """한글·한자는 2칸으로 계산한 표시 폭 (여러 줄이면 가장 긴 줄 기준)."""
    if value is None:
        return 0
    if isinstance(value, _dt.datetime):
        text = value.strftime("%Y-%m-%d") if (value.hour, value.minute, value.second) == (0, 0, 0) \
            else value.strftime("%Y-%m-%d %H:%M")
    elif isinstance(value, (_dt.date, _dt.time)):
        text = str(value)
    elif isinstance(value, bool):
        text = str(value)
    elif isinstance(value, int):
        text = f"{value:,}"
    elif isinstance(value, float):
        text = f"{int(value):,}" if value == int(value) else f"{value:,.2f}"
    else:
        text = str(value)
    widest = 0
    for line in text.split("\n"):
        w = sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in line)
        widest = max(widest, w)
    return widest


def _classify_columns(ws, last_col: int, right_cols: set, center_cols: set) -> dict:
    """열 번호 -> "right" | "center" | None(셀 단위 판정) 매핑.

    1) 사용자가 --right-cols / --center-cols 로 지정한 열이 최우선
    2) 제목이 식별번호성이면 가운데 (숫자형 증권번호가 금액으로 오인되는 것 방지)
    3) 제목이 금액성이고 본문 과반이 숫자면 열 전체를 오른쪽 (값이 0이어도 유지)
    """
    result: dict[int, str | None] = {}
    for col in range(1, last_col + 1):
        if col in right_cols:
            result[col] = "right"
            continue
        if col in center_cols:
            result[col] = "center"
            continue

        title = str(ws.cell(row=1, column=col).value or "").strip().lower()
        if any(k in title for k in _ID_HEADER):
            result[col] = "center"
            continue

        if any(k in title for k in _MONEY_HEADER):
            filled = numeric = 0
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=col).value
                if v is None or v == "":
                    continue
                filled += 1
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    numeric += 1
                elif isinstance(v, str) and _TEXT_AMOUNT.match(v):
                    numeric += 1
            if filled and numeric * 2 >= filled:
                result[col] = "right"
                continue

        result[col] = None
    return result


def _merged_anchor_map(ws) -> dict:
    """병합 셀의 좌상단 좌표 집합 (병합 셀은 좌상단에만 서식을 줘도 충분)."""
    return {(rng.min_row, rng.min_col) for rng in ws.merged_cells.ranges}


def format_sheet(
    ws,
    *,
    orientation: str,
    add_comma: bool,
    right_cols: set | None = None,
    center_cols: set | None = None,
) -> None:
    if ws.max_row < 1 or ws.sheet_state != "visible":
        return

    col_kind = _classify_columns(
        ws, ws.max_column, right_cols or set(), center_cols or set()
    )

    anchors = _merged_anchor_map(ws)
    merged_cells = {
        (r, c)
        for rng in ws.merged_cells.ranges
        for r in range(rng.min_row, rng.max_row + 1)
        for c in range(rng.min_col, rng.max_col + 1)
    } - anchors

    money_cols = {c for c, kind in col_kind.items() if kind == "right"}

    widths: dict[int, int] = {}

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            coord = (cell.row, cell.column)
            if coord in merged_cells:
                continue

            if cell.row == 1:
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = CENTER
            else:
                cell.font = BODY_FONT
                kind = col_kind.get(cell.column)
                if kind == "right":
                    cell.alignment = RIGHT
                elif kind == "center":
                    cell.alignment = CENTER
                else:
                    cell.alignment = RIGHT if _is_amount(cell) else CENTER
                if (
                    add_comma
                    and cell.column in money_cols
                    and isinstance(cell.value, (int, float))
                    and not isinstance(cell.value, bool)
                    and (cell.number_format or "General") == "General"
                ):
                    cell.number_format = "#,##0"

            cell.border = BORDER

            w = _display_width(cell.value)
            if w and w > widths.get(cell.column, 0):
                widths[cell.column] = w

    # 금액 열 표시 서식 통일.
    # 열 안에 서식이 섞여 있으면(#,##0.00 과 General 혼재) 우세한 쪽으로 맞추고,
    # 명시 서식이 아예 없는 금액 열(예: 값이 전부 0)은 같은 시트의 다른 금액 열을 따라간다.
    per_col_counts: dict[int, dict[str, int]] = {}
    sheet_counts: dict[str, int] = {}
    for col in money_cols:
        counts: dict[str, int] = {}
        for r in range(2, ws.max_row + 1):
            c = ws.cell(row=r, column=col)
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                fmt = c.number_format or "General"
                if fmt != "General":
                    counts[fmt] = counts.get(fmt, 0) + 1
                    sheet_counts[fmt] = sheet_counts.get(fmt, 0) + 1
        per_col_counts[col] = counts

    sheet_dominant = max(sheet_counts, key=sheet_counts.get) if sheet_counts else None
    for col in money_cols:
        counts = per_col_counts[col]
        dominant = max(counts, key=counts.get) if counts else sheet_dominant
        if not dominant:
            continue
        for r in range(2, ws.max_row + 1):
            c = ws.cell(row=r, column=col)
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                c.number_format = dominant

    # 본문에 값이 있는 열 (제목만 있는 기입용 빈 칸과 구분)
    filled_cols = {
        c.column
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column)
        for c in row
        if c.value not in (None, "")
    }

    # 열 너비 (병합 셀 제목이 과하게 넓히는 것을 막기 위해 상한 적용)
    for col, w in widths.items():
        letter = get_column_letter(col)
        new_width = min(MAX_COL_WIDTH, max(MIN_COL_WIDTH, w + 3))
        if col not in filled_cols:
            # 손으로 적는 빈 칸은 좁히지 않는다
            current = ws.column_dimensions[letter].width
            if current:
                new_width = max(new_width, min(MAX_COL_WIDTH, current))
        ws.column_dimensions[letter].width = new_width

    # 행 높이
    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT
    for r in range(2, ws.max_row + 1):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = BODY_ROW_HEIGHT

    # 화면 / 인쇄
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = orientation
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_title_rows = "1:1"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.page_margins.header = ws.page_margins.footer = 0.2


def format_file(
    path: Path,
    *,
    orientation: str,
    add_comma: bool,
    backup: bool,
    right_cols: set | None = None,
    center_cols: set | None = None,
) -> str:
    wb = load_workbook(path)          # 수식 보존 (data_only=False)
    for ws in wb.worksheets:
        format_sheet(
            ws,
            orientation=orientation,
            add_comma=add_comma,
            right_cols=right_cols,
            center_cols=center_cols,
        )
    if backup:
        shutil.copy2(path, path.with_suffix(path.suffix + ".bak"))
    wb.save(path)
    return f"{path.name}  ({len(wb.worksheets)}개 시트)"


def _parse_cols(spec: str) -> set:
    """ "G,I,J" 또는 "7,9,10" 을 열 번호 집합으로 변환."""
    cols = set()
    for token in spec.replace(" ", "").split(","):
        if not token:
            continue
        if token.isdigit():
            cols.add(int(token))
        else:
            cols.add(column_index_from_string(token.upper()))
    return cols


def main() -> int:
    ap = argparse.ArgumentParser(description="갱신보험료 엑셀 인쇄용 서식 정리")
    ap.add_argument("folder", nargs="?", default=".", help="대상 폴더 또는 .xlsx 파일 경로")
    ap.add_argument("--recursive", action="store_true", help="하위 폴더까지 처리")
    ap.add_argument("--backup", action="store_true", help="원본을 .xlsx.bak 으로 백업")
    ap.add_argument("--comma", action="store_true", help="금액 열에 1,000 단위 콤마 서식 추가")
    ap.add_argument(
        "--orientation",
        choices=("portrait", "landscape"),
        default="landscape",
        help="인쇄 방향 (기본: landscape)",
    )
    ap.add_argument(
        "--right-cols",
        default="",
        help='자동 판정을 무시하고 오른쪽 정렬할 열 (예: "G,I,J")',
    )
    ap.add_argument(
        "--center-cols",
        default="",
        help='자동 판정을 무시하고 가운데 정렬할 열 (예: "C,D")',
    )
    args = ap.parse_args()

    right_cols = _parse_cols(args.right_cols)
    center_cols = _parse_cols(args.center_cols)
    if right_cols & center_cols:
        sys.exit("--right-cols 와 --center-cols 에 같은 열을 동시에 넣을 수 없습니다.")

    target = Path(args.folder).expanduser()
    if not target.exists():
        sys.exit(f"경로를 찾을 수 없습니다: {target}")

    if target.is_file():
        files = [target]
    else:
        pattern = "**/*.xlsx" if args.recursive else "*.xlsx"
        files = sorted(p for p in target.glob(pattern) if not p.name.startswith("~$"))

    if not files:
        sys.exit(f"처리할 .xlsx 파일이 없습니다: {target}")

    ok, failed = 0, []
    for f in files:
        try:
            print("  [완료]", format_file(
                f,
                orientation=args.orientation,
                add_comma=args.comma,
                backup=args.backup,
                right_cols=right_cols,
                center_cols=center_cols,
            ))
            ok += 1
        except Exception as exc:  # 파일 하나가 실패해도 나머지는 계속
            failed.append((f.name, exc))
            print("  [실패]", f.name, "-", exc)

    print(f"\n총 {len(files)}개 중 {ok}개 처리 완료.")
    if failed:
        print("실패 목록:")
        for name, exc in failed:
            print(f"  - {name}: {exc}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())

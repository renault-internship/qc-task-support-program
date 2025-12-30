

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any, List, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook.workbook import Workbook
from openpyxl.cell.cell import MergedCell

from src.utils import (
    find_col_by_keywords_ws,
    parse_int_like,
    parse_excel_date,
    guess_last_data_row,
    AppError,
)

FILL_HIGHLIGHT = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# 전처리 1회 고정용 메타 시트
META_SHEET_NAME = "_PREPROCESS_META"
META_DONE_CELL = "A1"
META_TS_CELL = "A2"


# =========================================================
# 0) 전처리 1회만 가능: 마킹/체크
# =========================================================
def _is_already_preprocessed(wb: Workbook) -> bool:
    if META_SHEET_NAME not in wb.sheetnames:
        return False
    ws = wb[META_SHEET_NAME]
    v = ws[META_DONE_CELL].value
    return str(v).strip() == "1"


def _mark_preprocessed(wb: Workbook) -> None:
    if META_SHEET_NAME in wb.sheetnames:
        ws = wb[META_SHEET_NAME]
    else:
        ws = wb.create_sheet(META_SHEET_NAME)
        ws.sheet_state = "hidden"  # 숨김 처리

    ws[META_DONE_CELL].value = "1"
    ws[META_TS_CELL].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# =========================================================
# 1) 병합셀(MergedCell) 안전 처리
# =========================================================
def _resolve_merged_anchor(ws, row: int, col: int) -> Tuple[int, int]:
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return row, col

    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col

    return row, col


def _cell_safe(ws, row: int, col: int):
    ar, ac = _resolve_merged_anchor(ws, row, col)
    return ws.cell(row=ar, column=ac)


def set_cell_value_safe(ws, row: int, col: int, value: Any) -> None:
    _cell_safe(ws, row, col).value = value


def set_cell_fill_safe(ws, row: int, col: int, fill: PatternFill) -> None:
    _cell_safe(ws, row, col).fill = fill


# =========================================================
# 2) 기본 유틸
# =========================================================
def _is_blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def iter_data_rows(ws, data_start_row: int, last_row: int, anchor_col: int) -> List[int]:
    """
    anchor_col(보통 수리일자/클레임번호 등)에 값이 있는 행만 데이터 행으로 본다.
    국내 원본처럼 빈행+병합이 많은 경우 필수.
    """
    rows: List[int] = []
    for r in range(data_start_row, last_row + 1):
        v = ws.cell(row=r, column=anchor_col).value
        if not _is_blank(v):
            rows.append(r)
    return rows


# =========================================================
# 3) 컬럼 찾기(정확도 중요)
# =========================================================
def find_rate_col(ws, header_row: int) -> int:
    """
    구상율(Liability Ratio) 컬럼 정확히 찾기
    """
    try:
        return find_col_by_keywords_ws(ws, header_row, ["liability", "ratio"], mode="all")
    except Exception:
        return find_col_by_keywords_ws(ws, header_row, ["구상율"], mode="any")


def find_chargeback_col(ws, header_row: int) -> int:
    """
    구상금액(Chargeback Amount) 컬럼 정확히 찾기
    - '구상'만 넣으면 구상율에도 걸릴 수 있으므로 amount/금액을 반드시 포함
    """
    try:
        return find_col_by_keywords_ws(ws, header_row, ["chargeback", "amount"], mode="all")
    except Exception:
        pass

    try:
        return find_col_by_keywords_ws(ws, header_row, ["구상금액"], mode="any")
    except Exception:
        pass

    return find_col_by_keywords_ws(ws, header_row, ["구상", "금액"], mode="all")


def pick_mileage_col(ws, header_row: int) -> int:
    """
    국내 원본에서 주행거리 컬럼이 2개 이상 잡히는 케이스 대비.
    - 헤더에 '주행'/'mileage' 포함 후보 중 가장 오른쪽 우선
    """
    candidates: List[int] = []
    for c in range(1, ws.max_column + 1):
        hv = ws.cell(row=header_row, column=c).value
        if hv and isinstance(hv, str):
            s = hv.replace(" ", "")
            if ("주행" in s) or ("mileage" in hv.lower()):
                candidates.append(c)

    if candidates:
        return max(candidates)

    return find_col_by_keywords_ws(ws, header_row, ["mileage", "주행거리"], mode="any")


# =========================================================
# 4) 차계 병합 해제 + 채우기 (데이터 범위까지만)
# =========================================================
def unmerge_and_fill_column(ws, target_col: int, data_start_row: int, last_row: int) -> None:
    merged_ranges = list(ws.merged_cells.ranges)

    for mr in merged_ranges:
        if (mr.min_col <= target_col <= mr.max_col) and (mr.min_row >= data_start_row):
            top_left = ws.cell(mr.min_row, mr.min_col).value
            ws.unmerge_cells(str(mr))
            for r in range(mr.min_row, min(mr.max_row, last_row) + 1):
                set_cell_value_safe(ws, r, target_col, top_left)

    prev = None
    for r in range(data_start_row, last_row + 1):
        cur = ws.cell(row=r, column=target_col).value
        if _is_blank(cur):
            if not _is_blank(prev):
                set_cell_value_safe(ws, r, target_col, prev)
        else:
            prev = cur


# =========================================================
# 5) 구상율 변경(단일 진입점) + 바뀐 행 추적
# =========================================================
def set_rate(ws, row: int, rate_col: int, new_rate: float, changed_rows: set[int]) -> None:
    cell = _cell_safe(ws, row, rate_col)
    old = cell.value

    try:
        old_f = float(str(old).replace(",", "")) if not _is_blank(old) else None
    except Exception:
        old_f = None

    if old_f != float(new_rate):
        cell.value = float(new_rate)
        changed_rows.add(row)


# =========================================================
# 6) 구상금액 수식(데이터 행만)
# =========================================================
def set_chargeback_formula_rows(ws, data_rows: List[int], occ_col: int, rate_col: int, chb_col: int) -> None:
    for r in data_rows:
        occ_addr = ws.cell(row=r, column=occ_col).coordinate
        rate_addr = ws.cell(row=r, column=rate_col).coordinate
        set_cell_value_safe(ws, r, chb_col, f"={occ_addr}*({rate_addr}/100)")


# =========================================================
# 7) 아래 합계 행(SUM, 필터 무시)
# =========================================================
def add_sum_rows(ws, data_rows: List[int], occ_col: int, chb_col: int) -> None:
    if not data_rows:
        return

    first_row = data_rows[0]
    last_row = data_rows[-1]
    sum_start_row = last_row + 3

    set_cell_value_safe(ws, sum_start_row, occ_col - 1, "발생금액")
    set_cell_value_safe(
        ws,
        sum_start_row,
        occ_col,
        f"=SUM({ws.cell(row=first_row, column=occ_col).coordinate}:{ws.cell(row=last_row, column=occ_col).coordinate})",
    )

    set_cell_value_safe(ws, sum_start_row + 1, chb_col - 1, "구상금액")
    set_cell_value_safe(
        ws,
        sum_start_row + 1,
        chb_col,
        f"=SUM({ws.cell(row=first_row, column=chb_col).coordinate}:{ws.cell(row=last_row, column=chb_col).coordinate})",
    )


# =========================================================
# 8) 상단 서브토탈(SUBTOTAL 109, 필터 반영)
# =========================================================
def set_subtotal_if_empty(ws, target_col: int, data_rows: List[int], subtotal_row: int) -> None:
    if not data_rows:
        return

    cell = ws.cell(row=subtotal_row, column=target_col)
    if not _is_blank(cell.value):
        return

    first_row = data_rows[0]
    last_row = data_rows[-1]
    set_cell_value_safe(
        ws,
        subtotal_row,
        target_col,
        f"=SUBTOTAL(109,{ws.cell(row=first_row, column=target_col).coordinate}:{ws.cell(row=last_row, column=target_col).coordinate})",
    )



# =========================================================
# 9) 마일리지/보증기간 필터(데이터 행만)
# =========================================================
def apply_warranty_filters_ws(
    ws,
    header_row: int,
    data_rows: List[int],
    mileage_threshold: int,
    warranty_years: int,
    rate_col: int,
) -> set[int]:
    mileage_col = pick_mileage_col(ws, header_row)
    sale_col = find_col_by_keywords_ws(ws, header_row, ["sale date", "판매일", "sale"], mode="any")
    repair_col = find_col_by_keywords_ws(ws, header_row, ["repair date", "수리일자", "repair"], mode="any")

    warranty_days = int(warranty_years * 365)
    changed_rows: set[int] = set()

    for r in data_rows:
        mv = parse_int_like(ws.cell(row=r, column=mileage_col).value)
        if mv is not None and mv >= mileage_threshold:
            set_cell_fill_safe(ws, r, mileage_col, FILL_HIGHLIGHT)
            set_rate(ws, r, rate_col, 0, changed_rows)

        sale_dt = parse_excel_date(ws.cell(row=r, column=sale_col).value)
        repair_dt = parse_excel_date(ws.cell(row=r, column=repair_col).value)
        if sale_dt and repair_dt:
            if (repair_dt - sale_dt).days >= warranty_days:
                set_cell_fill_safe(ws, r, sale_col, FILL_HIGHLIGHT)
                set_rate(ws, r, rate_col, 0, changed_rows)

    for r in changed_rows:
        set_cell_fill_safe(ws, r, rate_col, FILL_HIGHLIGHT)

    return changed_rows


# =========================================================
# 10) 메인 처리(워크북 in-place)
# =========================================================
@dataclass
class CompanyConfig:
    sheet_index: int = 0
    header_row: int = 3
    data_start_row: int = 4
    mileage_threshold: int = 50000
    warranty_years: int = 2
    anchor_keywords: Tuple[str, ...] = ("repair date", "수리일자", "repair")


def process_wb_inplace(wb: Workbook, cfg: CompanyConfig) -> None:
    ws = wb.worksheets[cfg.sheet_index]

    vehicle_col = find_col_by_keywords_ws(ws, cfg.header_row, ["vehicle", "차계"], mode="any")
    occ_col = find_col_by_keywords_ws(ws, cfg.header_row, ["total cost", "발생", "발생금액"], mode="any")
    rate_col = find_rate_col(ws, cfg.header_row)
    chb_col = find_chargeback_col(ws, cfg.header_row)

    anchor_col = find_col_by_keywords_ws(ws, cfg.header_row, list(cfg.anchor_keywords), mode="any")
    last_row_guess = guess_last_data_row(ws, cfg.data_start_row, anchor_col=anchor_col, empty_run=30)

    unmerge_and_fill_column(ws, vehicle_col, cfg.data_start_row, last_row_guess)

    data_rows = iter_data_rows(ws, cfg.data_start_row, last_row_guess, anchor_col=anchor_col)
    if not data_rows:
        return

    apply_warranty_filters_ws(
        ws=ws,
        header_row=cfg.header_row,
        data_rows=data_rows,
        mileage_threshold=cfg.mileage_threshold,
        warranty_years=cfg.warranty_years,
        rate_col=rate_col,
    )

    set_chargeback_formula_rows(ws, data_rows, occ_col, rate_col, chb_col)
    add_sum_rows(ws, data_rows, occ_col, chb_col)

    # 상단 서브토탈: "구상금액" 기준
    subtotal_row = cfg.header_row - 1
    set_subtotal_if_empty(ws, target_col=chb_col, data_rows=data_rows, subtotal_row=subtotal_row)


# =========================================================
# 11) 파일 기반 처리(원하면 사용)
# =========================================================
def process_file(in_path: str, out_path: str, cfg: CompanyConfig) -> None:
    wb = load_workbook(in_path)
    process_wb_inplace(wb, cfg)
    wb.save(out_path)


# =========================================================
# 12) UI 엔트리
# =========================================================
def preprocess_inplace(wb: Workbook, company: str, keyword: str) -> None:
    """
    GUI 전처리 버튼 엔트리.
    company/keyword는 추후 룰 분기용.
    """
    try:
        # ✅ 전처리 1회만 허용
        if _is_already_preprocessed(wb):
            raise AppError("이미 전처리된 파일입니다. (전처리는 1회만 가능합니다)")

        cfg = CompanyConfig(
            sheet_index=0,
            header_row=3,
            data_start_row=4,
            mileage_threshold=50000,
            warranty_years=2,
        )

        process_wb_inplace(wb, cfg)

        # ✅ 전처리 완료 마킹
        _mark_preprocessed(wb)

    except AppError:
        raise
    except Exception as e:
        raise AppError(f"전처리 처리 중 오류: {e}") from e

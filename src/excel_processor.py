"""
excel_processor.py

목표(현재 단계)
- 와치스 원본 엑셀(병합/빈행 많음)에서도 안정적으로 전처리되게 만들기
- pandas/temp 파일 없이 openpyxl로만 처리
- MergedCell(병합셀) 쓰기 에러 방지
- "데이터 행만" 대상으로 수식/색칠/구상율 변경 적용 (빈 행/병합 빈행 스킵)

요구사항(반영)
1) 구상율이 변경된 행: 구상율 셀 색칠
2) 변경 원인 셀도 색칠
   - 주행거리 기준 초과: 주행거리 셀 색칠
   - 보증기간 기준 초과: 판매일/수리일자 셀 색칠
3) 구상금액은 전 행 수식으로 처리
   =발생금액*(구상율/100)
4) 발생금액/구상금액 합계(아래 2줄)는 필터 반영 X  -> SUM
5) 상단 서브토탈(헤더 위 1행)은 발생금액 기준, 필터 반영 O -> SUBTOTAL(109)

주의
- 국내 원본은 병합셀/빈행이 많아서 ws.max_row를 그대로 쓰면 안 됨
- guess_last_data_row()로 last_row를 잡되, 실제 "데이터 행"을 다시 한번 걸러서 처리함
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Tuple

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


# =========================================================
# 0) 병합셀(MergedCell) 안전 처리
# =========================================================
def _resolve_merged_anchor(ws, row: int, col: int) -> Tuple[int, int]:
    """
    (row, col)이 병합 범위 내부(MergedCell)이면 병합 범위의 좌상단 셀 좌표를 반환.
    병합이 아니면 그대로 반환.
    """
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
# 1) 데이터 행 판별(빈행/병합 빈행 스킵)
# =========================================================
def _is_blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def iter_data_rows(
    ws,
    data_start_row: int,
    last_row: int,
    anchor_col: int,
) -> List[int]:
    """
    anchor_col(보통 수리일자/클레임번호 등)에 값이 있는 행만 데이터 행으로 본다.
    국내 원본처럼 빈행+병합이 많은 경우, 이 필터가 핵심.
    """
    rows: List[int] = []
    for r in range(data_start_row, last_row + 1):
        v = ws.cell(row=r, column=anchor_col).value
        if not _is_blank(v):
            rows.append(r)
    return rows


# =========================================================
# 2) 차계 병합 해제 + 채우기 (데이터 범위까지만)
# =========================================================
def unmerge_and_fill_column(ws, target_col: int, data_start_row: int, last_row: int) -> None:
    """
    - data_start_row 이하(헤더/요약영역)는 건드리지 않음
    - target_col이 포함된 병합 범위를 해제하고, 해제된 영역의 target_col만 top-left 값으로 채움
    - 그 후 target_col에 대해 forward-fill (빈칸을 위 값으로 채움)
    """
    merged_ranges = list(ws.merged_cells.ranges)

    # 병합 해제 + 병합범위 target_col 채우기
    for mr in merged_ranges:
        if (mr.min_col <= target_col <= mr.max_col) and (mr.min_row >= data_start_row):
            top_left = ws.cell(mr.min_row, mr.min_col).value
            ws.unmerge_cells(str(mr))
            for r in range(mr.min_row, min(mr.max_row, last_row) + 1):
                set_cell_value_safe(ws, r, target_col, top_left)

    # 병합이 아닌 빈칸 ffill (last_row까지만)
    prev = None
    for r in range(data_start_row, last_row + 1):
        cur = ws.cell(row=r, column=target_col).value
        if _is_blank(cur):
            if not _is_blank(prev):
                set_cell_value_safe(ws, r, target_col, prev)
        else:
            prev = cur


# =========================================================
# 3) 구상율 변경(단일 진입점) + 바뀐 행 추적
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
# 4) 주행거리 컬럼이 2개 이상일 때 안전 처리
# =========================================================
def pick_mileage_col(ws, header_row: int) -> int:
    """
    국내 원본에서 주행거리 컬럼이 2개(X,Y)처럼 잡히는 케이스가 있음.
    - 우선 find_col_by_keywords_ws로 1개 찾고
    - 추가로 "주행"이 들어간 후보를 한 번 더 스캔해서 (가장 오른쪽) 후보로 사용
    - 실제 값 판별(숫자가 더 잘 나오는 쪽)은 추후 고도화 가능
    """
    # 1차: 기존 유틸로 찾기
    primary = find_col_by_keywords_ws(ws, header_row, ["mileage", "주행거리"], mode="any")

    # 2차: 헤더에서 후보 스캔(혹시 "주행거리(km)" "주행" 등이 여러 개면 가장 오른쪽 우선)
    candidates: List[int] = []
    for c in range(1, ws.max_column + 1):
        hv = ws.cell(row=header_row, column=c).value
        if hv and isinstance(hv, str):
            s = hv.replace(" ", "")
            if ("주행" in s) or ("mileage" in hv.lower()):
                candidates.append(c)

    if candidates:
        # 가장 오른쪽(보통 실제 값 컬럼이 오른쪽인 경우가 많음)
        return max(candidates)

    return primary


# =========================================================
# 5) 구상금액 수식(데이터 행만)
# =========================================================
def set_chargeback_formula_rows(
    ws,
    data_rows: List[int],
    occ_col: int,
    rate_col: int,
    chb_col: int,
) -> None:
    """
    구상금액 = 발생금액*(구상율/100)
    - 데이터 행만 적용(빈행/병합 빈행 스킵)
    """
    for r in data_rows:
        occ_addr = ws.cell(row=r, column=occ_col).coordinate
        rate_addr = ws.cell(row=r, column=rate_col).coordinate
        set_cell_value_safe(ws, r, chb_col, f"={occ_addr}*({rate_addr}/100)")


# =========================================================
# 6) 합계 행 추가(SUM, 필터 무시)
# =========================================================
def add_sum_rows(
    ws,
    data_rows: List[int],
    occ_col: int,
    chb_col: int,
    label_col_offset: int = 1,
) -> None:
    """
    데이터 아래 2줄 합계:
      - 발생금액(전체합) : SUM
      - 구상금액(전체합) : SUM
    """
    if not data_rows:
        return

    first_row = data_rows[0]
    last_row = data_rows[-1]
    sum_start_row = last_row + 3

    # 발생금액(전체합)
    set_cell_value_safe(ws, sum_start_row, occ_col - label_col_offset, "발생금액")
    set_cell_value_safe(
        ws,
        sum_start_row,
        occ_col,
        f"=SUM({ws.cell(row=first_row, column=occ_col).coordinate}:{ws.cell(row=last_row, column=occ_col).coordinate})",
    )

    # 구상금액(전체합)
    set_cell_value_safe(ws, sum_start_row + 1, chb_col - label_col_offset, "구상금액")
    set_cell_value_safe(
        ws,
        sum_start_row + 1,
        chb_col,
        f"=SUM({ws.cell(row=first_row, column=chb_col).coordinate}:{ws.cell(row=last_row, column=chb_col).coordinate})",
    )


# =========================================================
# 7) 상단 서브토탈(SUBTOTAL 109, 발생금액 기준)
# =========================================================
def set_subtotal_if_empty(
    ws,
    target_col: int,
    data_rows: List[int],
    subtotal_row: int,
    use_109: bool = True,
) -> None:
    """
    헤더 위 1행에 발생금액 기준 SUBTOTAL(109) 삽입.
    - 기존 셀이 비어있을 때만 넣음
    """
    if not data_rows:
        return

    cell = ws.cell(row=subtotal_row, column=target_col)
    if not _is_blank(cell.value):
        return

    func = 109 if use_109 else 9
    first_row = data_rows[0]
    last_row = data_rows[-1]

    set_cell_value_safe(
        ws,
        subtotal_row,
        target_col,
        f"=SUBTOTAL({func},{ws.cell(row=first_row, column=target_col).coordinate}:{ws.cell(row=last_row, column=target_col).coordinate})",
    )


# =========================================================
# 8) 마일리지/보증기간 필터(데이터 행만)
# =========================================================
def apply_warranty_filters_ws(
    ws,
    header_row: int,
    data_rows: List[int],
    changed_rows: set[int],
    mileage_threshold: int,
    warranty_years: int,
) -> int:
    """
    - 주행거리 기준 초과: 주행거리 셀 색칠 + 구상율 0
    - 보증기간 기준 초과: 판매일/수리일자 색칠 + 구상율 0
    - 변경된 행: 구상율 셀도 색칠
    """
    mileage_col = pick_mileage_col(ws, header_row)
    sale_col = find_col_by_keywords_ws(ws, header_row, ["sale date", "판매일", "sale"], mode="any")
    repair_col = find_col_by_keywords_ws(ws, header_row, ["repair date", "수리일자", "repair"], mode="any")
    rate_col = find_col_by_keywords_ws(ws, header_row, ["구상율", "liability ratio", "ratio"], mode="any")

    warranty_days = int(warranty_years * 365)

    for r in data_rows:
        # 1) 주행거리 기준
        mv = parse_int_like(ws.cell(row=r, column=mileage_col).value)
        if mv is not None and mv >= mileage_threshold:
            set_cell_fill_safe(ws, r, mileage_col, FILL_HIGHLIGHT)
            set_rate(ws, r, rate_col, 0, changed_rows)

        # 2) 보증기간 기준
        sale_dt = parse_excel_date(ws.cell(row=r, column=sale_col).value)
        repair_dt = parse_excel_date(ws.cell(row=r, column=repair_col).value)
        if sale_dt and repair_dt:
            if (repair_dt - sale_dt).days >= warranty_days:
                set_cell_fill_safe(ws, r, sale_col, FILL_HIGHLIGHT)
                set_cell_fill_safe(ws, r, repair_col, FILL_HIGHLIGHT)
                set_rate(ws, r, rate_col, 0, changed_rows)

    # 3) 구상율 변경행 색칠
    for r in changed_rows:
        set_cell_fill_safe(ws, r, rate_col, FILL_HIGHLIGHT)

    return rate_col


# =========================================================
# 9) 메인 처리(워크북 in-place)
# =========================================================
@dataclass
class CompanyConfig:
    sheet_index: int = 0
    header_row: int = 3
    data_start_row: int = 4
    mileage_threshold: int = 50000
    warranty_years: int = 2

    # last_row 추정용 anchor 키워드(대부분 수리일자)
    anchor_keywords: Tuple[str, ...] = ("repair date", "수리일자", "repair")


def process_wb_inplace(wb: Workbook, cfg: CompanyConfig) -> None:
    ws = wb.worksheets[cfg.sheet_index]

    # 핵심 컬럼 찾기
    vehicle_col = find_col_by_keywords_ws(ws, cfg.header_row, ["vehicle", "차계"], mode="any")
    occ_col = find_col_by_keywords_ws(ws, cfg.header_row, ["total cost", "발생", "발생금액"], mode="any")
    chb_col = find_col_by_keywords_ws(ws, cfg.header_row, ["chargeback", "구상", "구상금액"], mode="any")

    # last_row 추정 anchor (수리일자)
    anchor_col = find_col_by_keywords_ws(ws, cfg.header_row, list(cfg.anchor_keywords), mode="any")

    # ws.max_row는 믿으면 안 됨. anchor 기준으로 last_row 추정
    last_row_guess = guess_last_data_row(ws, cfg.data_start_row, anchor_col=anchor_col, empty_run=30)

    # 차계 병합 해제/채우기 (last_row_guess까지만)
    unmerge_and_fill_column(ws, vehicle_col, cfg.data_start_row, last_row_guess)

    # 실제 데이터 행 목록 (빈행/병합 빈행 스킵)
    data_rows = iter_data_rows(ws, cfg.data_start_row, last_row_guess, anchor_col=anchor_col)
    if not data_rows:
        return

    changed_rows: set[int] = set()

    # 마일리지/보증기간 필터 + 구상율 변경/색칠
    rate_col = apply_warranty_filters_ws(
        ws=ws,
        header_row=cfg.header_row,
        data_rows=data_rows,
        changed_rows=changed_rows,
        mileage_threshold=cfg.mileage_threshold,
        warranty_years=cfg.warranty_years,
    )

    # 구상금액 수식(데이터 행만)
    set_chargeback_formula_rows(ws, data_rows, occ_col, rate_col, chb_col)

    # 아래 합계 2줄 (필터 무시 SUM)
    add_sum_rows(ws, data_rows, occ_col, chb_col)

    # 상단 서브토탈(발생금액 기준, 필터 반영 SUBTOTAL 109)
    subtotal_row = cfg.header_row - 1
    set_subtotal_if_empty(ws, target_col=occ_col, data_rows=data_rows, subtotal_row=subtotal_row, use_109=True)


# =========================================================
# 10) 파일 기반 처리(원하면 사용)
# =========================================================
def process_file(in_path: str, out_path: str, cfg: CompanyConfig) -> None:
    wb = load_workbook(in_path)
    process_wb_inplace(wb, cfg)
    wb.save(out_path)


# =========================================================
# 11) UI 엔트리
# =========================================================
def preprocess_inplace(wb: Workbook, company: str, keyword: str) -> None:
    """
    GUI 전처리 버튼 엔트리.
    - company/keyword는 추후 룰 분기용
    - 현재는 AMS 기본 설정으로 처리
    """
    try:
        # TODO: 회사별로 header_row/data_start_row/threshold 다르면 여기서 분기
        cfg = CompanyConfig(
            sheet_index=0,
            header_row=3,
            data_start_row=4,
            mileage_threshold=50000,
            warranty_years=2,
        )
        process_wb_inplace(wb, cfg)
    except Exception as e:
        raise AppError(f"전처리 처리 중 오류: {e}") from e

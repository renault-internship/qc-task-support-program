"""
엑셀 파일 처리 모듈
"""
from typing import Dict, Any
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from src.utils import (
    find_col_by_keywords_ws,
    parse_int_like,
    parse_excel_date,
    guess_last_data_row
)


def unmerge_and_fill_column(ws, target_col: int, data_start_row: int, last_row: int):
    """
    차계 병합 해제 + 채우기 (last_row까지만)
    """
    merged_ranges = list(ws.merged_cells.ranges)

    # 병합 해제 + 병합범위 전체 채우기
    for mr in merged_ranges:
        if (mr.min_col <= target_col <= mr.max_col) and (mr.min_row >= data_start_row):
            top_left = ws.cell(mr.min_row, mr.min_col).value
            ws.unmerge_cells(str(mr))
            for r in range(mr.min_row, min(mr.max_row, last_row) + 1):
                ws.cell(row=r, column=target_col).value = top_left

    # 병합이 아닌 빈칸 ffill (last_row까지만)
    prev = None
    for r in range(data_start_row, last_row + 1):
        c = ws.cell(row=r, column=target_col)
        if c.value in (None, ""):
            if prev not in (None, ""):
                c.value = prev
        else:
            prev = c.value


def set_rate(ws, row: int, rate_col: int, new_rate: float, changed_rows: set[int]):
    """
    구상율 변경(단일 진입점) + 바뀐 행 추적
    """
    cell = ws.cell(row=row, column=rate_col)
    old = cell.value

    try:
        old_f = float(str(old).replace(",", "")) if old not in (None, "") else None
    except:
        old_f = None

    if old_f != float(new_rate):
        cell.value = float(new_rate)
        changed_rows.add(row)


def set_chargeback_formula_rows(ws, rows: set[int], occ_col: int, rate_col: int, chb_col: int):
    """
    바뀐 행만 구상금액 수식
    """
    for r in rows:
        occ_addr = ws.cell(row=r, column=occ_col).coordinate
        rate_addr = ws.cell(row=r, column=rate_col).coordinate
        ws.cell(row=r, column=chb_col).value = f"={occ_addr}*({rate_addr}/100)"


def add_sum_rows(ws, data_start_row: int, last_row: int, occ_col: int, chb_col: int):
    """
    발생/구상 합계 행 추가
    """
    sum_start_row = last_row + 3

    ws.cell(row=sum_start_row, column=occ_col - 1).value = "발생금액"
    ws.cell(row=sum_start_row, column=occ_col).value = (
        f"=SUM({ws.cell(row=data_start_row, column=occ_col).coordinate}:"
        f"{ws.cell(row=last_row, column=occ_col).coordinate})"
    )

    ws.cell(row=sum_start_row + 1, column=chb_col - 1).value = "구상금액"
    ws.cell(row=sum_start_row + 1, column=chb_col).value = (
        f"=SUM({ws.cell(row=data_start_row, column=chb_col).coordinate}:"
        f"{ws.cell(row=last_row, column=chb_col).coordinate})"
    )


def set_subtotal(ws, target_col: int, data_start_row: int, last_row: int, subtotal_row: int, use_109: bool = True):
    """
    SUBTOTAL 수식
    """
    func = 109 if use_109 else 9
    ws.cell(row=subtotal_row, column=target_col).value = (
        f"=SUBTOTAL({func},"
        f"{ws.cell(row=data_start_row, column=target_col).coordinate}:"
        f"{ws.cell(row=last_row, column=target_col).coordinate})"
    )


def apply_warranty_filters_ws(
    ws,
    header_row: int,
    data_start_row: int,
    last_row: int,
    changed_rows: set[int],
    mileage_threshold: int = 50000,
    warranty_years: int = 2,
):
    """
    마일리지/보증기간 필터 (색칠 + 구상율 0)
    """
    fill_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    mileage_col = find_col_by_keywords_ws(ws, header_row, ["mileage", "주행거리"], mode="any")
    sale_col    = find_col_by_keywords_ws(ws, header_row, ["sale date", "판매일", "sale"], mode="any")
    repair_col  = find_col_by_keywords_ws(ws, header_row, ["repair date", "수리일자", "repair"], mode="any")
    rate_col    = find_col_by_keywords_ws(ws, header_row, ["구상율", "liability ratio", "ratio"], mode="any")

    warranty_days = int(warranty_years * 365)

    for r in range(data_start_row, last_row + 1):
        mv = parse_int_like(ws.cell(row=r, column=mileage_col).value)
        if mv is not None and mv >= mileage_threshold:
            ws.cell(row=r, column=mileage_col).fill = fill_color
            set_rate(ws, r, rate_col, 0, changed_rows)

        sale_dt = parse_excel_date(ws.cell(row=r, column=sale_col).value)
        repair_dt = parse_excel_date(ws.cell(row=r, column=repair_col).value)

        if sale_dt and repair_dt:
            if (repair_dt - sale_dt).days >= warranty_days:
                ws.cell(row=r, column=sale_col).fill = fill_color
                set_rate(ws, r, rate_col, 0, changed_rows)

    return rate_col


def process_all(in_path: str, out_path: str, company_info: Dict[str, Any]):
    """
    전체 처리 파이프라인
    
    Args:
        in_path: 입력 엑셀 파일 경로
        out_path: 출력 엑셀 파일 경로
        company_info: 기업정보 딕셔너리 (DB에서 가져온 정보)
    """
    sheet_index = company_info.get("sheet_index", 0)
    header_row = company_info.get("header_row", 3)
    data_start_row = company_info.get("data_start_row", 4)
    mileage_threshold = company_info.get("mileage_threshold", 50000)
    warranty_years = company_info.get("warranty_years", 2)

    wb = load_workbook(in_path)
    ws = wb.worksheets[sheet_index]

    vehicle_col = find_col_by_keywords_ws(ws, header_row, ["vehicle", "차계"], mode="any")
    occ_col     = find_col_by_keywords_ws(ws, header_row, ["total cost", "발생", "발생금액"], mode="any")
    chb_col     = find_col_by_keywords_ws(ws, header_row, ["chargeback", "구상", "구상금액"], mode="any")
    repair_col  = find_col_by_keywords_ws(ws, header_row, ["repair date", "수리일자", "repair"], mode="any")

    last_row = guess_last_data_row(ws, data_start_row, anchor_col=repair_col, empty_run=30)

    unmerge_and_fill_column(ws, vehicle_col, data_start_row, last_row)

    changed_rows: set[int] = set()

    rate_col = apply_warranty_filters_ws(
        ws=ws,
        header_row=header_row,
        data_start_row=data_start_row,
        last_row=last_row,
        changed_rows=changed_rows,
        mileage_threshold=mileage_threshold,
        warranty_years=warranty_years,
    )

    set_chargeback_formula_rows(ws, changed_rows, occ_col, rate_col, chb_col)

    add_sum_rows(ws, data_start_row, last_row, occ_col, chb_col)

    # ✅ "겉보기 빈칸"까지 포함해서 SUBTOTAL 안전 삽입
    subtotal_row = header_row - 1
    v = ws.cell(row=subtotal_row, column=chb_col).value
    if v is None or str(v).strip() == "":
        set_subtotal(ws, chb_col, data_start_row, last_row, subtotal_row, use_109=True)

    wb.save(out_path)


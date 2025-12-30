# =========================
# src/gui/models.py
# =========================
"""
Excel 데이터 모델
- 병합 셀(merged cell) 표시값 채우기(방법 B): UI에서만 병합값이 아래로 계속 보이게
- 병합 셀 편집은 좌상단(top-left)만 허용
- 기존 기능(숫자/날짜 포맷, 구상율 컬럼만 편집, dirty 표시) 유지
"""
from __future__ import annotations

from datetime import datetime, date
from typing import Dict, Tuple, Any

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex


class ExcelSheetModel(QAbstractTableModel):
    """
    - openpyxl worksheet를 UI로 보여주는 모델
    - dirty: UI에서 수정된 값(메모리상)
    - 표시 포맷: 숫자 콤마, 날짜 yyyy-mm-dd
    - 편집 제어:
        * edit_all=True  -> 헤더(1행) 제외 전부 편집 (단, 병합셀은 좌상단만)
        * edit_all=False -> '구상율' 헤더 컬럼만 편집 (단, 병합셀은 좌상단만)

    - 병합 표시(방법 B):
        * 병합 범위 내부 셀은 '좌상단 셀' 값을 보여줌(값 아래로 채워 보이게)
        * 실제 ws 값은 절대 변경하지 않음 (export 시에도 그대로)
        * 편집은 좌상단만 가능하게 막음(데이터 꼬임 방지)
    """

    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.ws = ws
        self.max_row = ws.max_row
        self.max_col = ws.max_column

        self.dirty: Dict[Tuple[int, int], Any] = {}
        self.edit_all: bool = False
        self.editable_cols: set[int] = self._find_chargeback_rate_cols()

        # (r,c) -> (top_r, top_c) 병합 캐시
        self._merge_top_left: Dict[Tuple[int, int], Tuple[int, int]] = {}
        # (top_r, top_c) -> (min_row, min_col, max_row, max_col) 병합 범위 캐시(최적화용)
        self._merge_bounds_by_top: Dict[Tuple[int, int], Tuple[int, int, int, int]] = {}

        self._build_merge_cache()

    # ---------- 병합 캐시 ----------
    def _build_merge_cache(self):
        self._merge_top_left.clear()
        self._merge_bounds_by_top.clear()

        for mr in self.ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = mr.bounds
            top = (min_row, min_col)
            self._merge_bounds_by_top[top] = (min_row, min_col, max_row, max_col)

            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    self._merge_top_left[(r, c)] = top

    def _canonical_cell(self, r: int, c: int) -> Tuple[int, int]:
        """병합셀 내부면 좌상단 좌표로, 아니면 자기 자신."""
        return self._merge_top_left.get((r, c), (r, c))

    def _is_merged_non_topleft(self, r: int, c: int) -> bool:
        """병합 범위 안인데 좌상단이 아닌 셀인지"""
        top = self._merge_top_left.get((r, c))
        return (top is not None) and (top != (r, c))

    # ----- Qt 필수 -----
    def rowCount(self, parent=QModelIndex()):
        return self.max_row

    def columnCount(self, parent=QModelIndex()):
        return self.max_col

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        r = index.row() + 1
        c = index.column() + 1

        # 병합이면 좌상단 기준으로 값 조회(방법 B)
        cr, cc = self._canonical_cell(r, c)

        v = self.dirty.get((cr, cc), self.ws.cell(row=cr, column=cc).value)

        if role == Qt.EditRole:
            return "" if v is None else v

        if role == Qt.DisplayRole:
            return self._format_value(v)

        if role == Qt.BackgroundRole:
            # 수정된 셀 표시(병합이면 좌상단 기준)
            if (cr, cc) in self.dirty:
                from PySide6.QtGui import QBrush, QColor
                return QBrush(QColor(255, 250, 205))  # 연노랑
            return None

        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags

        r = index.row() + 1
        c = index.column() + 1
        base = Qt.ItemIsSelectable | Qt.ItemIsEnabled

        # 헤더 행(1행)은 편집 막기
        if r == 1:
            return base

        # 병합셀은 좌상단만 편집 가능
        if self._is_merged_non_topleft(r, c):
            return base

        if self.edit_all:
            return base | Qt.ItemIsEditable

        # 구상율 컬럼만 편집
        if c in self.editable_cols:
            return base | Qt.ItemIsEditable

        return base

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False

        r = index.row() + 1
        c = index.column() + 1

        if r == 1:
            return False

        # 병합셀 내부 클릭이면 좌상단으로 정규화
        cr, cc = self._canonical_cell(r, c)

        # 좌상단이 아닌 병합셀은 편집 막기
        if self._is_merged_non_topleft(r, c):
            return False

        if not self.edit_all and (cc not in self.editable_cols):
            return False

        new_val = self._parse_user_input(value)
        self.dirty[(cr, cc)] = new_val

        # 병합 범위가 있으면 범위만 갱신(최소 갱신)
        top = (cr, cc)
        if top in self._merge_bounds_by_top:
            min_row, min_col, max_row, max_col = self._merge_bounds_by_top[top]
            tl = self.index(min_row - 1, min_col - 1)
            br = self.index(max_row - 1, max_col - 1)
            self.dataChanged.emit(tl, br, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
        else:
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])

        return True

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.excel_col_name(section + 1)
        return str(section + 1)

    # ----- 유틸 -----
    def set_edit_all(self, on: bool):
        self.edit_all = bool(on)

    @staticmethod
    def excel_col_name(n: int) -> str:
        name = ""
        while n:
            n, rem = divmod(n - 1, 26)
            name = chr(65 + rem) + name
        return name

    @staticmethod
    def _format_value(v):
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, int):
            return f"{v:,}"
        if isinstance(v, float):
            s = f"{v:,.2f}"
            return s.rstrip("0").rstrip(".")
        if isinstance(v, (datetime, date)):
            return v.strftime("%Y-%m-%d")
        return str(v)

    @staticmethod
    def _parse_user_input(value):
        text = "" if value is None else str(value).strip()
        if text == "":
            return None

        raw = text.replace(",", "")

        try:
            if "." in raw:
                return float(raw)
            return int(raw)
        except ValueError:
            return text

    def _find_chargeback_rate_cols(self) -> set[int]:
        """
        1행(헤더)에서 '구상'+'율' 포함 컬럼을 찾아 편집 가능 컬럼으로 등록
        """
        editable = set()
        header_row = 1

        for c in range(1, self.max_col + 1):
            hv = self.ws.cell(row=header_row, column=c).value
            if hv and isinstance(hv, str):
                s = hv.replace(" ", "")
                if ("구상" in s and "율" in s) or ("chargeback" in hv.lower() and "rate" in hv.lower()):
                    editable.add(c)
        return editable

    def apply_dirty_to_sheet(self):
        """
        dirty를 실제 ws에 반영
        - 병합셀의 경우 dirty는 항상 좌상단 기준으로만 기록됨
        """
        for (r, c), v in self.dirty.items():
            self.ws.cell(row=r, column=c).value = v
        # dirty 유지(화면 표시/후속 반영용)

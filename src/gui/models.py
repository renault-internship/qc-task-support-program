"""
Excel 데이터 모델
"""
from datetime import datetime, date

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex


class ExcelSheetModel(QAbstractTableModel):
    """
    - openpyxl worksheet를 UI로 보여주는 모델
    - dirty: UI에서 수정된 값(메모리상)
    - 표시 포맷: 숫자 콤마, 날짜 yyyy-mm-dd
    - 편집 제어:
        * edit_all=True  -> 헤더(1행) 제외 전부 편집
        * edit_all=False -> '구상율' 헤더 컬럼만 편집
    """
    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.ws = ws
        self.max_row = ws.max_row
        self.max_col = ws.max_column

        self.dirty: dict[tuple[int, int], object] = {}
        self.edit_all: bool = False
        self.editable_cols: set[int] = self._find_chargeback_rate_cols()

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

        v = self.dirty.get((r, c), self.ws.cell(row=r, column=c).value)

        if role == Qt.EditRole:
            # 편집기는 타입 보존(가능한 한)
            return "" if v is None else v

        if role == Qt.DisplayRole:
            return self._format_value(v)

        if role == Qt.BackgroundRole:
            # 수정된 셀은 표시(엑셀 느낌)
            if (r, c) in self.dirty:
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

        # 보통 헤더 행(1행)은 편집 막는 게 안전
        if r == 1:
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

        # 편집 불가 셀 방어
        if r == 1:
            return False
        if (not self.edit_all) and (c not in self.editable_cols):
            return False

        new_val = self._parse_user_input(value)

        self.dirty[(r, c)] = new_val
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

        # 숫자 파싱 (int/float)
        try:
            if "." in raw:
                return float(raw)
            return int(raw)
        except ValueError:
            return text

    def _find_chargeback_rate_cols(self) -> set[int]:
        """
        1행(헤더)에서 '구상'+'율' 포함 컬럼을 찾아 편집 가능 컬럼으로 등록
        (파일마다 컬럼 위치가 다를 수 있어 탐지 기반으로 처리)
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
        for (r, c), v in self.dirty.items():
            self.ws.cell(row=r, column=c).value = v
        # 반영 후 dirty를 유지할지(되돌리기 위해) 정책 선택 가능
        # 지금은 "화면 표시(수정 강조)"를 위해 유지


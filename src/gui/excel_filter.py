# =========================
# src/gui/excel_filter.py
# =========================
from __future__ import annotations

from typing import Dict, Optional, Set, List

from PySide6.QtCore import Qt, QSortFilterProxyModel, QModelIndex
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton,
    QListWidget, QListWidgetItem, QLabel
)


_EMPTY_TOKEN = "(빈값)"


class ExcelFilterProxyModel(QSortFilterProxyModel):
    """
    - 기존 QSortFilterProxyModel(검색/정렬)은 유지
    - 추가로 '컬럼별 값 필터'를 AND 조건으로 적용
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self._col_allowed: Dict[int, Optional[Set[str]]] = {}  # col -> allowed set, None이면 필터 없음

    def clear_all_column_filters(self) -> None:
        self._col_allowed.clear()
        self.invalidateFilter()

    def clear_column_filter(self, col: int) -> None:
        if col in self._col_allowed:
            del self._col_allowed[col]
            self.invalidateFilter()

    def set_column_filter(self, col: int, allowed_values: Optional[Set[str]]) -> None:
        # allowed_values:
        #   None -> 필터 해제(전체 허용)
        #   set() -> 아무것도 허용 안함(전부 숨김)
        self._col_allowed[col] = allowed_values
        self.invalidateFilter()

    def get_column_filter(self, col: int) -> Optional[Set[str]]:
        return self._col_allowed.get(col)

    def _cell_text(self, source_row: int, source_col: int) -> str:
        src = self.sourceModel()
        if src is None:
            return ""
        idx = src.index(source_row, source_col)
        v = src.data(idx, Qt.DisplayRole)
        if v is None:
            return ""
        s = str(v).strip()
        return s

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        # 1) 기존 검색(정규식)은 super가 처리
        if not super().filterAcceptsRow(source_row, source_parent):
            return False

        # 2) 컬럼별 필터 AND
        src = self.sourceModel()
        if src is None:
            return True

        col_count = src.columnCount()
        for col, allowed in self._col_allowed.items():
            if allowed is None:
                continue
            if col < 0 or col >= col_count:
                continue

            txt = self._cell_text(source_row, col)
            key = _EMPTY_TOKEN if txt == "" else txt

            if key not in allowed:
                return False

        return True


class ColumnFilterDialog(QDialog):
    """
    엑셀 스타일 최소 구현:
    - 컬럼의 '고유값' 리스트(체크박스)
    - 검색창(리스트 내 필터)
    - 전체 선택/해제
    - 적용/해제
    """
    def __init__(self, source_model, proxy: ExcelFilterProxyModel, col: int, col_name: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"필터 - {col_name}")
        self.resize(360, 520)

        self.source_model = source_model
        self.proxy = proxy
        self.col = int(col)

        root = QVBoxLayout(self)

        root.addWidget(QLabel(f"컬럼: {col_name}"))

        # 리스트 검색
        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("값 검색")
        top.addWidget(self.search, 1)

        self.btn_all = QPushButton("전체 선택")
        self.btn_none = QPushButton("전체 해제")
        top.addWidget(self.btn_all)
        top.addWidget(self.btn_none)
        root.addLayout(top)

        # 값 리스트
        self.listw = QListWidget()
        self.listw.setSelectionMode(QListWidget.NoSelection)
        root.addWidget(self.listw, 1)

        # 하단 버튼
        bottom = QHBoxLayout()
        self.btn_clear = QPushButton("이 컬럼 필터 해제")
        self.btn_apply = QPushButton("적용")
        self.btn_cancel = QPushButton("취소")
        bottom.addWidget(self.btn_clear)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_apply)
        bottom.addWidget(self.btn_cancel)
        root.addLayout(bottom)

        # 데이터 채우기
        self._load_unique_values()

        # 기존 필터 상태 반영
        self._apply_existing_state()

        # 시그널
        self.search.textChanged.connect(self._on_search)
        self.btn_all.clicked.connect(self._check_all)
        self.btn_none.clicked.connect(self._uncheck_all)
        self.btn_clear.clicked.connect(self._clear_filter)
        self.btn_apply.clicked.connect(self._apply)
        self.btn_cancel.clicked.connect(self.reject)

    def _load_unique_values(self) -> None:
        self.listw.clear()

        vals: Set[str] = set()
        row_count = self.source_model.rowCount()
        for r in range(row_count):
            idx = self.source_model.index(r, self.col)
            v = self.source_model.data(idx, Qt.DisplayRole)
            s = "" if v is None else str(v).strip()
            vals.add(_EMPTY_TOKEN if s == "" else s)

        for v in sorted(vals):
            item = QListWidgetItem(v)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.listw.addItem(item)

    def _apply_existing_state(self) -> None:
        current = self.proxy.get_column_filter(self.col)
        if current is None:
            return  # 전체 선택 상태 유지

        # current가 있으면 그거만 체크
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            it.setCheckState(Qt.Checked if it.text() in current else Qt.Unchecked)

    def _on_search(self, text: str) -> None:
        q = (text or "").strip().lower()
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            it.setHidden(q not in it.text().lower())

    def _check_all(self) -> None:
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            if not it.isHidden():
                it.setCheckState(Qt.Checked)

    def _uncheck_all(self) -> None:
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            if not it.isHidden():
                it.setCheckState(Qt.Unchecked)

    def _clear_filter(self) -> None:
        self.proxy.clear_column_filter(self.col)
        self.accept()

    def _apply(self) -> None:
        selected: Set[str] = set()
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            if it.checkState() == Qt.Checked:
                selected.add(it.text())

        # 전체가 체크된 경우 -> 필터 해제(None)
        if selected and len(selected) == self.listw.count():
            self.proxy.set_column_filter(self.col, None)
        else:
            self.proxy.set_column_filter(self.col, selected)

        self.accept()

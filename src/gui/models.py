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
from typing import Dict, Tuple, Any, List, Optional
from collections import deque

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
        
        # 필터 상태 확인용 proxy_model 참조 (SUBTOTAL 계산 시 필요)
        self.proxy_model = None

        # (r,c) -> (top_r, top_c) 병합 캐시
        self._merge_top_left: Dict[Tuple[int, int], Tuple[int, int]] = {}
        # (top_r, top_c) -> (min_row, min_col, max_row, max_col) 병합 범위 캐시(최적화용)
        self._merge_bounds_by_top: Dict[Tuple[int, int], Tuple[int, int, int, int]] = {}

        self._build_merge_cache()
        
        # Undo/Redo 스택
        self._undo_stack: deque = deque(maxlen=100)  # 최대 100개까지 저장
        self._redo_stack: deque = deque(maxlen=100)
        self._is_undoing: bool = False  # Undo/Redo 중인지 플래그
    
    def set_proxy_model(self, proxy_model):
        """필터 상태 확인을 위한 proxy_model 참조 설정"""
        self.proxy_model = proxy_model

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

        # 병합된 셀의 경우, 좌상단이 아닌 셀에서는 빈 문자열 반환
        if self._is_merged_non_topleft(r, c):
            if role == Qt.DisplayRole or role == Qt.EditRole:
                return ""
            return None

        # 병합이면 좌상단 기준으로 값 조회
        cr, cc = self._canonical_cell(r, c)

        v = self.dirty.get((cr, cc), self.ws.cell(row=cr, column=cc).value)

        if role == Qt.EditRole:
            return "" if v is None else v

        if role == Qt.DisplayRole:
            # 수식이면 표시용으로 계산값을 보여주고, 아니면 원래 값 표시
            v_display = self._display_value(v, r=cr, c=cc)
            return self._format_value(v_display)

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

        # 현재 값 가져오기 (편집 전 값)
        old_val = self.dirty.get((cr, cc))
        if old_val is None:
            # dirty에 없으면 원본 워크시트에서 가져오기
            old_val = self.ws.cell(row=cr, column=cc).value

        new_val = self._parse_user_input(value)
        
        # Undo/Redo 중이 아니면 히스토리에 저장
        if not self._is_undoing:
            # 이전 값과 새 값이 다를 때만 히스토리에 추가
            if old_val != new_val:
                self._undo_stack.append({
                    'row': cr,
                    'col': cc,
                    'old_value': old_val,
                    'new_value': new_val
                })
                # 새 편집이 발생하면 redo 스택 초기화
                self._redo_stack.clear()
        
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
    
    # ================= Undo/Redo =================
    def can_undo(self) -> bool:
        """실행취소 가능한지 확인"""
        return len(self._undo_stack) > 0
    
    def can_redo(self) -> bool:
        """다시실행 가능한지 확인"""
        return len(self._redo_stack) > 0
    
    def undo(self) -> bool:
        """마지막 편집을 실행취소"""
        if not self.can_undo():
            return False
        
        self._is_undoing = True
        
        try:
            # undo 스택에서 마지막 편집 가져오기
            edit = self._undo_stack.pop()
            row, col = edit['row'], edit['col']
            old_val = edit['old_value']
            new_val = edit['new_value']
            
            # redo 스택에 저장 (다시 실행할 수 있도록)
            # redo는 old_val(현재 값)에서 new_val(편집 후 값)로 다시 적용하는 것
            self._redo_stack.append({
                'row': row,
                'col': col,
                'old_value': old_val,  # 현재 값 (편집 전 값)
                'new_value': new_val    # 다시 적용할 값 (편집 후 값)
            })
            
            # 값 복원
            if old_val is None:
                # 원래 값이 None이면 dirty에서 제거
                if (row, col) in self.dirty:
                    del self.dirty[(row, col)]
            else:
                self.dirty[(row, col)] = old_val
            
            # UI 업데이트
            index = self.index(row - 1, col - 1)
            top = (row, col)
            if top in self._merge_bounds_by_top:
                min_row, min_col, max_row, max_col = self._merge_bounds_by_top[top]
                tl = self.index(min_row - 1, min_col - 1)
                br = self.index(max_row - 1, max_col - 1)
                self.dataChanged.emit(tl, br, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
            else:
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
            
            return True
        finally:
            self._is_undoing = False
    
    def redo(self) -> bool:
        """취소한 편집을 다시 실행"""
        if not self.can_redo():
            return False
        
        self._is_undoing = True
        
        try:
            # redo 스택에서 마지막 편집 가져오기
            edit = self._redo_stack.pop()
            row, col = edit['row'], edit['col']
            old_val = edit['old_value']
            new_val = edit['new_value']
            
            # undo 스택에 저장 (다시 취소할 수 있도록)
            self._undo_stack.append({
                'row': row,
                'col': col,
                'old_value': old_val,
                'new_value': new_val
            })
            
            # 값 적용
            self.dirty[(row, col)] = new_val
            
            # UI 업데이트
            index = self.index(row - 1, col - 1)
            top = (row, col)
            if top in self._merge_bounds_by_top:
                min_row, min_col, max_row, max_col = self._merge_bounds_by_top[top]
                tl = self.index(min_row - 1, min_col - 1)
                br = self.index(max_row - 1, max_col - 1)
                self.dataChanged.emit(tl, br, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
            else:
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
            
            return True
        finally:
            self._is_undoing = False

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
            # 모든 float는 정수로 반올림하여 표시 (엑셀 스타일)
            return f"{int(round(v)):,}"
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
    def _display_value(self, v: Any, r: int, c: int) -> Any:
        """
        UI 표시용:
        - 값이 수식("=...")이면 계산 가능한 범위에서 숫자로 보여줌
        - 계산 못하면 원문 수식 그대로 보여줌
        """
        if not isinstance(v, str):
            return v
        s = v.strip()
        if not s.startswith("="):
            return v

        # 1) SUM 함수: =SUM(A1:A10)
        if "SUM(" in s.upper():
            try:
                return self._eval_sum(s, r, c)
            except Exception:
                return v
        
        # 2) SUBTOTAL 함수: =SUBTOTAL(9, A1:A10)
        if "SUBTOTAL(9," in s.upper() or "SUBTOTAL(9," in s:
            try:
                return self._eval_subtotal(s, r, c)
            except Exception:
                return v

        # 3) 단순 셀 참조: =T388
        import re
        cell_ref_match = re.fullmatch(r"=\s*([A-Z]{1,3}\d+)\s*", s, re.IGNORECASE)
        if cell_ref_match:
            try:
                ref_addr = cell_ref_match.group(1).upper()
                ref_row, ref_col = self._addr_to_row_col(ref_addr)
                # 참조된 셀의 값 읽기 (재귀적으로 수식 계산)
                ref_cell = self.ws.cell(row=ref_row, column=ref_col)
                ref_value = self.dirty.get((ref_row, ref_col), ref_cell.value)
                # 참조된 값이 수식이면 재귀적으로 계산
                if isinstance(ref_value, str) and ref_value.strip().startswith("="):
                    return self._display_value(ref_value, ref_row, ref_col)
                # 숫자면 그대로 반환
                if isinstance(ref_value, (int, float)):
                    return ref_value
                # 문자열 숫자면 파싱
                if isinstance(ref_value, str):
                    try:
                        return float(ref_value.replace(",", ""))
                    except:
                        return ref_value
                return ref_value
            except Exception:
                return v

        # 4) 아주 흔한 패턴: =T4*(U4/100)
        try:
            return self._eval_simple_mul_div(s)
        except Exception:
            return v

    def _eval_simple_mul_div(self, formula: str) -> float:
        """
        지원 범위: =A1*(B1/100) 또는 =A1*(B1/100.0) 비슷한 단순 산술
        정확한 소숫점 계산 (반올림 안 함)
        """
        import re

        # 예: =T4*(U4/100)
        m = re.fullmatch(r"=\s*([A-Z]{1,3}\d+)\s*\*\s*\(\s*([A-Z]{1,3}\d+)\s*/\s*(\d+(\.\d+)?)\s*\)\s*", formula, re.IGNORECASE)
        if not m:
            raise ValueError("not supported")

        a_addr = m.group(1).upper()
        b_addr = m.group(2).upper()
        denom = float(m.group(3))

        a = self._read_number(a_addr)
        b = self._read_number(b_addr)
        result = a * (b / denom)
        # 정확한 소숫점 계산 유지
        return result

    def _read_number(self, addr: str) -> float:
        """
        셀 주소(A1) -> 숫자값 읽기
        - 문자열 숫자("55,310")도 처리
        - 비어있으면 0
        """
        import re
        mm = re.fullmatch(r"([A-Z]{1,3})(\d+)", addr)
        if not mm:
            return 0.0

        col_letters = mm.group(1)
        row = int(mm.group(2))

        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - 64)

        # 병합이면 좌상단으로 정규화
        row, col = self._canonical_cell(row, col)

        vv = self.dirty.get((row, col), self.ws.cell(row=row, column=col).value)

        if vv is None:
            return 0.0
        if isinstance(vv, (int, float)):
            return float(vv)
        if isinstance(vv, str):
            t = vv.strip().replace(",", "")
            try:
                return float(t)
            except Exception:
                return 0.0
        return 0.0
    
    def _eval_sum(self, formula: str, row: int, col: int) -> float:
        """
        SUM 함수 계산: =SUM(A1:A10)
        병합 셀은 한 번만 계산 (중복 방지)
        """
        import re
        # =SUM(A1:A10) 또는 =SUM( A1 : A10 ) 패턴 매칭
        m = re.search(r"SUM\s*\(\s*([A-Z]{1,3}\d+)\s*:\s*([A-Z]{1,3}\d+)\s*\)", formula, re.IGNORECASE)
        if not m:
            raise ValueError("SUM range not found")
        
        start_addr = m.group(1).upper()
        end_addr = m.group(2).upper()
        
        # 주소를 행/열로 변환
        start_row, start_col = self._addr_to_row_col(start_addr)
        end_row, end_col = self._addr_to_row_col(end_addr)
        
        # 범위 내 모든 셀 값 합산 (병합 셀 중복 방지)
        total = 0.0
        processed_cells = set()  # 이미 처리한 병합 셀 추적
        
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                if 1 <= r <= self.max_row and 1 <= c <= self.max_col:
                    # 병합이면 좌상단으로 정규화
                    canonical_r, canonical_c = self._canonical_cell(r, c)
                    
                    # 이미 처리한 셀이면 스킵 (병합 셀 중복 방지)
                    if (canonical_r, canonical_c) in processed_cells:
                        continue
                    
                    processed_cells.add((canonical_r, canonical_c))
                    val = self._read_number_from_cell(canonical_r, canonical_c)
                    total += val
        
        return total
    
    def _eval_subtotal(self, formula: str, row: int, col: int) -> float:
        """
        SUBTOTAL 함수 계산: =SUBTOTAL(9, A1:A10)
        필터된 행만 합산 (Excel과 동일하게 동작)
        병합 셀은 한 번만 계산 (중복 방지)
        """
        import re
        # =SUBTOTAL(9, A1:A10) 패턴 매칭
        m = re.search(r"SUBTOTAL\s*\(\s*9\s*,\s*([A-Z]{1,3}\d+)\s*:\s*([A-Z]{1,3}\d+)\s*\)", formula, re.IGNORECASE)
        if not m:
            raise ValueError("SUBTOTAL range not found")
        
        start_addr = m.group(1).upper()
        end_addr = m.group(2).upper()
        
        # 주소를 행/열로 변환
        start_row, start_col = self._addr_to_row_col(start_addr)
        end_row, end_col = self._addr_to_row_col(end_addr)
        
        # 범위 내 셀 값 합산 (필터 상태 반영, 병합 셀 중복 방지)
        total = 0.0
        processed_cells = set()  # 이미 처리한 병합 셀 추적
        
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                if 1 <= r <= self.max_row and 1 <= c <= self.max_col:
                    # 병합이면 좌상단으로 정규화
                    canonical_r, canonical_c = self._canonical_cell(r, c)
                    
                    # 이미 처리한 셀이면 스킵 (병합 셀 중복 방지)
                    if (canonical_r, canonical_c) in processed_cells:
                        continue
                    
                    # 필터 상태 확인 (원본 행 기준)
                    if self._is_row_visible(r):
                        processed_cells.add((canonical_r, canonical_c))
                        val = self._read_number_from_cell(canonical_r, canonical_c)
                        total += val
        
        return total
    
    def _addr_to_row_col(self, addr: str) -> Tuple[int, int]:
        """셀 주소(A1)를 (행, 열) 튜플로 변환"""
        import re
        mm = re.fullmatch(r"([A-Z]{1,3})(\d+)", addr.upper())
        if not mm:
            raise ValueError(f"Invalid cell address: {addr}")
        
        col_letters = mm.group(1)
        row = int(mm.group(2))
        
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - 64)
        
        return row, col
    
    def _read_number_from_cell(self, row: int, col: int) -> float:
        """셀에서 숫자값 읽기 (병합 처리 포함)"""
        # 병합이면 좌상단으로 정규화
        row, col = self._canonical_cell(row, col)
        
        vv = self.dirty.get((row, col), self.ws.cell(row=row, column=col).value)
        
        if vv is None:
            return 0.0
        if isinstance(vv, (int, float)):
            return float(vv)
        if isinstance(vv, str):
            # 수식이면 계산 시도
            if vv.strip().startswith("="):
                try:
                    v_display = self._display_value(vv, row, col)
                    if isinstance(v_display, (int, float)):
                        return float(v_display)
                except Exception:
                    pass
            # 문자열 숫자 처리
            t = vv.strip().replace(",", "")
            try:
                return float(t)
            except Exception:
                return 0.0
        return 0.0
    
    def _is_row_visible(self, row: int) -> bool:
        """
        해당 행이 필터에 의해 보이는지 확인
        proxy_model이 없으면 모든 행이 보이는 것으로 간주
        """
        if not self.proxy_model:
            return True
        
        # proxy_model에서 해당 행이 필터 통과하는지 확인
        # source row는 0-based이므로 row - 1
        source_row = row - 1
        if source_row < 0:
            return True
        
        # filterAcceptsRow로 확인
        from PySide6.QtCore import QModelIndex
        return self.proxy_model.filterAcceptsRow(source_row, QModelIndex())

"""
기본 설정 페이지 - 기본 구상률 및 보증 설정 관리
"""
from typing import Dict, Any, Optional, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QTableWidget, QTableWidgetItem, QMessageBox, QDialog, QFormLayout,
    QGroupBox, QLabel, QSpinBox, QDoubleSpinBox, QListWidget, QListWidgetItem
)

from src.database import (
    get_all_common_project_liabilities,
    upsert_common_project_liability,
    delete_common_project_liability,
    get_global_warranty,
    update_global_warranty
)


class LiabilityRatioDialog(QDialog):
    """기본 구상률 추가/수정 다이얼로그"""
    def __init__(self, parent=None, project_code: str = None, liability_ratio: float = None):
        super().__init__(parent)
        self.is_edit_mode = project_code is not None
        
        title = "기본 구상률 수정" if self.is_edit_mode else "기본 구상률 추가"
        self.setWindowTitle(title)
        self.setFixedSize(400, 150)
        
        layout = QFormLayout()
        
        self.project_code_edit = QLineEdit()
        self.project_code_edit.setPlaceholderText("예: L43, H45, ALL")
        if project_code:
            self.project_code_edit.setText(project_code)
            self.project_code_edit.setReadOnly(True)  # 수정 모드에서는 프로젝트 코드 변경 불가
        layout.addRow("프로젝트 코드 *:", self.project_code_edit)
        
        self.liability_ratio_edit = QDoubleSpinBox()
        self.liability_ratio_edit.setRange(0.0, 100.0)
        self.liability_ratio_edit.setSingleStep(0.1)
        self.liability_ratio_edit.setDecimals(2)
        self.liability_ratio_edit.setSuffix(" %")
        if liability_ratio is not None:
            # 저장된 값(0.0 ~ 1.0)을 퍼센티지로 변환
            self.liability_ratio_edit.setValue(liability_ratio * 100)
        layout.addRow("구상률 (%) *:", self.liability_ratio_edit)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        ok_btn = QPushButton("확인")
        cancel_btn = QPushButton("취소")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow("", button_layout)
        
        self.setLayout(layout)
    
    def get_data(self) -> Dict[str, Any]:
        """입력 데이터 반환"""
        # 퍼센티지 입력값을 0.0 ~ 1.0 범위로 변환
        percentage = self.liability_ratio_edit.value()
        liability_ratio = percentage / 100.0
        
        return {
            "project_code": self.project_code_edit.text().strip(),
            "liability_ratio": liability_ratio
        }


class WarrantySettingsWidget(QWidget):
    """기본 보증 설정 관리 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_page = parent  # DefaultSettingsPageWidget 참조
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # 제목
        self.title_label = QLabel("기본 보증 설정")
        self.title_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(self.title_label)
        
        # 설명
        desc_label = QLabel("전역 보증 주행거리 및 보증 기간을 설정합니다.")
        desc_label.setStyleSheet("color: #666;")
        layout.addWidget(desc_label)
        
        # 보증 설정 폼
        form_group = QGroupBox("보증 설정")
        form_layout = QFormLayout()
        form_group.setLayout(form_layout)
        
        self.mileage_spin = QSpinBox()
        self.mileage_spin.setRange(0, 1000000)
        self.mileage_spin.setSuffix(" km")
        form_layout.addRow("보증 주행거리:", self.mileage_spin)
        
        self.period_spin = QSpinBox()
        self.period_spin.setRange(0, 20)
        self.period_spin.setSuffix(" 년")
        form_layout.addRow("보증 기간:", self.period_spin)
        
        layout.addWidget(form_group)
        
        # 저장 버튼
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.btn_save = QPushButton("저장")
        self.btn_save.clicked.connect(self.on_save)
        button_layout.addWidget(self.btn_save)
        layout.addLayout(button_layout)
        
        layout.addStretch()
        self.setLayout(layout)
        
        # 초기 데이터 로드
        self.load_data()
    
    def load_data(self):
        """보증 설정 데이터 로드"""
        mileage, period_years = get_global_warranty()
        self.mileage_spin.setValue(mileage)
        self.period_spin.setValue(period_years)
    
    def on_save(self):
        """보증 설정 저장"""
        mileage = self.mileage_spin.value()
        period_years = self.period_spin.value()
        
        if mileage <= 0:
            QMessageBox.warning(self, "오류", "보증 주행거리는 0보다 커야 합니다.")
            return
        
        if period_years <= 0:
            QMessageBox.warning(self, "오류", "보증 기간은 0보다 커야 합니다.")
            return
        
        try:
            update_global_warranty(mileage, period_years)
            QMessageBox.information(self, "완료", "보증 설정이 저장되었습니다.")
            self.load_data()
            if self.parent_page:
                self.parent_page.refresh_left_panel()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"보증 설정 저장 실패: {str(e)}")


class LiabilityRatioSettingsWidget(QWidget):
    """기본 구상률 설정 관리 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_page = parent  # DefaultSettingsPageWidget 참조
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # 제목
        self.title_label = QLabel("기본 구상률 설정")
        self.title_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(self.title_label)
        
        # 설명
        desc_label = QLabel("프로젝트 코드별 기본 구상률을 관리합니다.")
        desc_label.setStyleSheet("color: #666;")
        layout.addWidget(desc_label)
        
        # 버튼들
        button_layout = QHBoxLayout()
        self.btn_add = QPushButton("+ 추가")
        self.btn_edit = QPushButton("수정")
        self.btn_delete = QPushButton("삭제")
        self.btn_edit.setEnabled(False)
        self.btn_delete.setEnabled(False)
        
        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_edit)
        button_layout.addWidget(self.btn_delete)
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # 테이블
        self.table = QTableWidget()
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        layout.addWidget(self.table, 1)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self.btn_add.clicked.connect(self.on_add)
        self.btn_edit.clicked.connect(self.on_edit)
        self.btn_delete.clicked.connect(self.on_delete)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        
        # 초기 데이터 로드
        self.load_data()
    
    def load_data(self):
        """기본 구상률 데이터 로드"""
        self.liabilities = get_all_common_project_liabilities()
        self.refresh_table()
    
    def refresh_table(self):
        """테이블 새로고침"""
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["프로젝트 코드", "구상률"])
        self.table.setRowCount(len(self.liabilities))
        
        # 헤더 설정 (컬럼 설정 후에 다시 적용)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)  # 마지막 컬럼이 남은 공간을 채우도록
        
        for row, item in enumerate(self.liabilities):
            project_code = item.get("project_code", "")
            liability_ratio = item.get("liability_ratio", 0.0)
            
            # 퍼센트로 표시 (0.6 -> 60.0%)
            percentage = liability_ratio * 100
            
            self.table.setItem(row, 0, QTableWidgetItem(project_code))
            self.table.setItem(row, 1, QTableWidgetItem(f"{percentage:.1f}%"))
            
            # 중앙 정렬
            for col in range(2):
                item_widget = self.table.item(row, col)
                if item_widget:
                    item_widget.setTextAlignment(Qt.AlignCenter)
        
        # 컬럼 너비 자동 조정
        self.table.resizeColumnsToContents()
        
        # 첫 번째 컬럼 최소 너비 보장 (너무 짧아지지 않도록)
        if self.table.columnWidth(0) < 150:
            self.table.setColumnWidth(0, 150)
    
    def on_selection_changed(self):
        """선택 변경 시"""
        has_selection = len(self.table.selectedItems()) > 0
        self.btn_edit.setEnabled(has_selection)
        self.btn_delete.setEnabled(has_selection)
    
    def get_selected_project_code(self) -> Optional[str]:
        """선택된 행의 프로젝트 코드 반환"""
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())
        
        if not selected_rows:
            return None
        
        row = list(selected_rows)[0]
        project_code_item = self.table.item(row, 0)
        if project_code_item:
            return project_code_item.text()
        return None
    
    def on_add(self):
        """기본 구상률 추가"""
        dialog = LiabilityRatioDialog(self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            if not data["project_code"]:
                QMessageBox.warning(self, "오류", "프로젝트 코드를 입력해주세요.")
                return
            
            try:
                upsert_common_project_liability(
                    project_code=data["project_code"],
                    liability_ratio=data["liability_ratio"]
                )
                QMessageBox.information(self, "완료", "기본 구상률이 추가되었습니다.")
                self.load_data()
                if self.parent_page:
                    self.parent_page.refresh_left_panel()
            except Exception as e:
                QMessageBox.critical(self, "오류", f"기본 구상률 추가 실패: {str(e)}")
    
    def on_edit(self):
        """기본 구상률 수정"""
        project_code = self.get_selected_project_code()
        if not project_code:
            QMessageBox.warning(self, "오류", "수정할 항목을 선택해주세요.")
            return
        
        # 현재 구상률 찾기
        current_ratio = None
        for item in self.liabilities:
            if item.get("project_code") == project_code:
                current_ratio = item.get("liability_ratio")
                break
        
        if current_ratio is None:
            QMessageBox.warning(self, "오류", "선택한 항목을 찾을 수 없습니다.")
            return
        
        dialog = LiabilityRatioDialog(self, project_code=project_code, liability_ratio=current_ratio)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            try:
                upsert_common_project_liability(
                    project_code=data["project_code"],
                    liability_ratio=data["liability_ratio"]
                )
                QMessageBox.information(self, "완료", "기본 구상률이 수정되었습니다.")
                self.load_data()
                if self.parent_page:
                    self.parent_page.refresh_left_panel()
            except Exception as e:
                QMessageBox.critical(self, "오류", f"기본 구상률 수정 실패: {str(e)}")
    
    def on_delete(self):
        """기본 구상률 삭제"""
        project_code = self.get_selected_project_code()
        if not project_code:
            QMessageBox.warning(self, "오류", "삭제할 항목을 선택해주세요.")
            return
        
        reply = QMessageBox.question(
            self,
            "삭제 확인",
            f"프로젝트 코드 '{project_code}'의 기본 구상률을 삭제하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                success = delete_common_project_liability(project_code)
                if success:
                    QMessageBox.information(self, "완료", "기본 구상률이 삭제되었습니다.")
                    self.load_data()
                    if self.parent_page:
                        self.parent_page.refresh_left_panel()
                else:
                    QMessageBox.warning(self, "오류", "삭제할 항목을 찾을 수 없습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"기본 구상률 삭제 실패: {str(e)}")


class DefaultSettingsPageWidget(QWidget):
    """기본 설정 페이지"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(8)
        
        # 왼쪽: 메뉴 탭
        left_panel = QVBoxLayout()
        left_panel.setContentsMargins(0, 0, 0, 0)
        left_panel.setSpacing(4)
        
        # 메뉴 리스트
        self.menu_list = QListWidget()
        self.menu_list.setMaximumWidth(250)
        self.menu_list.addItem(QListWidgetItem("기본 구상률 설정"))
        self.menu_list.addItem(QListWidgetItem("기본 보증 설정"))
        self.menu_list.setCurrentRow(0)  # 첫 번째 항목 선택
        left_panel.addWidget(self.menu_list)
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        layout.addWidget(left_widget)
        
        # 오른쪽: 선택한 설정의 상세 관리 위젯
        self.settings_widget = QWidget()
        self.settings_widget_layout = QVBoxLayout()
        self.settings_widget_layout.setContentsMargins(0, 0, 0, 0)
        self.settings_widget.setLayout(self.settings_widget_layout)
        
        # 초기 위젯 (기본 구상률 설정)
        self.current_widget = LiabilityRatioSettingsWidget(self)
        self.current_widget.parent_page = self
        self.settings_widget_layout.addWidget(self.current_widget)
        
        layout.addWidget(self.settings_widget, 1)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self.menu_list.currentRowChanged.connect(self.on_menu_selected)
        
        # 초기 위젯 설정
        self.on_menu_selected(0)
    
    def on_menu_selected(self, index: int):
        """메뉴 선택 시"""
        # 위젯 교체
        self.settings_widget_layout.removeWidget(self.current_widget)
        self.current_widget.deleteLater()
        
        if index == 0:
            # 기본 구상률 설정
            self.current_widget = LiabilityRatioSettingsWidget(self)
        elif index == 1:
            # 기본 보증 설정
            self.current_widget = WarrantySettingsWidget(self)
        else:
            return
        
        self.current_widget.parent_page = self
        self.settings_widget_layout.addWidget(self.current_widget)
    
    def refresh_left_panel(self):
        """왼쪽 패널 새로고침 (위젯에서 데이터 변경 시 호출) - 더 이상 필요 없음"""
        pass

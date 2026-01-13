"""
기본 설정 페이지 - 기본 구상률 / 보증 설정 / 차계-프로젝트 매핑 관리
"""

from __future__ import annotations

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
    update_global_warranty,

    # ✅ 차계-프로젝트 매핑 DB CRUD
    get_all_vehicle_project_maps,
    upsert_vehicle_project_map,
    delete_vehicle_project_map,
)


# =========================================================
# 기본 구상률 Dialog / Widget
# =========================================================

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
            self.project_code_edit.setReadOnly(True)
        layout.addRow("프로젝트 코드 *:", self.project_code_edit)

        self.liability_ratio_edit = QDoubleSpinBox()
        self.liability_ratio_edit.setRange(0.0, 100.0)
        self.liability_ratio_edit.setSingleStep(0.1)
        self.liability_ratio_edit.setDecimals(2)
        self.liability_ratio_edit.setSuffix(" %")
        if liability_ratio is not None:
            self.liability_ratio_edit.setValue(float(liability_ratio) * 100.0)
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
        percentage = float(self.liability_ratio_edit.value())
        liability_ratio = percentage / 100.0
        return {
            "project_code": self.project_code_edit.text().strip(),
            "liability_ratio": liability_ratio
        }


class LiabilityRatioSettingsWidget(QWidget):
    """기본 구상률 설정 관리 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_page = parent

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.title_label = QLabel("기본 구상률 설정")
        self.title_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(self.title_label)

        desc_label = QLabel("프로젝트 코드별 기본 구상률을 관리합니다.")
        desc_label.setStyleSheet("color: #666;")
        layout.addWidget(desc_label)

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

        self.table = QTableWidget()
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        layout.addWidget(self.table, 1)

        self.setLayout(layout)

        self.btn_add.clicked.connect(self.on_add)
        self.btn_edit.clicked.connect(self.on_edit)
        self.btn_delete.clicked.connect(self.on_delete)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

        self.liabilities: List[Dict[str, Any]] = []
        self.load_data()

    def load_data(self):
        self.liabilities = get_all_common_project_liabilities()
        self.refresh_table()

    def refresh_table(self):
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["프로젝트 코드", "구상률"])
        self.table.setRowCount(len(self.liabilities))

        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)

        for row, item in enumerate(self.liabilities):
            project_code = str(item.get("project_code", "") or "")
            liability_ratio = float(item.get("liability_ratio", 0.0) or 0.0)
            percentage = liability_ratio * 100.0

            self.table.setItem(row, 0, QTableWidgetItem(project_code))
            self.table.setItem(row, 1, QTableWidgetItem(f"{percentage:.1f}%"))

            for col in range(2):
                it = self.table.item(row, col)
                if it:
                    it.setTextAlignment(Qt.AlignCenter)

        self.table.resizeColumnsToContents()
        if self.table.columnWidth(0) < 150:
            self.table.setColumnWidth(0, 150)

    def on_selection_changed(self):
        has_selection = len(self.table.selectedItems()) > 0
        self.btn_edit.setEnabled(has_selection)
        self.btn_delete.setEnabled(has_selection)

    def get_selected_project_code(self) -> Optional[str]:
        selected_rows = {it.row() for it in self.table.selectedItems()}
        if not selected_rows:
            return None
        row = next(iter(selected_rows))
        it = self.table.item(row, 0)
        return it.text() if it else None

    def on_add(self):
        dialog = LiabilityRatioDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        if not data["project_code"]:
            QMessageBox.warning(self, "오류", "프로젝트 코드를 입력해주세요.")
            return

        try:
            upsert_common_project_liability(
                project_code=data["project_code"],
                liability_ratio=float(data["liability_ratio"]),
            )
            QMessageBox.information(self, "완료", "기본 구상률이 추가되었습니다.")
            self.load_data()
            if self.parent_page:
                self.parent_page.refresh_left_panel()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"기본 구상률 추가 실패: {str(e)}")

    def on_edit(self):
        project_code = self.get_selected_project_code()
        if not project_code:
            QMessageBox.warning(self, "오류", "수정할 항목을 선택해주세요.")
            return

        current_ratio = None
        for item in self.liabilities:
            if str(item.get("project_code")) == project_code:
                current_ratio = item.get("liability_ratio")
                break

        if current_ratio is None:
            QMessageBox.warning(self, "오류", "선택한 항목을 찾을 수 없습니다.")
            return

        dialog = LiabilityRatioDialog(self, project_code=project_code, liability_ratio=float(current_ratio))
        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        try:
            upsert_common_project_liability(
                project_code=data["project_code"],
                liability_ratio=float(data["liability_ratio"]),
            )
            QMessageBox.information(self, "완료", "기본 구상률이 수정되었습니다.")
            self.load_data()
            if self.parent_page:
                self.parent_page.refresh_left_panel()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"기본 구상률 수정 실패: {str(e)}")

    def on_delete(self):
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
        if reply != QMessageBox.Yes:
            return

        try:
            ok = delete_common_project_liability(project_code)
            if ok:
                QMessageBox.information(self, "완료", "기본 구상률이 삭제되었습니다.")
                self.load_data()
                if self.parent_page:
                    self.parent_page.refresh_left_panel()
            else:
                QMessageBox.warning(self, "오류", "삭제할 항목을 찾을 수 없습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"기본 구상률 삭제 실패: {str(e)}")


# =========================================================
# 기본 보증 Widget
# =========================================================

class WarrantySettingsWidget(QWidget):
    """기본 보증 설정 관리 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_page = parent

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.title_label = QLabel("기본 보증 설정")
        self.title_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(self.title_label)

        desc_label = QLabel("전역 보증 주행거리 및 보증 기간을 설정합니다.")
        desc_label.setStyleSheet("color: #666;")
        layout.addWidget(desc_label)

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

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.btn_save = QPushButton("저장")
        self.btn_save.clicked.connect(self.on_save)
        button_layout.addWidget(self.btn_save)
        layout.addLayout(button_layout)

        layout.addStretch()
        self.setLayout(layout)

        self.load_data()

    def load_data(self):
        mileage, period_years = get_global_warranty()
        self.mileage_spin.setValue(int(mileage))
        self.period_spin.setValue(int(period_years))

    def on_save(self):
        mileage = int(self.mileage_spin.value())
        period_years = int(self.period_spin.value())

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


# =========================================================
# ✅ 차계-프로젝트 매핑 Dialog / Widget
# =========================================================

class VehicleProjectMapDialog(QDialog):
    """
    vehicle_prefix(차계 키) / project_code 입력 다이얼로그

    - 추가 모드: vehicle_prefix 입력 가능
    - 수정 모드: vehicle_prefix는 UNIQUE라서 고치기 귀찮아지니까 ReadOnly로 막고 project_code만 수정
      (vehicle_prefix까지 바꾸고 싶으면 삭제 후 재추가 방식 권장)
    """
    def __init__(
        self,
        parent=None,
        vehicle_prefix: Optional[str] = None,
        project_code: Optional[str] = None,
    ):
        super().__init__(parent)
        self.is_edit_mode = vehicle_prefix is not None

        self.setWindowTitle("차계 설정 수정" if self.is_edit_mode else "차계 설정정 추가")
        self.setFixedSize(420, 160)

        layout = QFormLayout()

        self.vehicle_prefix_edit = QLineEdit()
        self.vehicle_prefix_edit.setPlaceholderText("예: J, H, K 또는 J111, H611, K601")
        if vehicle_prefix:
            self.vehicle_prefix_edit.setText(str(vehicle_prefix))
            self.vehicle_prefix_edit.setReadOnly(True)
        layout.addRow("차계 *:", self.vehicle_prefix_edit)

        self.project_code_edit = QLineEdit()
        self.project_code_edit.setPlaceholderText("예: AR1, AR2, LJL, LFD, HZG")
        if project_code:
            self.project_code_edit.setText(str(project_code))
        layout.addRow("프로젝트 코드 *:", self.project_code_edit)

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

    def get_data(self) -> Dict[str, str]:
        return {
            "vehicle_prefix": self.vehicle_prefix_edit.text().strip(),
            "project_code": self.project_code_edit.text().strip(),
        }


class VehicleProjectMapSettingsWidget(QWidget):
    """차계(vehicle_prefix) -> 프로젝트 코드 매핑 관리 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_page = parent

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        title = QLabel("차계-프로젝트 설정")
        title.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(title)

        desc = QLabel("차계 별 프로젝트 코드 매칭을을 관리합니다. ")
        desc.setStyleSheet("color: #666;")
        layout.addWidget(desc)

        button_layout = QHBoxLayout()
        self.btn_add = QPushButton("+ 추가")
        self.btn_edit = QPushButton("수정")
        self.btn_delete = QPushButton("삭제")
        self.btn_refresh = QPushButton("새로고침")
        self.btn_edit.setEnabled(False)
        self.btn_delete.setEnabled(False)

        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_edit)
        button_layout.addWidget(self.btn_delete)
        button_layout.addWidget(self.btn_refresh)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)

        self.setLayout(layout)

        self.btn_add.clicked.connect(self.on_add)
        self.btn_edit.clicked.connect(self.on_edit)
        self.btn_delete.clicked.connect(self.on_delete)
        self.btn_refresh.clicked.connect(self.load_data)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

        self.rows: List[Dict[str, Any]] = []
        self.load_data()

    def load_data(self):
        self.rows = get_all_vehicle_project_maps()
        self.refresh_table()

    def refresh_table(self):
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["ID", "차계", "프로젝트 코드"])
        self.table.setRowCount(len(self.rows))

        for r, item in enumerate(self.rows):
            vid = str(item.get("id", "") or "")
            vp  = str(item.get("vehicle_prefix", "") or "")
            pc  = str(item.get("project_code", "") or "")

            self.table.setItem(r, 0, QTableWidgetItem(vid))
            self.table.setItem(r, 1, QTableWidgetItem(vp))
            self.table.setItem(r, 2, QTableWidgetItem(pc))

            for c in range(3):
                it = self.table.item(r, c)
                if it:
                    it.setTextAlignment(Qt.AlignCenter)

        # 화면엔 2개만 보이게
        self.table.setColumnHidden(0, True)

        self.table.resizeColumnsToContents()
        if self.table.columnWidth(1) < 150:
            self.table.setColumnWidth(1, 150)
        if self.table.columnWidth(2) < 120:
            self.table.setColumnWidth(2, 120)




    def on_selection_changed(self):
        has_selection = len(self.table.selectedItems()) > 0
        self.btn_edit.setEnabled(has_selection)
        self.btn_delete.setEnabled(has_selection)

    def _get_selected_row_index(self) -> Optional[int]:
        selected_rows = {it.row() for it in self.table.selectedItems()}
        if not selected_rows:
            return None
        return next(iter(selected_rows))

    def _get_selected_id(self) -> Optional[int]:
        idx = self._get_selected_row_index()
        if idx is None:
            return None
        it = self.table.item(idx, 0)
        if not it:
            return None
        try:
            return int(it.text())
        except Exception:
            return None

    def on_add(self):
        dialog = VehicleProjectMapDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        vp = data["vehicle_prefix"].strip().upper()
        pc = data["project_code"].strip().upper()

        if not vp:
            QMessageBox.warning(self, "오류", "차계 키(vehicle_prefix)를 입력해주세요.")
            return
        if not pc:
            QMessageBox.warning(self, "오류", "프로젝트 코드(project_code)를 입력해주세요.")
            return

        try:
            upsert_vehicle_project_map(vp, pc)
            QMessageBox.information(self, "완료", "차계 설정이 저장되었습니다.")
            self.load_data()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"차계 설정 저장 실패: {str(e)}")

    def on_edit(self):
        idx = self._get_selected_row_index()
        if idx is None:
            QMessageBox.warning(self, "오류", "수정할 항목을 선택해주세요.")
            return

        current = self.rows[idx]
        vp = str(current.get("vehicle_prefix", "") or "")
        pc = str(current.get("project_code", "") or "")

        dialog = VehicleProjectMapDialog(self, vehicle_prefix=vp, project_code=pc)
        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        # edit 모드에서는 vp는 readOnly라 동일
        vp2 = data["vehicle_prefix"].strip().upper()
        pc2 = data["project_code"].strip().upper()

        if not vp2 or not pc2:
            QMessageBox.warning(self, "오류", "차계 키/프로젝트 코드는 비워둘 수 없습니다.")
            return

        try:
            upsert_vehicle_project_map(vp2, pc2)
            QMessageBox.information(self, "완료", "차계 설정이 수정되었습니다.")
            self.load_data()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"차계 설정 수정 실패: {str(e)}")

    def on_delete(self):
        id_ = self._get_selected_id()
        if id_ is None:
            QMessageBox.warning(self, "오류", "삭제할 항목을 선택해주세요.")
            return

        idx = self._get_selected_row_index()
        vp = ""
        pc = ""
        if idx is not None and idx < len(self.rows):
            vp = str(self.rows[idx].get("vehicle_prefix", "") or "")
            pc = str(self.rows[idx].get("project_code", "") or "")

        reply = QMessageBox.question(
            self,
            "삭제 확인",
            f"삭제하시겠습니까?\n(ID={id_}, vehicle_prefix={vp}, project_code={pc})",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        try:
            ok = delete_vehicle_project_map(id_)
            if ok:
                QMessageBox.information(self, "완료", "차계 설정이 삭제되었습니다.")
                self.load_data()
            else:
                QMessageBox.warning(self, "오류", "삭제할 항목을 찾을 수 없습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"차계 설정 삭제 실패: {str(e)}")


# =========================================================
# 기본 설정 페이지 (메뉴 3개)
# =========================================================

class DefaultSettingsPageWidget(QWidget):
    """기본 설정 페이지"""
    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QHBoxLayout()
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(8)

        left_panel = QVBoxLayout()
        left_panel.setContentsMargins(0, 0, 0, 0)
        left_panel.setSpacing(4)

        self.menu_list = QListWidget()
        self.menu_list.setMaximumWidth(250)

        self.menu_list.addItem(QListWidgetItem("기본 구상률 설정"))
        self.menu_list.addItem(QListWidgetItem("기본 보증 설정"))
        self.menu_list.addItem(QListWidgetItem("차계-프로젝트 설정"))  # ✅ 추가

        self.menu_list.setCurrentRow(0)
        left_panel.addWidget(self.menu_list)

        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        layout.addWidget(left_widget)

        self.settings_widget = QWidget()
        self.settings_widget_layout = QVBoxLayout()
        self.settings_widget_layout.setContentsMargins(0, 0, 0, 0)
        self.settings_widget.setLayout(self.settings_widget_layout)

        self.current_widget: QWidget = LiabilityRatioSettingsWidget(self)
        self.current_widget.parent_page = self
        self.settings_widget_layout.addWidget(self.current_widget)

        layout.addWidget(self.settings_widget, 1)
        self.setLayout(layout)

        self.menu_list.currentRowChanged.connect(self.on_menu_selected)
        self.on_menu_selected(0)

    def on_menu_selected(self, index: int):
        self.settings_widget_layout.removeWidget(self.current_widget)
        self.current_widget.deleteLater()

        if index == 0:
            self.current_widget = LiabilityRatioSettingsWidget(self)
        elif index == 1:
            self.current_widget = WarrantySettingsWidget(self)
        elif index == 2:
            self.current_widget = VehicleProjectMapSettingsWidget(self)  # ✅ 추가
        else:
            return

        self.current_widget.parent_page = self
        self.settings_widget_layout.addWidget(self.current_widget)

    def refresh_left_panel(self):
        # 지금 구조에선 별도 처리 없음(호출만 유지)
        pass

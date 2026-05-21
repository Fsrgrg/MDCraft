"""模板面板：模板列表 + 操作按钮"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QListWidget, QListWidgetItem,
    QPushButton, QHBoxLayout, QLabel, QFileDialog, QMessageBox, QInputDialog
)
from PyQt6.QtCore import pyqtSignal
from app.templates.manager import TemplateManager
from app.models.config import FormatConfig


class TemplatePanel(QWidget):
    template_selected = pyqtSignal(FormatConfig)
    template_list_changed = pyqtSignal()
    save_requested = pyqtSignal()  # 保存前发出，让主窗口同步设置页面

    def __init__(self, tm: TemplateManager, parent=None):
        super().__init__(parent)
        self._tm = tm
        self._current_config: FormatConfig | None = None
        self._setup_ui()
        self._refresh_list()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        title = QLabel('模板列表')
        title.setStyleSheet('font-weight: bold; font-size: 13px;')
        layout.addWidget(title)

        self._list = QListWidget()
        self._list.currentItemChanged.connect(self._on_select)
        layout.addWidget(self._list, 1)

        btn_layout = QHBoxLayout()
        btn_new = QPushButton('新建')
        btn_new.clicked.connect(self._on_new)
        btn_save = QPushButton('保存')
        btn_save.clicked.connect(self._on_save)
        btn_del = QPushButton('删除')
        btn_del.clicked.connect(self._on_delete)
        btn_layout.addWidget(btn_new)
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_del)
        layout.addLayout(btn_layout)

        btn_layout2 = QHBoxLayout()
        btn_import = QPushButton('导入')
        btn_import.clicked.connect(self._on_import)
        btn_export = QPushButton('导出')
        btn_export.clicked.connect(self._on_export)
        btn_dup = QPushButton('复制')
        btn_dup.clicked.connect(self._on_duplicate)
        btn_layout2.addWidget(btn_import)
        btn_layout2.addWidget(btn_export)
        btn_layout2.addWidget(btn_dup)
        layout.addLayout(btn_layout2)

    def _refresh_list(self, select_name: str = None):
        self._list.blockSignals(True)
        self._list.clear()
        for name in self._tm.list_templates():
            item = QListWidgetItem(name)
            self._list.addItem(item)
            if select_name and name == select_name:
                self._list.setCurrentItem(item)
        self._list.blockSignals(False)
        self.template_list_changed.emit()

    def _on_select(self, current, previous):
        if current is None:
            return
        name = current.text()
        try:
            self._current_config = self._tm.load(name)
            self.template_selected.emit(self._current_config)
        except Exception as e:
            QMessageBox.warning(self, '加载失败', str(e))

    def _on_new(self):
        name, ok = QInputDialog.getText(self, '新建模板', '模板名称：')
        if not ok or not name.strip():
            return
        name = name.strip()
        if self._tm.exists(name):
            QMessageBox.warning(self, '名称冲突', f'模板 "{name}" 已存在。')
            return
        config = FormatConfig()
        config.name = name
        self._tm.save(config)
        self._refresh_list(select_name=name)

    def _on_save(self):
        if self._current_config is None:
            QMessageBox.information(self, '提示', '请先选择一个模板。')
            return
        self.save_requested.emit()
        self._tm.save(self._current_config)
        self._refresh_list(select_name=self._current_config.name)

    def _on_delete(self):
        item = self._list.currentItem()
        if item is None:
            return
        name = item.text()
        reply = QMessageBox.question(self, '确认删除', f'确定要删除模板 "{name}" 吗？',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._tm.delete(name)
        self._current_config = None
        self._refresh_list()

    def _on_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '导入模板', '', '模板文件 (*.hputpl);;JSON 文件 (*.json);;所有文件 (*)')
        if not path:
            return
        try:
            config = self._tm.import_from(path)
            self._refresh_list(select_name=config.name)
            QMessageBox.information(self, '导入成功', f'模板 "{config.name}" 已导入。')
        except Exception as e:
            QMessageBox.warning(self, '导入失败', str(e))

    def _on_export(self):
        if self._current_config is None:
            QMessageBox.information(self, '提示', '请先选择一个模板。')
            return
        path, _ = QFileDialog.getSaveFileName(
            self, '导出模板', f'{self._current_config.name}.hputpl',
            '模板文件 (*.hputpl);;JSON 文件 (*.json)')
        if not path:
            return
        self._tm.export_to(self._current_config.name, path)
        QMessageBox.information(self, '导出成功', f'模板已导出到 {path}。')

    def _on_duplicate(self):
        if self._current_config is None:
            QMessageBox.information(self, '提示', '请先选择一个模板。')
            return
        name, ok = QInputDialog.getText(self, '复制模板', '新模板名称：',
                                        text=self._current_config.name + '_副本')
        if not ok or not name.strip():
            return
        name = name.strip()
        self.save_requested.emit()
        self._tm.save(self._current_config)
        new_path = self._tm.duplicate(self._current_config.name, name)
        if new_path is None:
            QMessageBox.warning(self, '复制失败', '名称冲突或源模板不存在。')
            return
        self._refresh_list(select_name=name)

    def get_current_config(self) -> FormatConfig | None:
        return self._current_config

    def set_current_config(self, config: FormatConfig):
        self._current_config = config

"""目录设置"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QScrollArea, QGroupBox, QFormLayout,
    QCheckBox, QComboBox
)
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class TOCSettingsPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._cfg: FormatConfig | None = None
        self._setup_ui()

    def _setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        w = QWidget()
        layout = QVBoxLayout(w)

        # 基本设置
        gb = QGroupBox('目录选项')
        form = QFormLayout(gb)
        self.auto_insert = QCheckBox('自动插入目录')
        self.auto_insert.setChecked(True)
        form.addRow(self.auto_insert)
        self.levels = QComboBox()
        self.levels.addItems(['1-3', '1-2', '1-4', '1-1'])
        form.addRow('包含标题层级：', self.levels)
        layout.addWidget(gb)

        self.heading_editor = ElementStyleEditor('"目录" 标题')
        layout.addWidget(self.heading_editor)

        self.item_editor = ElementStyleEditor('目录条目')
        layout.addWidget(self.item_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.auto_insert.setChecked(config.toc.auto_insert)
        self.levels.setCurrentText(config.toc.heading_levels)
        self.heading_editor.load_style(config.toc.toc_heading_style)
        self.item_editor.load_style(config.toc.toc_item_style)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.toc.auto_insert = self.auto_insert.isChecked()
        self._cfg.toc.heading_levels = self.levels.currentText()
        self._cfg.toc.toc_heading_style = self.heading_editor.save_style()
        self._cfg.toc.toc_item_style = self.item_editor.save_style()

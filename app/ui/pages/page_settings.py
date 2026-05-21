"""页面设置"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QComboBox, QDoubleSpinBox
)
from app.models.config import FormatConfig


class PageSettingsPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._cfg: FormatConfig | None = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        gb = QGroupBox('纸张与边距')
        form = QFormLayout(gb)

        self.paper_type = QComboBox()
        self.paper_type.addItems(['A4 (210×297 mm)', '自定义'])
        self.paper_type.currentIndexChanged.connect(self._on_paper_change)
        form.addRow('纸张类型：', self.paper_type)

        self.page_width = QDoubleSpinBox()
        self.page_width.setRange(10.0, 50.0)
        self.page_width.setSuffix(' cm')
        self.page_width.setValue(21.0)
        self.page_width.valueChanged.connect(self._on_change)
        form.addRow('页面宽度：', self.page_width)

        self.page_height = QDoubleSpinBox()
        self.page_height.setRange(10.0, 50.0)
        self.page_height.setSuffix(' cm')
        self.page_height.setValue(29.7)
        self.page_height.valueChanged.connect(self._on_change)
        form.addRow('页面高度：', self.page_height)

        self.top_margin = QDoubleSpinBox()
        self.top_margin.setRange(0.0, 10.0)
        self.top_margin.setSuffix(' cm')
        self.top_margin.valueChanged.connect(self._on_change)
        form.addRow('上边距：', self.top_margin)

        self.bottom_margin = QDoubleSpinBox()
        self.bottom_margin.setRange(0.0, 10.0)
        self.bottom_margin.setSuffix(' cm')
        self.bottom_margin.valueChanged.connect(self._on_change)
        form.addRow('下边距：', self.bottom_margin)

        self.left_margin = QDoubleSpinBox()
        self.left_margin.setRange(0.0, 10.0)
        self.left_margin.setSuffix(' cm')
        self.left_margin.valueChanged.connect(self._on_change)
        form.addRow('左边距：', self.left_margin)

        self.right_margin = QDoubleSpinBox()
        self.right_margin.setRange(0.0, 10.0)
        self.right_margin.setSuffix(' cm')
        self.right_margin.valueChanged.connect(self._on_change)
        form.addRow('右边距：', self.right_margin)

        layout.addWidget(gb)
        layout.addStretch()

    def _on_paper_change(self, idx):
        if idx == 0:
            self.page_width.setValue(21.0)
            self.page_height.setValue(29.7)

    def _on_change(self):
        if self._cfg is None:
            return
        self._cfg.page.page_width = self.page_width.value()
        self._cfg.page.page_height = self.page_height.value()
        self._cfg.page.top_margin = self.top_margin.value()
        self._cfg.page.bottom_margin = self.bottom_margin.value()
        self._cfg.page.left_margin = self.left_margin.value()
        self._cfg.page.right_margin = self.right_margin.value()

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        pg = config.page
        self.page_width.blockSignals(True)
        self.page_height.blockSignals(True)
        self.top_margin.blockSignals(True)
        self.bottom_margin.blockSignals(True)
        self.left_margin.blockSignals(True)
        self.right_margin.blockSignals(True)
        self.page_width.setValue(pg.page_width)
        self.page_height.setValue(pg.page_height)
        self.top_margin.setValue(pg.top_margin)
        self.bottom_margin.setValue(pg.bottom_margin)
        self.left_margin.setValue(pg.left_margin)
        self.right_margin.setValue(pg.right_margin)
        self.page_width.blockSignals(False)
        self.page_height.blockSignals(False)
        self.top_margin.blockSignals(False)
        self.bottom_margin.blockSignals(False)
        self.left_margin.blockSignals(False)
        self.right_margin.blockSignals(False)

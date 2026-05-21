"""三线表样式设置"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QScrollArea, QGroupBox, QFormLayout,
    QComboBox, QDoubleSpinBox
)
from app.models.config import FormatConfig
from app.ui.pages.element_editor import CN_FONTS, EN_FONTS, ALIGNMENTS, ALIGN_REVERSE


class TableStyleSettingsPage(QWidget):
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

        # 字体设置
        gb_font = QGroupBox('表格文字格式')
        form = QFormLayout(gb_font)

        self.cn_font = QComboBox()
        self.cn_font.addItems(CN_FONTS)
        self.cn_font.setEditable(True)
        form.addRow('中文字体：', self.cn_font)

        self.en_font = QComboBox()
        self.en_font.addItems(EN_FONTS)
        self.en_font.setEditable(True)
        form.addRow('英文字体：', self.en_font)

        self.font_size = QDoubleSpinBox()
        self.font_size.setRange(5.0, 72.0)
        self.font_size.setSuffix(' pt')
        self.font_size.setSingleStep(0.5)
        form.addRow('字号：', self.font_size)

        self.alignment = QComboBox()
        self.alignment.addItems(list(ALIGNMENTS.keys()))
        form.addRow('对齐方式：', self.alignment)

        self.line_spacing = QDoubleSpinBox()
        self.line_spacing.setRange(0.5, 10.0)
        self.line_spacing.setSingleStep(0.1)
        form.addRow('行距：', self.line_spacing)

        layout.addWidget(gb_font)

        # 三线表边框
        gb_border = QGroupBox('三线表边框')
        form2 = QFormLayout(gb_border)

        self.top_width = QDoubleSpinBox()
        self.top_width.setRange(0.5, 6.0)
        self.top_width.setSingleStep(0.25)
        self.top_width.setSuffix(' pt')
        self.top_width.setValue(1.5)
        form2.addRow('顶线宽度：', self.top_width)

        self.bottom_width = QDoubleSpinBox()
        self.bottom_width.setRange(0.5, 6.0)
        self.bottom_width.setSingleStep(0.25)
        self.bottom_width.setSuffix(' pt')
        self.bottom_width.setValue(1.5)
        form2.addRow('底线宽度：', self.bottom_width)

        self.header_width = QDoubleSpinBox()
        self.header_width.setRange(0.25, 3.0)
        self.header_width.setSingleStep(0.25)
        self.header_width.setSuffix(' pt')
        self.header_width.setValue(0.75)
        form2.addRow('表头下线宽度：', self.header_width)

        layout.addWidget(gb_border)
        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        ts = config.table_style
        self.cn_font.setCurrentText(ts.cn_font)
        self.en_font.setCurrentText(ts.en_font)
        self.font_size.setValue(ts.font_size)
        self.alignment.setCurrentText(ALIGN_REVERSE.get(ts.alignment, '居中'))
        self.line_spacing.setValue(ts.line_spacing)
        self.top_width.setValue(ts.top_border_width)
        self.bottom_width.setValue(ts.bottom_border_width)
        self.header_width.setValue(ts.header_bottom_width)

    def save_to_config(self):
        if self._cfg is None:
            return
        ts = self._cfg.table_style
        ts.cn_font = self.cn_font.currentText()
        ts.en_font = self.en_font.currentText()
        ts.font_size = self.font_size.value()
        ts.alignment = ALIGNMENTS.get(self.alignment.currentText(), 'center')
        ts.line_spacing = self.line_spacing.value()
        ts.top_border_width = self.top_width.value()
        ts.bottom_border_width = self.bottom_width.value()
        ts.header_bottom_width = self.header_width.value()

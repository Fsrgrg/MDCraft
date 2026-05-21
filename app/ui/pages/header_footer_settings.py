"""页眉页脚设置"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QScrollArea, QGroupBox, QFormLayout,
    QComboBox, QDoubleSpinBox, QCheckBox, QLineEdit
)
from app.models.config import FormatConfig
from app.ui.pages.element_editor import CN_FONTS, EN_FONTS


class HeaderFooterSettingsPage(QWidget):
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

        # 页眉文字
        gb_header = QGroupBox('页眉内容')
        form = QFormLayout(gb_header)

        self.odd_source = QComboBox()
        self.odd_source.addItems(['所在章节标题（自动）', '固定文字'])
        form.addRow('奇数页页眉：', self.odd_source)

        self.odd_fixed = QLineEdit()
        self.odd_fixed.setPlaceholderText('输入固定页眉文字...')
        form.addRow('奇数页固定文字：', self.odd_fixed)

        self.even_text = QLineEdit()
        form.addRow('偶数页页眉：', self.even_text)

        self.cn_font = QComboBox()
        self.cn_font.addItems(CN_FONTS)
        self.cn_font.setEditable(True)
        form.addRow('页眉中文字体：', self.cn_font)

        self.en_font = QComboBox()
        self.en_font.addItems(EN_FONTS)
        self.en_font.setEditable(True)
        form.addRow('页眉英文字体：', self.en_font)

        self.font_size = QDoubleSpinBox()
        self.font_size.setRange(5.0, 24.0)
        self.font_size.setSuffix(' pt')
        form.addRow('页眉字号：', self.font_size)

        self.header_line = QCheckBox('页眉下方横线')
        self.header_line.setChecked(True)
        form.addRow(self.header_line)

        layout.addWidget(gb_header)

        # 页码设置
        gb_page = QGroupBox('页码设置')
        form2 = QFormLayout(gb_page)

        self.page_num_fmt = QComboBox()
        self.page_num_fmt.addItems(['阿拉伯数字 (1,2,3...)', '大写罗马数字 (Ⅰ,Ⅱ,Ⅲ...)', '小写罗马数字 (ⅰ,ⅱ,ⅲ...)'])
        form2.addRow('正文页码格式：', self.page_num_fmt)

        layout.addWidget(gb_page)
        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        hf = config.header_footer
        self.odd_source.setCurrentIndex(0 if hf.odd_header_source == 'chapter' else 1)
        self.odd_fixed.setText(hf.odd_header_fixed_text)
        self.even_text.setText(hf.even_header_text)
        self.cn_font.setCurrentText(hf.cn_font)
        self.en_font.setCurrentText(hf.en_font)
        self.font_size.setValue(hf.font_size)
        self.header_line.setChecked(hf.header_line)
        fmt_map = {'decimal': 0, 'upperRoman': 1, 'lowerRoman': 2}
        self.page_num_fmt.setCurrentIndex(fmt_map.get(hf.page_num_format, 0))

    def save_to_config(self):
        if self._cfg is None:
            return
        hf = self._cfg.header_footer
        hf.odd_header_source = 'chapter' if self.odd_source.currentIndex() == 0 else 'fixed'
        hf.odd_header_fixed_text = self.odd_fixed.text()
        hf.even_header_text = self.even_text.text()
        hf.cn_font = self.cn_font.currentText()
        hf.en_font = self.en_font.currentText()
        hf.font_size = self.font_size.value()
        hf.header_line = self.header_line.isChecked()
        fmt_map = {0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman'}
        hf.page_num_format = fmt_map.get(self.page_num_fmt.currentIndex(), 'decimal')

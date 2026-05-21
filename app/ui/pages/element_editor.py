"""可复用的 ElementStyle 编辑组件"""

from PyQt6.QtWidgets import (
    QGroupBox, QFormLayout, QComboBox, QDoubleSpinBox,
    QCheckBox, QHBoxLayout, QWidget, QLabel
)
from PyQt6.QtCore import Qt
from app.models.config import ElementStyle

CN_FONTS = ['宋体', '黑体', '楷体', '仿宋', '微软雅黑', '方正书宋', '华文楷体']
EN_FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Cambria Math']
ALIGNMENTS = {'左对齐': 'left', '居中': 'center', '右对齐': 'right', '两端对齐': 'justify'}
ALIGN_REVERSE = {v: k for k, v in ALIGNMENTS.items()}


class ElementStyleEditor(QGroupBox):
    def __init__(self, title: str = '', parent=None):
        super().__init__(title, parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QFormLayout(self)

        # 中文字体
        self.cn_font = QComboBox()
        self.cn_font.addItems(CN_FONTS)
        self.cn_font.setEditable(True)
        layout.addRow('中文字体：', self.cn_font)

        # 英文字体
        self.en_font = QComboBox()
        self.en_font.addItems(EN_FONTS)
        self.en_font.setEditable(True)
        layout.addRow('英文字体：', self.en_font)

        # 字号
        self.font_size = QDoubleSpinBox()
        self.font_size.setRange(5.0, 72.0)
        self.font_size.setSuffix(' pt')
        self.font_size.setSingleStep(0.5)
        layout.addRow('字号：', self.font_size)

        # 加粗
        self.bold = QCheckBox()
        layout.addRow('加粗：', self.bold)

        # 对齐
        self.alignment = QComboBox()
        self.alignment.addItems(list(ALIGNMENTS.keys()))
        layout.addRow('对齐方式：', self.alignment)

        # 行距
        row = QHBoxLayout()
        self.line_spacing = QDoubleSpinBox()
        self.line_spacing.setRange(0.5, 10.0)
        self.line_spacing.setSingleStep(0.1)
        self.line_spacing.setDecimals(1)
        row.addWidget(self.line_spacing)
        self.line_spacing_rule = QComboBox()
        self.line_spacing_rule.addItems(['多倍行距', '固定值(pt)'])
        row.addWidget(self.line_spacing_rule)
        row.addStretch()
        layout.addRow('行距：', row)

        # 首行缩进
        indent_row = QHBoxLayout()
        self.indent_enabled = QCheckBox('启用')
        self.indent_value = QDoubleSpinBox()
        self.indent_value.setRange(0.0, 5.0)
        self.indent_value.setSingleStep(0.1)
        self.indent_value.setSuffix(' cm')
        self.indent_value.setValue(0.74)
        self.indent_enabled.toggled.connect(self.indent_value.setEnabled)
        indent_row.addWidget(self.indent_enabled)
        indent_row.addWidget(self.indent_value)
        indent_row.addStretch()
        layout.addRow('首行缩进：', indent_row)

        # 段前距
        self.space_before = QDoubleSpinBox()
        self.space_before.setRange(0, 100)
        self.space_before.setSuffix(' pt')
        layout.addRow('段前距：', self.space_before)

        # 段后距
        self.space_after = QDoubleSpinBox()
        self.space_after.setRange(0, 100)
        self.space_after.setSuffix(' pt')
        layout.addRow('段后距：', self.space_after)

    def load_style(self, es: ElementStyle):
        self.cn_font.setCurrentText(es.cn_font)
        self.en_font.setCurrentText(es.en_font)
        self.font_size.setValue(es.font_size)
        self.bold.setChecked(es.bold)
        key = ALIGN_REVERSE.get(es.alignment, '左对齐')
        self.alignment.setCurrentText(key)
        self.line_spacing.setValue(es.line_spacing)
        rule_text = '多倍行距' if es.line_spacing_rule == 'multiple' else '固定值(pt)'
        self.line_spacing_rule.setCurrentText(rule_text)
        if es.first_line_indent is not None:
            self.indent_enabled.setChecked(True)
            self.indent_value.setValue(es.first_line_indent)
        else:
            self.indent_enabled.setChecked(False)
        self.space_before.setValue(es.space_before)
        self.space_after.setValue(es.space_after)

    def save_style(self) -> ElementStyle:
        rule = 'multiple' if self.line_spacing_rule.currentText() == '多倍行距' else 'exact'
        indent = self.indent_value.value() if self.indent_enabled.isChecked() else None
        return ElementStyle(
            cn_font=self.cn_font.currentText(),
            en_font=self.en_font.currentText(),
            font_size=self.font_size.value(),
            bold=self.bold.isChecked(),
            alignment=ALIGNMENTS.get(self.alignment.currentText(), 'left'),
            line_spacing=self.line_spacing.value(),
            line_spacing_rule=rule,
            first_line_indent=indent,
            space_before=self.space_before.value(),
            space_after=self.space_after.value(),
        )

"""题目页设置"""

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class TitleSettingsPage(QWidget):
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

        self.uni_editor = ElementStyleEditor('学校名称')
        self.uni_editor.cn_font.setCurrentText('黑体')
        self.uni_editor.font_size.setValue(36.0)
        layout.addWidget(self.uni_editor)

        self.doctype_editor = ElementStyleEditor('文档类型（如"本科毕业设计（论文）"）')
        self.doctype_editor.cn_font.setCurrentText('黑体')
        self.doctype_editor.font_size.setValue(22.0)
        layout.addWidget(self.doctype_editor)

        self.paper_editor = ElementStyleEditor('论文题目')
        self.paper_editor.cn_font.setCurrentText('黑体')
        self.paper_editor.font_size.setValue(18.0)
        layout.addWidget(self.paper_editor)

        self.date_editor = ElementStyleEditor('日期/院系信息')
        self.date_editor.cn_font.setCurrentText('宋体')
        self.date_editor.font_size.setValue(15.0)
        layout.addWidget(self.date_editor)

        self.table_editor = ElementStyleEditor('题目页表格')
        self.table_editor.cn_font.setCurrentText('宋体')
        self.table_editor.font_size.setValue(15.0)
        layout.addWidget(self.table_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.uni_editor.load_style(config.title_uni)
        self.doctype_editor.load_style(config.title_doctype)
        self.paper_editor.load_style(config.title_paper)
        self.date_editor.load_style(config.title_date)
        self.table_editor.load_style(config.title_table)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.title_uni = self.uni_editor.save_style()
        self._cfg.title_doctype = self.doctype_editor.save_style()
        self._cfg.title_paper = self.paper_editor.save_style()
        self._cfg.title_date = self.date_editor.save_style()
        self._cfg.title_table = self.table_editor.save_style()

"""英文摘要设置"""

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class EnAbstractSettingsPage(QWidget):
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

        self.label_editor = ElementStyleEditor('"ABSTRACT" 标题')
        layout.addWidget(self.label_editor)

        self.body_editor = ElementStyleEditor('英文摘要正文')
        layout.addWidget(self.body_editor)

        self.kw_label_editor = ElementStyleEditor('"KEY WORDS" 标签')
        layout.addWidget(self.kw_label_editor)

        self.kw_body_editor = ElementStyleEditor('英文关键词内容')
        layout.addWidget(self.kw_body_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.label_editor.load_style(config.en_abstract_label)
        self.body_editor.load_style(config.en_abstract_body)
        self.kw_label_editor.load_style(config.en_keywords_label)
        self.kw_body_editor.load_style(config.en_keywords_body)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.en_abstract_label = self.label_editor.save_style()
        self._cfg.en_abstract_body = self.body_editor.save_style()
        self._cfg.en_keywords_label = self.kw_label_editor.save_style()
        self._cfg.en_keywords_body = self.kw_body_editor.save_style()

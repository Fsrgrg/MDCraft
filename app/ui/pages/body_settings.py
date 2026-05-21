"""正文、公式、源代码、注释设置"""

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class BodySettingsPage(QWidget):
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

        self.body_editor = ElementStyleEditor('正文段落')
        layout.addWidget(self.body_editor)

        self.formula_editor = ElementStyleEditor('公式')
        layout.addWidget(self.formula_editor)

        self.src_editor = ElementStyleEditor('源代码')
        layout.addWidget(self.src_editor)

        self.note_editor = ElementStyleEditor('文末注释')
        layout.addWidget(self.note_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.body_editor.load_style(config.body)
        self.formula_editor.load_style(config.formula)
        self.src_editor.load_style(config.source_code)
        self.note_editor.load_style(config.end_note)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.body = self.body_editor.save_style()
        self._cfg.formula = self.formula_editor.save_style()
        self._cfg.source_code = self.src_editor.save_style()
        self._cfg.end_note = self.note_editor.save_style()

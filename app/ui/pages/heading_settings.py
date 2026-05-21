"""正文标题设置 H1-H4"""

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class HeadingSettingsPage(QWidget):
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

        self.h1_editor = ElementStyleEditor('一级标题 (Heading 1) — 章标题')
        layout.addWidget(self.h1_editor)

        self.h2_editor = ElementStyleEditor('二级标题 (Heading 2) — 节标题')
        layout.addWidget(self.h2_editor)

        self.h3_editor = ElementStyleEditor('三级标题 (Heading 3)')
        layout.addWidget(self.h3_editor)

        self.h4_editor = ElementStyleEditor('四级标题 (Heading 4)')
        layout.addWidget(self.h4_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.h1_editor.load_style(config.heading_1)
        self.h2_editor.load_style(config.heading_2)
        self.h3_editor.load_style(config.heading_3)
        self.h4_editor.load_style(config.heading_4)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.heading_1 = self.h1_editor.save_style()
        self._cfg.heading_2 = self.h2_editor.save_style()
        self._cfg.heading_3 = self.h3_editor.save_style()
        self._cfg.heading_4 = self.h4_editor.save_style()

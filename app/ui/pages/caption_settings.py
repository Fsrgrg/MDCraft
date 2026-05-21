"""图表题注设置"""

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class CaptionSettingsPage(QWidget):
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

        self.table_cap_editor = ElementStyleEditor('表题注（如"表 1 实验结果"）')
        layout.addWidget(self.table_cap_editor)

        self.figure_cap_editor = ElementStyleEditor('图题注（如"图 1 系统架构"）')
        layout.addWidget(self.figure_cap_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.table_cap_editor.load_style(config.table_caption)
        self.figure_cap_editor.load_style(config.figure_caption)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.table_caption = self.table_cap_editor.save_style()
        self._cfg.figure_caption = self.figure_cap_editor.save_style()

"""参考文献 + 附录设置"""

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from app.models.config import FormatConfig
from app.ui.pages.element_editor import ElementStyleEditor


class ReferenceSettingsPage(QWidget):
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

        self.ref_heading_editor = ElementStyleEditor('"参考文献" 标题')
        layout.addWidget(self.ref_heading_editor)

        self.ref_item_editor = ElementStyleEditor('参考文献条目')
        layout.addWidget(self.ref_item_editor)

        self.appendix_heading_editor = ElementStyleEditor('"附录" 大标题')
        layout.addWidget(self.appendix_heading_editor)

        self.appendix_sub_editor = ElementStyleEditor('附录子标题')
        layout.addWidget(self.appendix_sub_editor)

        layout.addStretch()
        scroll.setWidget(w)

    def set_config(self, config: FormatConfig):
        self._cfg = config

    def refresh(self, config: FormatConfig):
        self._cfg = config
        self.ref_heading_editor.load_style(config.ref_heading)
        self.ref_item_editor.load_style(config.ref_item)
        self.appendix_heading_editor.load_style(config.appendix_heading)
        self.appendix_sub_editor.load_style(config.appendix_sub)

    def save_to_config(self):
        if self._cfg is None:
            return
        self._cfg.ref_heading = self.ref_heading_editor.save_style()
        self._cfg.ref_item = self.ref_item_editor.save_style()
        self._cfg.appendix_heading = self.appendix_heading_editor.save_style()
        self._cfg.appendix_sub = self.appendix_sub_editor.save_style()

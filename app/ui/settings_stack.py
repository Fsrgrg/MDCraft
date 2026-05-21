"""设置页面框架：左侧分类导航 + 右侧 QStackedWidget"""

from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QListWidget, QListWidgetItem, QStackedWidget
)
from app.models.config import FormatConfig

CATEGORIES = [
    ('页面设置', 'page'),
    ('题目页', 'title'),
    ('中文摘要', 'abstract'),
    ('英文摘要', 'en_abstract'),
    ('正文标题', 'headings'),
    ('正文段落', 'body'),
    ('图表题注', 'captions'),
    ('参考文献与附录', 'references'),
    ('目录', 'toc'),
    ('表格样式', 'table_style'),
    ('页眉页脚', 'header_footer'),
]


class SettingsStack(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._pages = {}
        self._config: FormatConfig | None = None
        self._current_key: str = ''
        self._setup_ui()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._nav = QListWidget()
        self._nav.setFixedWidth(130)
        self._nav.setStyleSheet('font-size: 12px;')
        for label, key in CATEGORIES:
            item = QListWidgetItem(label)
            item.setData(1, key)
            self._nav.addItem(item)
        self._nav.currentRowChanged.connect(self._on_nav_change)
        layout.addWidget(self._nav)

        self._stack = QStackedWidget()
        layout.addWidget(self._stack, 1)

    def register_page(self, key: str, page: QWidget):
        self._pages[key] = page
        self._stack.addWidget(page)

    def set_config(self, config: FormatConfig):
        self._config = config
        for page in self._pages.values():
            if hasattr(page, 'set_config'):
                page.set_config(config)

    def save_current_page(self):
        """保存当前页面修改到 config"""
        if self._current_key and self._current_key in self._pages:
            page = self._pages[self._current_key]
            if hasattr(page, 'save_to_config'):
                page.save_to_config()

    def save_all_pages(self):
        """保存所有页面修改到 config"""
        for page in self._pages.values():
            if hasattr(page, 'save_to_config'):
                page.save_to_config()

    def _on_nav_change(self, index: int):
        item = self._nav.item(index)
        if item is None:
            return
        key = item.data(1)
        if key == self._current_key:
            return
        # 离开当前页面前保存
        self.save_current_page()
        self._current_key = key
        if key in self._pages:
            page = self._pages[key]
            if hasattr(page, 'refresh') and self._config:
                page.refresh(self._config)
            self._stack.setCurrentWidget(page)

    def show_page(self, key: str):
        for i in range(self._nav.count()):
            item = self._nav.item(i)
            if item is not None and item.data(1) == key:
                self._nav.setCurrentRow(i)
                break

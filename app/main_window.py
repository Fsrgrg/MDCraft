"""主窗口"""

import os
from PyQt6.QtWidgets import (
    QMainWindow, QToolBar, QStatusBar, QSplitter, QFileDialog,
    QMessageBox, QLabel, QComboBox, QPushButton, QWidget, QVBoxLayout,
    QHBoxLayout
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QAction, QIcon

from app.models.config import FormatConfig
from app.templates.manager import TemplateManager
from app.engine.formatter import Formatter
from app.ui.template_panel import TemplatePanel
from app.ui.settings_stack import SettingsStack, CATEGORIES

# 导入各设置页面
from app.ui.pages.page_settings import PageSettingsPage
from app.ui.pages.title_settings import TitleSettingsPage
from app.ui.pages.abstract_settings import AbstractSettingsPage
from app.ui.pages.en_abstract_settings import EnAbstractSettingsPage
from app.ui.pages.heading_settings import HeadingSettingsPage
from app.ui.pages.body_settings import BodySettingsPage
from app.ui.pages.caption_settings import CaptionSettingsPage
from app.ui.pages.reference_settings import ReferenceSettingsPage
from app.ui.pages.toc_settings import TOCSettingsPage
from app.ui.pages.table_style_settings import TableStyleSettingsPage
from app.ui.pages.header_footer_settings import HeaderFooterSettingsPage


class FormatWorker(QThread):
    progress = pyqtSignal(str, int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, config: FormatConfig, input_path: str, output_path: str):
        super().__init__()
        self._config = config
        self._input = input_path
        self._output = output_path

    def run(self):
        try:
            formatter = Formatter(self._config)
            result = formatter.format(
                self._input, self._output,
                progress=lambda msg, pct: self.progress.emit(msg, pct))
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('论文排版工具 - HPU Thesis Formatter')
        self.resize(1200, 750)

        self._tm = TemplateManager()
        self._tm.ensure_builtin_exists()

        self._config: FormatConfig | None = None
        self._doc_path: str | None = None
        self._worker: FormatWorker | None = None

        self._setup_menu()
        self._setup_toolbar()
        self._setup_central()
        self._setup_statusbar()

        # 加载默认模板
        self._load_template('HPU_默认')

    # ---- 菜单栏 ----
    def _setup_menu(self):
        mb = self.menuBar()

        file_menu = mb.addMenu('文件(&F)')
        act_open = file_menu.addAction('打开文档...', self._on_open_doc)
        act_open.setShortcut('Ctrl+O')
        file_menu.addSeparator()
        act_exit = file_menu.addAction('退出(&X)', self.close)
        act_exit.setShortcut('Ctrl+Q')

        tmpl_menu = mb.addMenu('模板(&T)')
        act_save = tmpl_menu.addAction('保存当前模板', self._on_save_template)
        act_save.setShortcut('Ctrl+S')
        tmpl_menu.addAction('另存为新模板...', self._on_save_as_template)
        tmpl_menu.addSeparator()
        tmpl_menu.addAction('恢复 HPU 默认', self._on_reset_default)

        help_menu = mb.addMenu('帮助(&H)')
        help_menu.addAction('关于...', self._on_about)

    # ---- 工具栏 ----
    def _setup_toolbar(self):
        tb = QToolBar('主工具栏')
        tb.setIconSize(QSize(24, 24))
        tb.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.addToolBar(tb)

        tb.addAction('📂 打开文档', self._on_open_doc)
        tb.addSeparator()

        tb.addWidget(QLabel('  模板：'))
        self._tmpl_combo = QComboBox()
        self._tmpl_combo.setMinimumWidth(160)
        self._tmpl_combo.currentTextChanged.connect(self._on_template_combo)
        tb.addWidget(self._tmpl_combo)
        tb.addSeparator()

        self._format_btn = QPushButton('▶ 一键排版')
        self._format_btn.setStyleSheet(
            'QPushButton { background-color: #1976D2; color: white; '
            'padding: 6px 20px; font-size: 14px; font-weight: bold; '
            'border-radius: 4px; }'
            'QPushButton:hover { background-color: #1565C0; }'
            'QPushButton:disabled { background-color: #BDBDBD; }'
        )
        self._format_btn.clicked.connect(self._on_format)
        tb.addWidget(self._format_btn)

    # ---- 中央区域 ----
    def _setup_central(self):
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 右侧：设置页面（先创建，模板面板的信号需要引用它）
        self._settings = SettingsStack()
        self._register_pages()
        splitter.addWidget(self._settings)

        # 左侧：模板面板
        self._tmpl_panel = TemplatePanel(self._tm)
        self._tmpl_panel.template_selected.connect(self._on_template_loaded)
        self._tmpl_panel.template_list_changed.connect(self._refresh_combo)
        self._tmpl_panel.save_requested.connect(self._settings.save_all_pages)
        splitter.addWidget(self._tmpl_panel)

        splitter.setSizes([260, 940])
        self.setCentralWidget(splitter)

    def _register_pages(self):
        """注册所有设置页面"""
        pages = [
            ('page', PageSettingsPage()),
            ('title', TitleSettingsPage()),
            ('abstract', AbstractSettingsPage()),
            ('en_abstract', EnAbstractSettingsPage()),
            ('headings', HeadingSettingsPage()),
            ('body', BodySettingsPage()),
            ('captions', CaptionSettingsPage()),
            ('references', ReferenceSettingsPage()),
            ('toc', TOCSettingsPage()),
            ('table_style', TableStyleSettingsPage()),
            ('header_footer', HeaderFooterSettingsPage()),
        ]
        for key, page in pages:
            self._settings.register_page(key, page)

        # 默认显示第一页
        self._settings.show_page('page')

    # ---- 状态栏 ----
    def _setup_statusbar(self):
        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._status_label = QLabel('就绪')
        self._status.addWidget(self._status_label)

    # ---- 模板操作 ----
    def _load_template(self, name: str):
        try:
            self._config = self._tm.load(name)
            self._settings.set_config(self._config)
            # 刷新当前页面
            self._settings._current_key = ''
            self._settings.show_page('page')
            self._status_label.setText(f'已加载模板：{name}')
        except Exception as e:
            QMessageBox.warning(self, '加载失败', str(e))

    def _on_template_loaded(self, config: FormatConfig):
        self._config = config
        self._settings.set_config(config)
        self._settings.show_page('page')
        self._settings._current_key = ''
        self._status_label.setText(f'已加载模板：{config.name}')

    def _on_template_combo(self, name: str):
        if name and self._tm.exists(name):
            self._load_template(name)

    def _refresh_combo(self):
        self._tmpl_combo.blockSignals(True)
        current = self._tmpl_combo.currentText()
        self._tmpl_combo.clear()
        self._tmpl_combo.addItems(self._tm.list_templates())
        if current and self._tm.exists(current):
            self._tmpl_combo.setCurrentText(current)
        elif self._config:
            self._tmpl_combo.setCurrentText(self._config.name)
        self._tmpl_combo.blockSignals(False)

    def _on_save_template(self):
        self._settings.save_all_pages()
        if self._config is None:
            return
        self._tm.save(self._config)
        self._refresh_combo()
        self._status_label.setText(f'模板 "{self._config.name}" 已保存')

    def _on_save_as_template(self):
        from PyQt6.QtWidgets import QInputDialog
        self._settings.save_all_pages()
        if self._config is None:
            return
        name, ok = QInputDialog.getText(self, '另存为', '新模板名称：',
                                        text=self._config.name)
        if not ok or not name.strip():
            return
        name = name.strip()
        old_name = self._config.name
        self._config.name = name
        self._tm.save(self._config)
        # 如果名称变了，删除旧模板
        if old_name != name and self._tm.exists(old_name):
            self._tm.delete(old_name)
        self._refresh_combo()
        self._status_label.setText(f'模板已另存为：{name}')

    def _on_reset_default(self):
        reply = QMessageBox.question(
            self, '确认重置', '确定要恢复为 HPU 默认模板吗？当前修改将丢失。',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        config = self._tm.load_builtin()
        config.name = 'HPU_默认'
        self._tm.save(config)
        self._load_template('HPU_默认')
        self._refresh_combo()

    # ---- 文档操作 ----
    def _on_open_doc(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '打开文档', '',
            'Word 文档 (*.docx);;所有文件 (*)')
        if not path:
            return
        self._doc_path = path
        fname = os.path.basename(path)
        self._status_label.setText(f'已打开文档：{fname}')
        self.setWindowTitle(f'论文排版工具 - {fname}')

    # ---- 一键排版 ----
    def _on_format(self):
        if self._doc_path is None:
            QMessageBox.information(self, '提示', '请先打开一个 .docx 文档。')
            return
        if self._config is None:
            QMessageBox.information(self, '提示', '请先选择一个模板。')
            return
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.information(self, '提示', '排版正在进行中，请稍候...')
            return

        # 弹出保存路径选择
        output_path, _ = QFileDialog.getSaveFileName(
            self, '保存排版结果', self._doc_path,
            'Word 文档 (*.docx);;所有文件 (*)')
        if not output_path:
            return

        # 保存当前设置
        self._settings.save_all_pages()

        self._format_btn.setEnabled(False)
        self._status_label.setText('正在排版...')

        self._worker = FormatWorker(self._config, self._doc_path, output_path)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_format_done)
        self._worker.error.connect(self._on_format_error)
        self._worker.start()

    def _on_progress(self, msg: str, pct: int):
        self._status_label.setText(f'[{pct}%] {msg}')

    def _on_format_done(self, output_path: str):
        self._format_btn.setEnabled(True)
        self._status_label.setText(f'排版完成！已保存至：{output_path}')
        QMessageBox.information(
            self, '排版完成',
            f'文档排版已完成！\n\n保存至：{output_path}\n\n'
            '请在 Word 中打开并更新目录域（右键目录 → 更新域）。')

    def _on_format_error(self, err: str):
        self._format_btn.setEnabled(True)
        self._status_label.setText('排版失败')
        QMessageBox.critical(self, '排版失败', f'排版过程中发生错误：\n\n{err}')

    # ---- 关于 ----
    def _on_about(self):
        QMessageBox.about(
            self, '关于',
            '论文排版工具 v1.0\n\n'
            '基于河南理工大学本科毕业设计（论文）撰写规范\n'
            '支持自定义排版模板，一键应用到 .docx 文档。\n\n'
            '技术栈：Python + PyQt6 + python-docx')

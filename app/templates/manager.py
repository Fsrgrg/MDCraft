"""模板管理器：JSON 读写、验证、导入导出"""

import os
import json
import shutil
from typing import List, Optional
from app.models.config import FormatConfig


class TemplateManager:
    def __init__(self, templates_dir: Optional[str] = None):
        if templates_dir is None:
            appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
            templates_dir = os.path.join(appdata, 'HPUThesisFormatter', 'templates')
        self._dir = templates_dir
        os.makedirs(self._dir, exist_ok=True)

    @property
    def directory(self) -> str:
        return self._dir

    def list_templates(self) -> List[str]:
        """返回所有模板名称（不含扩展名）"""
        names = []
        if not os.path.isdir(self._dir):
            return names
        for f in os.listdir(self._dir):
            if f.endswith('.hputpl'):
                names.append(f[:-7])
        return sorted(names)

    def get_path(self, name: str) -> str:
        return os.path.join(self._dir, f'{name}.hputpl')

    def exists(self, name: str) -> bool:
        return os.path.isfile(self.get_path(name))

    def load(self, name: str) -> FormatConfig:
        """加载模板，若不存在则抛出 FileNotFoundError"""
        path = self.get_path(name)
        with open(path, 'r', encoding='utf-8') as f:
            return FormatConfig.from_json(f.read())

    def save(self, config: FormatConfig) -> str:
        """保存模板，返回文件路径"""
        path = self.get_path(config.name)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(config.to_json())
        return path

    def delete(self, name: str) -> bool:
        """删除模板，返回是否成功"""
        path = self.get_path(name)
        if os.path.isfile(path):
            os.remove(path)
            return True
        return False

    def rename(self, old_name: str, new_name: str) -> bool:
        """重命名模板"""
        old_path = self.get_path(old_name)
        new_path = self.get_path(new_name)
        if not os.path.isfile(old_path):
            return False
        if os.path.isfile(new_path):
            return False
        os.rename(old_path, new_path)
        return True

    def duplicate(self, name: str, new_name: str) -> Optional[str]:
        """复制模板，返回新模板路径"""
        src = self.get_path(name)
        if not os.path.isfile(src):
            return None
        dst = self.get_path(new_name)
        if os.path.isfile(dst):
            return None
        shutil.copy2(src, dst)
        return dst

    def export_to(self, name: str, filepath: str) -> bool:
        """导出模板到指定路径"""
        src = self.get_path(name)
        if not os.path.isfile(src):
            return False
        shutil.copy2(src, filepath)
        return True

    def import_from(self, filepath: str) -> Optional[FormatConfig]:
        """从文件导入模板，验证后保存到模板目录"""
        with open(filepath, 'r', encoding='utf-8') as f:
            config = FormatConfig.from_json(f.read())
        if self.exists(config.name):
            config.name = config.name + '_导入'
        self.save(config)
        return config

    def load_builtin(self) -> FormatConfig:
        """加载内置默认模板（HPU）"""
        import sys
        import pkgutil
        import app.resources

        data = pkgutil.get_data('app.resources', 'hpu_default.json')
        if data is None:
            # Fallback for PyInstaller: look in sys._MEIPASS
            meipass = getattr(sys, '_MEIPASS', '')
            if meipass:
                path = os.path.join(meipass, 'app', 'resources', 'hpu_default.json')
                if os.path.isfile(path):
                    with open(path, 'r', encoding='utf-8') as f:
                        return FormatConfig.from_json(f.read())
            # Fallback: try relative to this file
            res_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'resources')
            path = os.path.join(res_dir, 'hpu_default.json')
            with open(path, 'r', encoding='utf-8') as f:
                return FormatConfig.from_json(f.read())
        return FormatConfig.from_json(data.decode('utf-8'))

    def ensure_builtin_exists(self) -> str:
        """确保内置模板已复制到用户模板目录，返回模板名"""
        name = 'HPU_默认'
        if not self.exists(name):
            config = self.load_builtin()
            config.name = name
            self.save(config)
        return name

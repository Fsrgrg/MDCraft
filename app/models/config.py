"""排版配置数据模型 — 所有参数可序列化为 JSON"""

from dataclasses import dataclass, field, asdict
from typing import Optional
import json


@dataclass
class ElementStyle:
    """单个文档元素的排版样式"""
    cn_font: str = '宋体'
    en_font: str = 'Times New Roman'
    font_size: float = 12.0           # pt
    bold: bool = False
    alignment: str = 'left'           # left | center | right | justify
    line_spacing: float = 1.3         # 倍数 or 固定磅值
    line_spacing_rule: str = 'multiple'  # multiple | exact
    first_line_indent: Optional[float] = None  # cm, None=无缩进
    space_before: float = 0.0         # pt
    space_after: float = 0.0          # pt

    def to_dict(self) -> dict:
        d = asdict(self)
        if self.first_line_indent is None:
            d['first_line_indent'] = None
        return d

    @classmethod
    def from_dict(cls, d: dict) -> 'ElementStyle':
        defaults = {
            'cn_font': '宋体', 'en_font': 'Times New Roman',
            'font_size': 12.0, 'bold': False, 'alignment': 'left',
            'line_spacing': 1.3, 'line_spacing_rule': 'multiple',
            'first_line_indent': None, 'space_before': 0.0, 'space_after': 0.0
        }
        return cls(**{k: d.get(k, defaults[k]) for k in defaults})


@dataclass
class PageConfig:
    """页面设置"""
    page_width: float = 21.0          # cm
    page_height: float = 29.7         # cm
    top_margin: float = 2.54          # cm
    bottom_margin: float = 2.54       # cm
    left_margin: float = 3.18         # cm
    right_margin: float = 3.18        # cm

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> 'PageConfig':
        return cls(
            page_width=d.get('page_width', 21.0),
            page_height=d.get('page_height', 29.7),
            top_margin=d.get('top_margin', 2.54),
            bottom_margin=d.get('bottom_margin', 2.54),
            left_margin=d.get('left_margin', 3.18),
            right_margin=d.get('right_margin', 3.18),
        )


@dataclass
class TableStyleConfig:
    """表格样式"""
    cn_font: str = '宋体'
    en_font: str = 'Times New Roman'
    font_size: float = 10.5           # pt
    alignment: str = 'center'
    line_spacing: float = 1.0
    top_border_width: float = 1.5     # pt (三线表顶线)
    bottom_border_width: float = 1.5  # pt (三线表底线)
    header_bottom_width: float = 0.75 # pt (表头下线)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> 'TableStyleConfig':
        return cls(
            cn_font=d.get('cn_font', '宋体'),
            en_font=d.get('en_font', 'Times New Roman'),
            font_size=d.get('font_size', 10.5),
            alignment=d.get('alignment', 'center'),
            line_spacing=d.get('line_spacing', 1.0),
            top_border_width=d.get('top_border_width', 1.5),
            bottom_border_width=d.get('bottom_border_width', 1.5),
            header_bottom_width=d.get('header_bottom_width', 0.75),
        )


@dataclass
class HeaderFooterConfig:
    """页眉页脚设置"""
    odd_header_source: str = 'chapter'  # 'chapter' | 'fixed'
    odd_header_fixed_text: str = ''
    even_header_text: str = '河南理工大学本科毕业设计（论文）'
    cn_font: str = '宋体'
    en_font: str = 'Times New Roman'
    font_size: float = 9.0             # pt
    page_num_format: str = 'decimal'    # 'decimal' | 'upperRoman' | 'lowerRoman'
    header_line: bool = True            # 页眉下方横线

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> 'HeaderFooterConfig':
        return cls(
            odd_header_source=d.get('odd_header_source', 'chapter'),
            odd_header_fixed_text=d.get('odd_header_fixed_text', ''),
            even_header_text=d.get('even_header_text', '河南理工大学本科毕业设计（论文）'),
            cn_font=d.get('cn_font', '宋体'),
            en_font=d.get('en_font', 'Times New Roman'),
            font_size=d.get('font_size', 9.0),
            page_num_format=d.get('page_num_format', 'decimal'),
            header_line=d.get('header_line', True),
        )


@dataclass
class TOCConfig:
    """目录设置"""
    auto_insert: bool = True           # 自动插入目录
    heading_levels: str = '1-3'        # 包含的标题层级
    toc_heading_style: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=15.0, alignment='center', line_spacing=2.0))
    toc_item_style: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='left', line_spacing=1.3))

    def to_dict(self) -> dict:
        return {
            'auto_insert': self.auto_insert,
            'heading_levels': self.heading_levels,
            'toc_heading_style': self.toc_heading_style.to_dict(),
            'toc_item_style': self.toc_item_style.to_dict(),
        }

    @classmethod
    def from_dict(cls, d: dict) -> 'TOCConfig':
        return cls(
            auto_insert=d.get('auto_insert', True),
            heading_levels=d.get('heading_levels', '1-3'),
            toc_heading_style=ElementStyle.from_dict(d.get('toc_heading_style', {})),
            toc_item_style=ElementStyle.from_dict(d.get('toc_item_style', {})),
        )


@dataclass
class FormatConfig:
    """完整的排版配置"""
    name: str = '未命名模板'
    description: str = ''

    page: PageConfig = field(default_factory=PageConfig)

    # 题目页
    title_uni: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=36.0, alignment='center', line_spacing=1.5))
    title_doctype: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=22.0, alignment='center', line_spacing=1.5))
    title_paper: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=18.0, alignment='center', line_spacing=1.5))
    title_date: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=15.0, alignment='center', line_spacing=1.5))
    title_table: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=15.0, alignment='center', line_spacing=1.5))

    # 中文摘要
    abstract_label: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=15.0, alignment='center', line_spacing=2.0))
    abstract_body: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='justify', line_spacing=1.3,
        first_line_indent=0.74))
    keywords_label: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, bold=True, alignment='justify',
        line_spacing=1.3, first_line_indent=0.74))
    keywords_body: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='justify', line_spacing=1.3,
        first_line_indent=0.74))

    # 英文摘要
    en_abstract_label: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=15.0, bold=True, alignment='center', line_spacing=2.0))
    en_abstract_body: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='justify', line_spacing=1.3,
        first_line_indent=0.74))
    en_keywords_label: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, bold=True, alignment='justify',
        line_spacing=1.3, first_line_indent=0.74))
    en_keywords_body: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='justify', line_spacing=1.3,
        first_line_indent=0.74))

    # 正文标题
    heading_1: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=15.0, alignment='center', line_spacing=2.0))
    heading_2: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=14.0, bold=True, alignment='left', line_spacing=1.5))
    heading_3: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='left', line_spacing=1.3,
        first_line_indent=0.74))
    heading_4: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='left', line_spacing=1.3,
        first_line_indent=0.74))

    # 正文
    body: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='justify', line_spacing=1.3,
        first_line_indent=0.74))

    # 图表题注
    table_caption: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=10.5, alignment='center', line_spacing=1.0))
    figure_caption: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=10.5, alignment='center', line_spacing=1.0))

    # 参考文献
    ref_heading: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=15.0, alignment='center', line_spacing=2.0))
    ref_item: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=10.5, alignment='justify', line_spacing=1.0))

    # 附录
    appendix_heading: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='黑体', font_size=15.0, alignment='center', line_spacing=2.0))
    appendix_sub: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='left', line_spacing=1.3,
        first_line_indent=0.74))

    # 源代码
    source_code: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=9.0, alignment='left', line_spacing=1.0))

    # 文末注释
    end_note: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=10.5, alignment='left', line_spacing=1.0))

    # 公式
    formula: ElementStyle = field(default_factory=lambda: ElementStyle(
        cn_font='宋体', font_size=12.0, alignment='center', line_spacing=1.3))

    # 表格样式
    table_style: TableStyleConfig = field(default_factory=TableStyleConfig)

    # 页眉页脚
    header_footer: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)

    # 目录
    toc: TOCConfig = field(default_factory=TOCConfig)

    # ---- 序列化 ----

    _style_fields = [
        'title_uni', 'title_doctype', 'title_paper', 'title_date', 'title_table',
        'abstract_label', 'abstract_body', 'keywords_label', 'keywords_body',
        'en_abstract_label', 'en_abstract_body', 'en_keywords_label', 'en_keywords_body',
        'heading_1', 'heading_2', 'heading_3', 'heading_4',
        'body', 'table_caption', 'figure_caption',
        'ref_heading', 'ref_item',
        'appendix_heading', 'appendix_sub',
        'source_code', 'end_note', 'formula',
    ]

    def to_dict(self) -> dict:
        d = {
            'name': self.name,
            'description': self.description,
            'page': self.page.to_dict(),
        }
        for f in self._style_fields:
            d[f] = getattr(self, f).to_dict()
        d['table_style'] = self.table_style.to_dict()
        d['header_footer'] = self.header_footer.to_dict()
        d['toc'] = self.toc.to_dict()
        return d

    @classmethod
    def from_dict(cls, d: dict) -> 'FormatConfig':
        config = cls(
            name=d.get('name', '未命名模板'),
            description=d.get('description', ''),
            page=PageConfig.from_dict(d.get('page', {})),
        )
        for f in cls._style_fields:
            if f in d:
                setattr(config, f, ElementStyle.from_dict(d[f]))
        config.table_style = TableStyleConfig.from_dict(d.get('table_style', {}))
        config.header_footer = HeaderFooterConfig.from_dict(d.get('header_footer', {}))
        config.toc = TOCConfig.from_dict(d.get('toc', {}))
        return config

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    @classmethod
    def from_json(cls, s: str) -> 'FormatConfig':
        return cls.from_dict(json.loads(s))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""党政机关公文 DOCX 生成器（GB/T 9704-2012）

三种用法：
  1. 配置文件：python3 generate_gongwen.py --config content.json
  2. 纯文本正文：python3 generate_gongwen.py --title "..." --body-file body.txt \
         --recipient "..." --signature "..." --date "..." --output out.docx
  3. 作为模块调用：from generate_gongwen import build_document

依赖：pip install python-docx
"""

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from lxml import etree

# ── 排版常量（依据 GB/T 9704-2012） ──────────────────────

FONT_TITLE = '方正小标宋简体'      # 2号小标宋体
FONT_BODY = '仿宋_GB2312'          # 3号仿宋体
FONT_KAI = '楷体_GB2312'           # 第二层楷体
FONT_HEI = '黑体'                  # 第一层黑体
FONT_PAGE = '宋体'                 # 4号宋体

SIZE_TITLE = 22                    # 二号
SIZE_BODY = 16                     # 三号
SIZE_PAGE = 14                     # 四号

LINE_TITLE = 32                    # 标题行距（磅，固定值）
LINE_BODY = 28                     # 正文行距（磅，固定值）

# ── 字体回退链 ───────────────────────────────────────────
# 国标字体名在不同系统上实际名称不同：Windows 常见「仿宋_GB2312」，
# macOS 通常只有「方正仿宋_GBK」「华文仿宋」，Linux 多用「FangSong」。
# 若写死国标名而本机没有，Word 会静默回退到默认字体，导致格式走样。
# 这里按「国标名 → 等效字体 → 通用兜底」逐项回退，取本机第一个可用的。
FONT_CHAINS = {
    'title': ['方正小标宋简体', '小标宋体', '方正小标宋_GBK', '华文中宋',
              'STZhongsong', '宋体-简', '宋体'],
    'body':  ['仿宋_GB2312', '方正仿宋_GBK', '方正仿宋', '华文仿宋',
              '仿宋', 'FangSong', '宋体'],
    'kai':   ['楷体_GB2312', '楷体-简', '华文楷体', '楷体', 'Kai',
              'Kaiti SC', '宋体'],
    'hei':   ['黑体', '黑体-简', '华文黑体', 'Hei', 'Heiti SC', 'SimHei'],
    'page':  ['宋体', '宋体-简', '华文宋体', 'SimSun', 'Songti SC'],
}

# ── 段落角色 ─────────────────────────────────────────────
#
# 角色由调用方（AI）显式指定，脚本只负责把角色渲染成对应版式，
# 不根据编号猜语义。因为「1.」既可能是三级标题，也可能只是一条
# 编号的工作事项——把整段叙述误加粗，比漏加粗难看得多。
#
# 推荐在 content.body 中显式写 {"type": "<role>", "text": "..."}。

ROLE_STYLES = {
    'heading1': ('hei', False),        # 一、      黑体
    'heading2': ('kai', False),        # （一）    楷体
    'heading3': ('body', True),        # 1.        仿宋加粗
    'heading4': ('body', True),        # （1）     仿宋加粗
    'heading5': ('body', True),        # ①         仿宋加粗
    'body': ('body', False),           # 普通正文
    'numbered_body': ('body', False),  # 编号正文：保留编号，按正文排，不加粗
}

# 旧字段名兼容
ROLE_ALIASES = {
    'level1': 'heading1', 'level2': 'heading2', 'level3': 'heading3',
    'level4': 'heading4', 'level5': 'heading5',
    'text': 'body', 'para': 'body', 'plain': 'body',
}

# 兜底自动识别：仅在元素是字符串（未显式指定角色）时使用。
# 原则：宁可漏判为正文，也不要把整段叙述误加粗。
#   一、/（一）几乎不会是正文罗列 → 仍按标题
#   1./（1）/① 常用于罗列工作事项 → 兜底一律按编号正文（不加粗）
LEVEL_PATTERNS = [
    (re.compile(r'^[一二三四五六七八九十]+、'), 'heading1'),
    (re.compile(r'^（[一二三四五六七八九十]+）'), 'heading2'),
    (re.compile(r'^\d+\.'), 'numbered_body'),
    (re.compile(r'^（\d+）'), 'numbered_body'),
    (re.compile(r'^[①-⑳]'), 'numbered_body'),
]


# ── 字体可用性检测与回退 ─────────────────────────────────

_font_cache = None


def available_fonts():
    """列出本机可用字体名集合。检测失败时返回空集合（此时不回退，保留国标名）。"""
    global _font_cache
    if _font_cache is not None:
        return _font_cache
    import subprocess
    names = set()
    try:
        if sys.platform == 'darwin':
            out = subprocess.run(['system_profiler', 'SPFontsDataType'],
                                 capture_output=True, text=True, timeout=90).stdout
            names = set(m.strip() for m in re.findall(r'^\s*Family:\s*(.+)$', out, re.M))
        elif sys.platform.startswith('linux'):
            out = subprocess.run(['fc-list', ':', 'family'],
                                 capture_output=True, text=True, timeout=30).stdout
            for line in out.splitlines():
                names.update(f.strip() for f in line.split(','))
        elif sys.platform == 'win32':
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
            i = 0
            while True:
                try:
                    name = winreg.EnumValue(key, i)[0]
                    names.add(re.sub(r'\s*\((?:TrueType|OpenType|All res)\)\s*$',
                                     '', name))
                    i += 1
                except OSError:
                    break
    except Exception:
        names = set()
    _font_cache = names
    return names


def resolve_fonts():
    """返回 {role: 本机实际可用的字体名}。"""
    avail = available_fonts()
    resolved = {}
    for role, chain in FONT_CHAINS.items():
        resolved[role] = next((f for f in chain if f in avail), chain[0])
    return resolved


def apply_font_fallback():
    """把模块级字体常量改写为回退链中本机第一个可用的字体，并重建层级规则。

    返回 {role: (国标名, 实际使用名)}，便于向用户报告替换情况。
    """
    global FONT_TITLE, FONT_BODY, FONT_KAI, FONT_HEI, FONT_PAGE

    resolved = resolve_fonts()
    before = {'title': FONT_TITLE, 'body': FONT_BODY, 'kai': FONT_KAI,
              'hei': FONT_HEI, 'page': FONT_PAGE}
    FONT_TITLE = resolved['title']
    FONT_BODY = resolved['body']
    FONT_KAI = resolved['kai']
    FONT_HEI = resolved['hei']
    FONT_PAGE = resolved['page']
    return {r: (before[r], resolved[r]) for r in resolved}


def resolve_role(item):
    """把 body 元素解析为 (文本, 角色)。

    item 可为：
      {"type": "heading3"|"numbered_body"|"body"|"blank"|..., "text": "..."}
      或纯字符串（走兜底自动识别）
    返回 (text, role)；role 为 'blank' 时调用方应插入空行。
    """
    if isinstance(item, str):
        return item, detect_level(item) or 'body'

    text = item.get('text', '')
    role = item.get('type') or 'body'
    role = ROLE_ALIASES.get(role, role)
    if role == 'blank':
        return text, 'blank'
    if role not in ROLE_STYLES:
        # 未知角色：不猜，按最安全的正文处理
        role = 'body'
    return text, role


# ── 基础排版函数 ─────────────────────────────────────────

def set_font(run, font_name, font_size, bold=False):
    """设置字体。中文必须设置 w:eastAsia，否则 Word 回退默认字体。"""
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def set_line_spacing(paragraph, pt_value):
    """固定值行距。lineRule=exact。"""
    pPr = paragraph._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(int(pt_value * 20)))   # 磅 -> twip
    spacing.set(qn('w:lineRule'), 'exact')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')


def _set_doc_grid(section, line_pitch_twips=560):
    """声明行网格（GB/T 9704-2012 5.2.3：一般每面 22 行）。

    只锁行距网格，不锁字符网格（type 用 lines 而非 linesAndChars）——
    锁字符会强制每行字数，导致中英混排被拉伸变形。
    linePitch 与正文固定行距取同一值，两者一致时不会产生冲突。
    """
    sectPr = section._sectPr
    docGrid = sectPr.find(qn('w:docGrid'))
    if docGrid is None:
        docGrid = OxmlElement('w:docGrid')
        sectPr.append(docGrid)
    docGrid.set(qn('w:type'), 'lines')
    docGrid.set(qn('w:linePitch'), str(line_pitch_twips))


def set_first_line_indent(paragraph, chars=2):
    """首行缩进 chars 个字符。

    python-docx 的 Normal 样式自带 firstLineChars=200，直接追加会叠加导致缩进翻倍，
    必须先清除已有缩进属性，再同时写入字符值与绝对值（Word 以字符值优先）。
    """
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    for attr in ('w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars',
                 'w:hanging', 'w:hangingChars', 'w:right', 'w:rightChars'):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]
    ind.set(qn('w:firstLineChars'), str(chars * 100))
    ind.set(qn('w:firstLine'), str(chars * 320))


def set_right_indent_chars(paragraph, chars=1):
    """右缩进 chars 个字符（用于单页码居右空一字）。"""
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:rightChars'), str(chars * 100))
    ind.set(qn('w:right'), str(chars * 320))


def set_left_indent_chars(paragraph, chars=1):
    """左缩进 chars 个字符（用于双页码居左空一字）。"""
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:leftChars'), str(chars * 100))
    ind.set(qn('w:left'), str(chars * 320))


def add_paragraph(doc, text, font_name=None, font_size=SIZE_BODY,
                  bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  indent=True, line_spacing=LINE_BODY):
    # font_name 默认 None：不能在签名里写 FONT_BODY，默认参数在定义时就绑定，
    # 字体回退改写 FONT_BODY 后不会生效。
    para = doc.add_paragraph()
    para.alignment = align
    set_line_spacing(para, line_spacing)
    if indent:
        set_first_line_indent(para, 2)
    run = para.add_run(text)
    set_font(run, font_name or FONT_BODY, font_size, bold)
    return para


def add_blank(doc, line_spacing=LINE_BODY):
    para = doc.add_paragraph()
    set_line_spacing(para, line_spacing)
    return para


# VML（老式 w:pict 图片）命名空间，python-docx 的 nsmap 里没有，需手写
VML_NS = 'urn:schemas-microsoft-com:vml'


def _rebuild_images(src_doc, dst_doc, new_p):
    """重建被复制段落里图片的关系 ID。

    deepcopy 过来的 r:embed 仍指向源文档的 image part，在新文档里是悬空引用，
    必须把图片重新加进目标文档并换成新的 rId，否则图片显示为空白甚至报损坏。
    """
    def rebuild(rid, setter):
        if not rid or rid not in src_doc.part.related_parts:
            return
        img_part = src_doc.part.related_parts[rid]
        try:
            import io
            # get_or_add_image 接受字节流，返回 (rId, image)
            new_rid, _img = dst_doc.part.get_or_add_image(io.BytesIO(img_part.blob))
            setter(new_rid)
        except Exception as exc:
            print(f'警告：版头中的图片未能复制（{exc}），该图可能显示为空白')

    # Drawing（Office 2007+）与 VML（w:pict，老文档）两种图片都要处理
    for blip in new_p.iter(qn('a:blip')):
        rid = blip.get(qn('r:embed'))
        rebuild(rid, lambda v, b=blip: b.set(qn('r:embed'), v))
    # 注意：python-docx 的 nsmap 没有 VML 前缀，qn('v:...') 会 KeyError，须写完整 URI
    for imagedata in new_p.iter(f'{{{VML_NS}}}imagedata'):
        rid = imagedata.get(qn('r:id'))
        rebuild(rid, lambda v, e=imagedata: e.set(qn('r:id'), v))


def _theme_font_map(src_doc):
    """解析源文档 theme 部件的 major/minor 字体名。

    返回 {'majorHAnsi': 'Calibri', 'minorEastAsia': '宋体', ...}。
    文档默认字体常写成主题引用（如 w:eastAsiaTheme="minorEastAsia"），
    不解析出真实字体名就无法在新文档里还原。
    """
    out = {}
    root = None
    try:
        for part in src_doc.part.package.iter_parts():
            if 'theme' in str(part.partname):
                root = etree.fromstring(part.blob)
                break
    except Exception:
        return out
    if root is None:
        return out
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for scheme, prefix in (('majorFont', 'major'), ('minorFont', 'minor')):
        node = root.find('{{{a}}}themeElements/{{{a}}}fontScheme/{{{a}}}{s}'
                         .format(a=A, s=scheme))
        if node is None:
            continue
        for tag, suffix in (('latin', 'HAnsi'), ('ea', 'EastAsia'), ('cs', 'Bidi')):
            el = node.find('{{{a}}}{t}'.format(a=A, t=tag))
            if el is not None:
                out[prefix + suffix] = el.get('typeface') or ''
    return out


def _resolve_default_fonts(src_doc, fonts_el):
    """把文档默认字体（w:rFonts）里的主题引用解析成真实字体名。

    返回 {'ascii': ..., 'hAnsi': ..., 'eastAsia': ..., 'cs': ...}，
    只含解析成功的键。
    """
    attrs = {}
    for w, key in (('w:ascii', 'ascii'), ('w:hAnsi', 'hAnsi'),
                   ('w:eastAsia', 'eastAsia'), ('w:cs', 'cs')):
        v = fonts_el.get(qn(w))
        if v:
            attrs[key] = v
    theme = None
    for w, key in (('w:asciiTheme', 'ascii'), ('w:hAnsiTheme', 'hAnsi'),
                   ('w:eastAsiaTheme', 'eastAsia'), ('w:cstheme', 'cs')):
        v = fonts_el.get(qn(w))
        if not v or key in attrs:
            continue
        if theme is None:
            theme = _theme_font_map(src_doc)
        # v 形如 minorEastAsia / majorHAnsi；取不到时按默认档位兜底
        name = theme.get(v) or theme.get(
            'minor' + ('EastAsia' if key == 'eastAsia' else 'HAnsi'), '')
        if name:
            attrs[key] = name
    return attrs


def sync_doc_defaults(src_doc, dst_doc):
    """把源文档的文档级默认字体（docDefaults）同步到新文档。

    源文档里没有显式设字体的 run（红头常常如此——靠样式或文档默认显示），
    在源文档中由 docDefaults 兜底。新建文档的 docDefaults 是 python-docx
    模板的主题引用，与源文档不同，这些 run 复制过来就会换字体。

    正文、页码、表格都由本脚本显式设字体，不受此项影响。
    """
    def _rpr_defaults(d):
        styles = d.styles.element
        dd = styles.find(qn('w:docDefaults'))
        if dd is None:
            return None
        rpd = dd.find(qn('w:rPrDefault'))
        return rpd.find(qn('w:rPr')) if rpd is not None else None

    src_rpr = _rpr_defaults(src_doc)
    dst_rpr = _rpr_defaults(dst_doc)
    if src_rpr is None or dst_rpr is None:
        return None
    src_fonts = src_rpr.find(qn('w:rFonts'))
    if src_fonts is None:
        return None

    resolved = _resolve_default_fonts(src_doc, src_fonts)
    if not resolved:
        return None

    dst_fonts = dst_rpr.find(qn('w:rFonts'))
    if dst_fonts is None:
        dst_fonts = OxmlElement('w:rFonts')
        dst_rpr.insert(0, dst_fonts)
    # 具体字体名与主题引用不能并存，先清主题引用
    for a in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        dst_fonts.attrib.pop(qn(a), None)
    for key, val in resolved.items():
        dst_fonts.set(qn('w:' + key), val)

    # 语言标记影响中文的断行与字体选择，一并同步
    src_lang = src_rpr.find(qn('w:lang'))
    if src_lang is not None:
        dst_lang = dst_rpr.find(qn('w:lang'))
        if dst_lang is None:
            dst_lang = OxmlElement('w:lang')
            dst_rpr.append(dst_lang)
        for a in ('w:eastAsia', 'w:val', 'w:bidi'):
            v = src_lang.get(qn(a))
            if v:
                dst_lang.set(qn(a), v)
    return resolved


def _letterhead_fonts(src_doc, indices):
    """列出版头段落实际用到的字体名（含未显式设置时的文档默认字体）。"""
    fonts = []
    seen = set()

    def _add(name):
        if name and name not in seen:
            seen.add(name)
            fonts.append(name)

    elements = [p._element for p in src_doc.paragraphs]
    for i in indices:
        if i >= len(elements):
            continue
        for run in elements[i].iter(qn('w:r')):
            rPr = run.find(qn('w:rPr'))
            if rPr is None:
                continue
            rf = rPr.find(qn('w:rFonts'))
            if rf is None:
                continue
            for a in ('w:eastAsia', 'w:ascii', 'w:hAnsi'):
                _add(rf.get(qn(a)))
    # 未显式设字体的 run 会用文档默认字体，也算版头实际字体
    src_rpr = src_doc.styles.element.find(qn('w:docDefaults'))
    if src_rpr is not None:
        rpd = src_rpr.find(qn('w:rPrDefault'))
        if rpd is not None:
            rpr = rpd.find(qn('w:rPr'))
            if rpr is not None:
                rf = rpr.find(qn('w:rFonts'))
                if rf is not None:
                    for name in _resolve_default_fonts(src_doc, rf).values():
                        _add(name)
    return fonts


def copy_letterhead(doc, spec):
    """把源文档的版头（红头 / 发文字号 / 分隔线）原样复制到新文档开头。

    整段 deepcopy w:p，字体、颜色、字号、对齐、边框一概照搬。
    红头是各单位的门面，本技能**不重做红头**——重排等于替用户改门面。
    """
    import copy as copy_mod

    src_path = spec.get('source')
    indices = spec.get('indices') or []
    if not src_path or not indices:
        return 0
    if not os.path.exists(src_path):
        print(f'警告：找不到源文档 {src_path}，版头未保留（正文照常输出）')
        return 0

    src = Document(src_path)

    # 先同步文档默认字体：版头里没显式设字体的 run 靠它兜底，
    # 不复制过来就会在新文档里换成别的字体（表现为"红头变样"）。
    defaults = sync_doc_defaults(src, doc)
    if defaults:
        print(f'✓ 已沿用源文档默认字体：'
              + '，'.join(f'{k}={v}' for k, v in defaults.items()))

    elements = [p._element for p in src.paragraphs]
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    copied = 0
    for i in indices:
        if i >= len(elements):
            continue
        new_p = copy_mod.deepcopy(elements[i])
        _rebuild_images(src, doc, new_p)
        # sectPr 必须是 body 的最后一个子元素，不能直接 append 到末尾
        if sectPr is not None:
            sectPr.addprevious(new_p)
        else:
            body.append(new_p)
        copied += 1

    if copied:
        _report_letterhead_fonts(src, indices)
    return copied


def _report_letterhead_fonts(src_doc, indices):
    """提示版头用到的字体在本机是否齐全。

    版头按用户要求原样保留、不做字体回退，所以本机缺字体时由 Word 自行替换，
    可能和正文（走回退链）不是同一张字体。只提示、不擅自改。
    """
    names = _letterhead_fonts(src_doc, indices)
    if not names:
        return
    avail = available_fonts()
    if not avail:
        return                      # 探测失败就不猜，避免误报
    missing = [n for n in names if n not in avail]
    if missing:
        print('↷ 版头字体本机未安装（已原样保留字体名，由 Word 替换显示）：'
              + '、'.join(missing))
        print('    如需与正文统一，可在源文件中把红头改为本机已有的字体，'
              '或安装对应字体')


def _row_width(row, srow):
    """行的逻辑列数：有合并信息时按 gridSpan 累加，否则按元素个数。"""
    if not srow:
        return len(row)
    return sum(max(int((s or {}).get('h', 1) or 1), 1) for s in srow)


def _apply_merges(table, spans, rows):
    """还原合并单元格。

    直接写 w:gridSpan / w:vMerge，不用 python-docx 的 cell.merge()：后者会把
    被合并格的内容追加进来（这里被合并格是留空占位，追加后会留下多余空段落），
    且合并后列索引随即错位。

    每行内部**从右向左**处理：gridSpan 要删掉后续 tc，右侧先删才不会让
    左侧的索引失效。纵向合并只加属性、不动 tc，顺序无所谓。
    """
    for ri, srow in enumerate(spans or []):
        if ri >= len(rows):
            break
        items, c = [], 0
        for span in srow or []:
            s = span or {}
            colspan = max(int(s.get('h', 1) or 1), 1)
            items.append((c, colspan, s.get('v')))
            c += colspan
        for c, colspan, vmerge in reversed(items):
            try:
                cell = table.cell(ri, c)
            except (IndexError, ValueError):
                continue
            tcPr = cell._tc.get_or_add_tcPr()
            # gridSpan / vMerge 必须排在 vAlign 之前（CT_TcPr 的 schema 顺序），
            # 而 vAlign 是填内容时先写进去的，所以这里用 insert 而非 append
            va = tcPr.find(qn('w:vAlign'))
            pos = list(tcPr).index(va) if va is not None else len(tcPr)
            if colspan > 1:
                gs = OxmlElement('w:gridSpan')
                gs.set(qn('w:val'), str(colspan))
                tcPr.insert(pos, gs)
                pos += 1
                # 删掉本行被吞掉的后续 tc，否则行内列数会超出表格列数
                tr = cell._tc.getparent()
                idx = list(tr).index(cell._tc)
                for _ in range(colspan - 1):
                    if idx + 1 < len(tr) and tr[idx + 1].tag == qn('w:tc'):
                        del tr[idx + 1]
            if vmerge in ('restart', 'continue'):
                vm = OxmlElement('w:vMerge')
                vm.set(qn('w:val'), vmerge)
                tcPr.insert(pos, vm)


def add_table(doc, rows, header=True, font_name=None,
              size_header=12, size_body=10.5, align='center',
              spans=None, ncol=None):
    """插入附表。

    国标未对附表字号作强制规定，通行做法是小四（12磅）表头 + 五号（10.5磅）
    数据，字体随正文用仿宋，整体居中并加边框。数据与表头均居中对齐。

    rows: 二维列表，第一行在 header=True 时视为表头。
    spans: 与 rows 同形的合并信息，由 inspect_docx.py 从 gridSpan/vMerge 解析；
           被合并覆盖的位置在 rows 中为 ''，由 _apply_merges 还原成真正的合并格。
    ncol: 表格逻辑列数（有横向合并时大于 len(rows[i])）。
    """
    font_name = font_name or FONT_BODY
    if not rows:
        return None
    if not ncol:
        ncol = max((_row_width(rows[i], spans[i] if spans and i < len(spans) else None)
                    for i in range(len(rows))), default=0)
    if ncol <= 0:
        return None
    table = doc.add_table(rows=0, cols=ncol)
    try:
        table.style = 'Table Grid'      # 全线边框
    except KeyError:
        pass
    table.alignment = (WD_TABLE_ALIGNMENT.CENTER if align == 'center'
                       else WD_TABLE_ALIGNMENT.LEFT)
    table.autofit = True

    align_map = {'left': WD_ALIGN_PARAGRAPH.LEFT,
                 'center': WD_ALIGN_PARAGRAPH.CENTER,
                 'right': WD_ALIGN_PARAGRAPH.RIGHT}
    para_align = align_map.get(align, WD_ALIGN_PARAGRAPH.CENTER)

    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        srow = spans[ri] if spans and ri < len(spans) else None
        is_header = header and ri == 0
        size = size_header if is_header else size_body
        ci = 0                      # 逻辑列游标，按 gridSpan 推进
        for j, text in enumerate(row):
            if ci >= ncol:
                break
            span = (srow[j] if srow and j < len(srow) else None) or {}
            colspan = max(int(span.get('h', 1) or 1), 1)
            cell = cells[ci]
            # 新建单元格自带一个空段落，直接取用即可；
            # clear_content() 会把这个段落也删掉，导致 paragraphs[0] 越界。
            para = cell.paragraphs[0]
            para.alignment = para_align
            set_line_spacing(para, 16 if is_header else 14)
            run = para.add_run(str(text or ''))
            set_font(run, font_name, size, bold=is_header)
            # 垂直居中
            tcPr = cell._tc.get_or_add_tcPr()
            va = OxmlElement('w:vAlign')
            va.set(qn('w:val'), 'center')
            tcPr.append(va)
            ci += colspan

    if spans:
        _apply_merges(table, spans, rows)
    return table


def detect_level(text):
    """兜底：按编号前缀猜角色，返回角色字符串。无匹配返回 None。

    只在元素是字符串（未显式指定 type）时调用。猜不出语义，
    偏向安全侧：阿拉伯数字编号一律按编号正文，不加粗。
    """
    stripped = text.strip()
    for pattern, role in LEVEL_PATTERNS:
        if pattern.match(stripped):
            return role
    return None


# ── 页码 ─────────────────────────────────────────────────

# 页码样式。默认 center：多数单位内行文、呈报稿按居中排，也便于双面打印装订。
#   center    居中（默认）——所有页页码居中，不区分奇偶页
#   odd_even  奇偶页不同——单页居右空一字、双页居左空一字（GB/T 9704-2012 7.5）
#   none      不显示页码
PAGE_NUMBER_MODES = ('center', 'odd_even', 'none')
PAGE_NUMBER_DEFAULT = 'center'

# 页码别名，容错常见写法
PAGE_NUMBER_ALIASES = {
    'center': 'center', 'centre': 'center', 'middle': 'center', '居中': 'center',
    'odd_even': 'odd_even', 'odd-even': 'odd_even', 'oddeven': 'odd_even',
    'parity': 'odd_even', '奇偶': 'odd_even', '奇偶页': 'odd_even',
    'none': 'none', 'off': 'none', 'no': 'none', 'hide': 'none',
    '不显示': 'none', '无': 'none',
}


def resolve_page_number_mode(value):
    """把配置值归一到 PAGE_NUMBER_MODES 之一；无法识别时返回 None。"""
    if value is None:
        return PAGE_NUMBER_DEFAULT
    key = str(value).strip().lower()
    key = PAGE_NUMBER_ALIASES.get(key, key)
    return key if key in PAGE_NUMBER_MODES else None


def _page_field_runs(paragraph):
    """在段落中追加 '— PAGE —' 域代码。"""
    run1 = paragraph.add_run('— ')
    set_font(run1, FONT_PAGE, SIZE_PAGE)

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    run_field = paragraph.add_run()
    run_field._element.append(fld_begin)

    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    run_instr = paragraph.add_run()
    set_font(run_instr, FONT_PAGE, SIZE_PAGE)
    run_instr._element.append(instr)

    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run_end = paragraph.add_run()
    run_end._element.append(fld_end)

    run2 = paragraph.add_run(' —')
    set_font(run2, FONT_PAGE, SIZE_PAGE)


def _make_footer_paragraph(footer, align, indent_fn):
    para = footer.paragraphs[0]
    para.alignment = align
    set_line_spacing(para, SIZE_PAGE)
    if indent_fn:
        indent_fn(para)
    _page_field_runs(para)
    return para


def add_page_number(section, mode=PAGE_NUMBER_DEFAULT):
    """添加页码。

    mode='center'   所有页页码居中（默认）
    mode='odd_even' 单页码居右空一字、双页码居左空一字（GB/T 9704-2012 7.5）
    mode='none'     不添加页码
    """
    if mode == 'none':
        return

    if mode == 'center':
        # 只挂 default 页脚并居中；不开启奇偶页不同，避免 Word 去找空的 even 页脚
        footer = section.footer
        footer.is_linked_to_previous = False
        _make_footer_paragraph(footer, WD_ALIGN_PARAGRAPH.CENTER, None)
        section.footer_distance = Mm(4.58)
        return

    # odd_even
    # 1) 开启奇偶页不同
    settings = section.part.document.settings.element
    if settings.find(qn('w:evenAndOddHeaders')) is None:
        flag = OxmlElement('w:evenAndOddHeaders')
        # CT_Settings 子元素有 schema 顺序：evenAndOddHeaders 位于 rsids 之前。
        # 直接 insert(0) 插到最前面 Word 多半容忍，但严格解析器会判不合规。
        rsids = settings.find(qn('w:rsids'))
        if rsids is not None:
            rsids.addprevious(flag)
        else:
            settings.append(flag)

    # 2) 奇数页页脚（default）：居右空一字
    footer_odd = section.footer
    footer_odd.is_linked_to_previous = False
    _make_footer_paragraph(footer_odd, WD_ALIGN_PARAGRAPH.RIGHT, set_right_indent_chars)

    # 3) 偶数页页脚（even）：居左空一字。取不到时降级，仅保留奇数页页码
    footer_even = _get_even_footer(section)
    if footer_even is not None:
        footer_even.is_linked_to_previous = False
        _make_footer_paragraph(footer_even, WD_ALIGN_PARAGRAPH.LEFT, set_left_indent_chars)

    # 4) 页脚位置：4.58mm（叠加 4 号字半高后，一字线距版心下边缘约 7mm）
    section.footer_distance = Mm(4.58)


def _get_even_footer(section):
    """获取偶数页页脚。python-docx 未公开 even footer，需内部构造；失败则返回 None 降级。"""
    try:
        from docx.enum.section import WD_HEADER_FOOTER
        from docx.section import _Footer
        return _Footer(section._sectPr, section.part, WD_HEADER_FOOTER.EVEN_PAGE)
    except Exception:
        return None


# ── 主体构建 ─────────────────────────────────────────────

def build_document(content, output_path):
    """根据 content 字典生成公文 DOCX。

    content 支持字段：
      title       标题（必填）
      recipient   主送机关，如「省人民政府：」（可选）
      body        正文列表，元素可为字符串（兜底识别）或
                  {"type": "heading1".."heading5"|"body"|"numbered_body"
                          |"blank"|"table", "text": "..."}
                  {"type": "table", "index": 0} 为附表占位
      tables      附表列表 [{"rows": [...], "header": true,
                             "spans": [...], "ncol": 3}]
      attachments 附件名称列表（可选）
      signature   发文机关署名（可选）
      date        成文日期（可选）
      page_number 页码样式：center（无源文件时默认，居中）/
                  odd_even（奇偶页不同）/ none（不显示）
                  有源文件时由 inspect_docx.py 按源文档推断
      output      输出路径（可选，优先级低于 output_path 参数）
    """
    title = (content.get('title') or '').strip()
    if not title:
        raise ValueError('缺少 title：公文标题为必填项，'
                         '请传 --title 或在 content 中设置 title')
    doc = Document()

    # 字体回退：把国标字体名换成本机实际可用的等效字体
    font_report = apply_font_fallback()

    # 页面设置（5.2.1）
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(37)
    section.bottom_margin = Mm(35)
    section.left_margin = Mm(28)
    section.right_margin = Mm(26)

    # 文档网格（5.2.3）：python-docx 默认模板残留学生 linePitch=360(18磅) 且缺 type，
    # 需显式改写为 28 磅行网格（560 twips），与固定行距一致，避免 Word 各行版本解释不一致。
    _set_doc_grid(section, line_pitch_twips=int(LINE_BODY * 20))

    # 版头（红头 / 发文字号 / 分隔线）：源文档有就原样搬过来，不重新排版
    if content.get('header'):
        copy_letterhead(doc, content['header'])

    # 标题：小标宋 二号 居中 行距32磅
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(title_para, LINE_TITLE)
    title_run = title_para.add_run(title)
    set_font(title_run, FONT_TITLE, SIZE_TITLE)

    add_blank(doc)

    # 主送机关：顶格，全角冒号
    recipient = content.get('recipient')
    if recipient:
        if not recipient.endswith('：'):
            recipient += '：'
        rec_para = doc.add_paragraph()
        rec_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_line_spacing(rec_para, LINE_BODY)
        rec_run = rec_para.add_run(recipient)
        set_font(rec_run, FONT_BODY, SIZE_BODY)

    # 正文
    font_by_role = {'hei': FONT_HEI, 'kai': FONT_KAI, 'body': FONT_BODY}
    tables = content.get('tables') or []
    used_tables = set()

    for item in content.get('body', []):
        # 表格占位：{"type": "table", "index": 0}
        if isinstance(item, dict) and item.get('type') == 'table':
            idx = item.get('index', 0)
            if idx < len(tables):
                t = tables[idx]
                add_table(doc, t.get('rows', []),
                          header=t.get('header', True),
                          align=t.get('align', 'center'),
                          spans=t.get('spans'), ncol=t.get('ncol'))
                used_tables.add(idx)
            else:
                # 静默跳过等于丢内容，必须出声
                print(f'警告：body 引用了表格 index={idx}，'
                      f'但 tables 只有 {len(tables)} 个，该表格未输出')
            continue

        text, role = resolve_role(item)
        if role == 'blank' or not text.strip():
            add_blank(doc)
            continue
        font_role, bold = ROLE_STYLES[role]
        add_paragraph(doc, text, font_name=font_by_role[font_role], bold=bold)

    # 附件说明（7.3.4）
    attachments = content.get('attachments') or []
    if attachments:
        add_blank(doc)
        if len(attachments) == 1:
            text = '附件：' + attachments[0]
        else:
            text = '附件：' + '　'.join(f'{i}. {a}'
                                    for i, a in enumerate(attachments, 1))
        attach_para = add_paragraph(doc, text)
        # 回行与首字对齐：此处用悬挂缩进近似
        set_first_line_indent(attach_para, 2)

    # 发文机关署名与成文日期（7.3.5）
    signature = content.get('signature')
    date = content.get('date')
    if signature or date:
        add_blank(doc)
    if signature:
        sig_para = doc.add_paragraph()
        sig_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_line_spacing(sig_para, LINE_BODY)
        set_right_indent_chars(sig_para, 2)   # 右空二字
        sig_run = sig_para.add_run(signature)
        set_font(sig_run, FONT_BODY, SIZE_BODY)
    if date:
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_line_spacing(date_para, LINE_BODY)
        set_right_indent_chars(date_para, 4)   # 成文日期右空四字
        date_run = date_para.add_run(date)
        set_font(date_run, FONT_BODY, SIZE_BODY)

    # 未在 body 中显式定位的表格，追加到文末（落款之后）
    for idx, t in enumerate(tables):
        if idx not in used_tables:
            add_blank(doc)
            add_table(doc, t.get('rows', []),
                      header=t.get('header', True),
                      align=t.get('align', 'center'),
                      spans=t.get('spans'), ncol=t.get('ncol'))

    # 页码
    mode = resolve_page_number_mode(content.get('page_number'))
    if mode is None:
        raise ValueError(
            f"page_number 取值无法识别：{content.get('page_number')!r}，"
            f"可选 {PAGE_NUMBER_MODES}")
    add_page_number(section, mode)

    out = output_path or content.get('output')
    if not out:
        raise ValueError('未指定输出路径：请传 output_path 或在 content 中设置 output')
    out_dir = os.path.dirname(os.path.abspath(out))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(out)

    changed = {k: v for k, v in font_report.items() if v[0] != v[1]}
    if changed:
        print('字体回退：本机缺少国标字体，已自动替换为等效字体')
        for role, (old, new) in changed.items():
            print(f'  {old} → {new}')
        print('  （在装有国标字体的机器上生成则不会发生替换）')
    return out


# ── 字体自检 ─────────────────────────────────────────────

ROLE_LABEL = {'title': '标题', 'body': '正文/主送机关', 'kai': '第二层',
              'hei': '第一层', 'page': '页码'}


def check_fonts():
    """检查本机字体，并报告每个用途的回退链与最终选用的字体。"""
    avail = available_fonts()
    if not avail:
        return None          # 检测失败，无法判断
    resolved = resolve_fonts()
    report = {}
    for role, chain in FONT_CHAINS.items():
        report[role] = {
            'label': ROLE_LABEL[role],
            'chain': chain,
            'picked': resolved[role],
            'exact': resolved[role] == chain[0],   # 是否命中国标名
        }
    return report


# ── CLI ─────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='生成符合 GB/T 9704-2012 的公文 DOCX')
    parser.add_argument('--config', help='JSON 配置文件路径')
    parser.add_argument('--title', help='公文标题')
    parser.add_argument('--recipient', help='主送机关')
    parser.add_argument('--body-file', help='正文纯文本文件，每行一段')
    parser.add_argument('--attachments', nargs='*', help='附件名称列表')
    parser.add_argument('--signature', help='发文机关署名')
    parser.add_argument('--date', help='成文日期')
    parser.add_argument('--output', '-o', help='输出 .docx 路径')
    parser.add_argument('--page-number', dest='page_number',
                        choices=PAGE_NUMBER_MODES, default=None,
                        help='页码样式：center 居中（默认）/ odd_even 奇偶页不同 / none 不显示')
    parser.add_argument('--check-fonts', action='store_true', help='检查字体并退出')
    args = parser.parse_args()

    if args.check_fonts:
        result = check_fonts()
        if result is None:
            print('当前平台不支持自动字体检测，请手动确认已安装：'
                  '方正小标宋简体、仿宋_GB2312、楷体_GB2312、黑体、宋体')
            print('（未检测到字体清单时不会做回退替换，仍按国标名写入）')
        else:
            for role in FONT_CHAINS:
                r = result[role]
                mark = '✓' if r['exact'] else '↷'
                print(f"{mark} {r['label']}：实际使用「{r['picked']}」")
                if not r['exact']:
                    print(f"    国标名「{r['chain'][0]}」本机没有，"
                          f"已回退；回退链 {' → '.join(r['chain'])}")
        return 0

    if args.config:
        with open(args.config, encoding='utf-8') as f:
            content = json.load(f)
        # 命令行显式指定时覆盖配置文件
        if args.page_number:
            content['page_number'] = args.page_number
    else:
        if not args.title:
            parser.error('需要 --config，或至少提供 --title')
        body = []
        if args.body_file:
            with open(args.body_file, encoding='utf-8') as f:
                body = [line.strip() for line in f if line.strip()]
        content = {
            'title': args.title,
            'recipient': args.recipient,
            'body': body,
            'attachments': args.attachments or [],
            'signature': args.signature,
            'date': args.date,
            'page_number': args.page_number or PAGE_NUMBER_DEFAULT,
        }

    out = args.output or content.get('output')
    if not out:
        parser.error('未指定输出路径，请用 --output 指定')

    try:
        path = build_document(content, out)
    except ValueError as e:
        # 直接抛 traceback 对用户没有意义，给出一句能照着改的提示
        print(f'错误：{e}', file=sys.stderr)
        return 1
    print(f'公文已生成: {os.path.abspath(path)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())

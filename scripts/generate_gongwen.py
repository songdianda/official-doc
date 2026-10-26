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
import shutil
import sys
import tempfile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docx.text.paragraph import Paragraph
from lxml import etree

# 版记沉底要转 PDF 测量，渲染能力复用页码实测脚本（同为可选依赖）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from measure_page_number import find_soffice, to_pdf
except Exception:                      # noqa: BLE001 - 缺依赖就退回不校准
    find_soffice = to_pdf = None

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
    # 附件标题：附件另页起排时那一行标题，二号小标宋居中（GB/T 9704-2012 7.3.7）。
    # 印发类通知（"现将《XXX办法》印发给你们"）后面几乎都跟着一份附件，
    # 原来只能拿 heading3（仿宋加粗）凑，跟源文件二号小标宋居中差得远。
    'attachment_title': ('title', False),
}

# 居中排的角色：正文类角色一侓两端对齐，只有附件标题要居中
CENTER_ROLES = {'attachment_title'}

# 落款元素：放在 body 里表示"就在源文档的这个位置输出"，而不是甩到文末
SIGN_ROLES = {'signature', 'date'}

# 段落级 align 字段允许的取值（公文只用得到这四种）
ALIGN_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
    '两端对齐': WD_ALIGN_PARAGRAPH.JUSTIFY,
    '居中': WD_ALIGN_PARAGRAPH.CENTER,
    '居左': WD_ALIGN_PARAGRAPH.LEFT,
    '居右': WD_ALIGN_PARAGRAPH.RIGHT,
}


def resolve_align(value):
    """把 align 字段转成对齐枚举；未指定或无法识别返回 None（用角色默认值）。"""
    if value is None:
        return None
    return ALIGN_MAP.get(str(value).strip().lower())

# 旧字段名兼容
ROLE_ALIASES = {
    'level1': 'heading1', 'level2': 'heading2', 'level3': 'heading3',
    'level4': 'heading4', 'level5': 'heading5',
    'text': 'body', 'para': 'body', 'plain': 'body',
    'attachment': 'attachment_title', 'attachment_name': 'attachment_title',
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
    if role in SIGN_ROLES:
        return text, role          # 落款就地输出，保留位置
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


def set_pagination(paragraph, keep_next=False, keep_lines=False,
                   page_break_before=False):
    """段落分页控制：与下段同页 / 段中不分页 / 段前分页。

    「标题与后段同页」是 Word 里最常用的一个开关，藏在 w:pPr 里，
    看不到也不影响字体字号，但丢掉之后标题会孤零零留在页末——
    属于排版事故，不是美观问题。源文档设了就必须跟着走。

    schema 顺序：w:keepNext / w:keepLines / w:pageBreakBefore 必须排在
    w:numPr、w:spacing、w:ind、w:jc **之前**，直接 append 会被严格的
    解析器判为不合规。
    """
    pPr = paragraph._element.get_or_add_pPr()
    for tag, flag in (('w:keepNext', keep_next),
                      ('w:keepLines', keep_lines),
                      ('w:pageBreakBefore', page_break_before)):
        node = pPr.find(qn(tag))
        if flag:
            if node is not None:
                continue
            node = OxmlElement(tag)
            anchor = None
            for t in ('w:numPr', 'w:spacing', 'w:ind', 'w:jc'):
                anchor = pPr.find(qn(t))
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(node)
            else:
                pPr.append(node)
        elif node is not None:
            pPr.remove(node)


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


_IND_ATTRS = ('w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars',
              'w:hanging', 'w:hangingChars', 'w:right', 'w:rightChars')


def set_indent_chars(paragraph, left=None, hanging=None,
                     first_line=None, right=None):
    """按字符设置缩进，一次写全。

    python-docx 的 Normal 样式自带 firstLineChars=200，直接追加会叠加导致缩进翻倍，
    必须先清除已有缩进属性，再同时写入字符值与绝对值（Word 以字符值优先）。

    四个方向都可指定；未指定的方向会被清除，避免旧值与新值叠加。
      left       整段左缩进
      hanging    悬挂缩进：首行不动，后续行相对首行右移（用于回行对齐）
      first_line 首行缩进
      right      整段右缩进

    **注意：字符值的实际宽度由文档默认字号决定，不是当前段落字号。**
    用它排三号字公文时「2 字符」只有 7.8mm（按默认 11pt 折算），而国标要的是
    11.3mm。公文正文排版请用 `set_indent_pt`。本函数仅在确实需要按字符单位
    表达时保留。
    """
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    for attr in _IND_ATTRS:
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]
    for chars, char_attr, abs_attr in (
            (left, 'w:leftChars', 'w:left'),
            (hanging, 'w:hangingChars', 'w:hanging'),
            (first_line, 'w:firstLineChars', 'w:firstLine'),
            (right, 'w:rightChars', 'w:right'),
    ):
        if chars:
            ind.set(qn(char_attr), str(int(chars * 100)))
            ind.set(qn(abs_attr), str(int(chars * 320)))


def set_indent_pt(paragraph, left=None, right=None,
                  first_line=None, hanging=None):
    """按磅设置缩进，只写绝对值、不写字符值。

    用于页码「居右空一字」「居左空一字」：此处的「一字」按正文三号字（16 磅 =
    5.64mm）计。若同时写 rightChars，Word 会优先按字符值解释，而字符值按文档
    默认字号（多数模板是五号）折算，实测只有 3.8mm，达不到一字的 5.64mm。
    """
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    for attr in _IND_ATTRS:
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]
    for pt, attr in ((left, 'w:left'), (right, 'w:right'),
                     (first_line, 'w:firstLine'), (hanging, 'w:hanging')):
        if pt:
            ind.set(qn(attr), str(int(round(pt * 20))))    # 磅 → twips


def set_first_line_indent(paragraph, chars=2):
    """首行缩进 chars 个字符（按正文三号字折算成磅值写入）。

    只写绝对值、不写 firstLineChars：字符值的实际宽度由文档默认字号决定
    （python-docx 模板默认 11pt），用它排三号字的公文会得到 7.8mm 而不是
    11.3mm——实测渲染确认。写死磅值，「2 字符」在任何渲染器下都是 11.3mm。
    """
    set_indent_pt(paragraph, first_line=chars * SIZE_BODY)


def set_right_indent_chars(paragraph, chars=1):
    """右缩进 chars 个字符（按正文字号折算，用于署名右空二字等）。"""
    set_indent_pt(paragraph, right=chars * SIZE_BODY)


def set_left_indent_chars(paragraph, chars=1):
    """左缩进 chars 个字符（按正文字号折算）。"""
    set_indent_pt(paragraph, left=chars * SIZE_BODY)


def _is_wide_char(ch):
    """全角字符判定（汉字、全角标点、假名、谚文等）。"""
    o = ord(ch)
    return (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF
            or 0xAC00 <= o <= 0xD7A3 or 0xF900 <= o <= 0xFAFF
            or 0xFE30 <= o <= 0xFE6F or 0xFF00 <= o <= 0xFF60
            or 0xFFE0 <= o <= 0xFFE6)


def text_width_pt(text, size=SIZE_BODY):
    """估算一段文字的排版宽度（磅）：全角按字号、半角按半字号。

    只用来算落款与成文日期的相对位置。差半个字宽不影响观感，
    真要精确到毫米得渲染测量（scripts/measure_page_number.py 的做法）。
    """
    return sum(size if _is_wide_char(c) else size * 0.5 for c in (text or ''))


def sign_date_indents(signature, date, size=SIZE_BODY, right_chars=2):
    """落款：署名与成文日期共用一条中心线，日期以单位名称为中。

    GB/T 9704-2012 7.3.5 的字面是「署名右空二字、成文日期右空四字」，
    那是两组各自的右对齐——日期比署名长时会明显左探出去，盖印章时
    印章盖不住两者中间。实际操作里普遍要求**成文日期相对署名居中**。

    做法：取两者较宽者为基准右空二字，另一者按宽度差减半缩进，
    于是两条文字的中心落在同一条竖线上。

    返回 (署名右缩进磅值, 日期右缩进磅值)，都以版心右边缘为起点。
    """
    sw = text_width_pt(signature, size) if signature else 0.0
    dw = text_width_pt(date, size) if date else 0.0
    block = max(sw, dw, 0.0)
    center = right_chars * size + block / 2.0
    return (center - sw / 2.0, center - dw / 2.0)


def add_paragraph(doc, text, font_name=None, font_size=SIZE_BODY,
                  bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  indent=True, line_spacing=LINE_BODY,
                  keep_next=False, keep_lines=False,
                  page_break_before=False):
    # font_name 默认 None：不能在签名里写 FONT_BODY，默认参数在定义时就绑定，
    # 字体回退改写 FONT_BODY 后不会生效。
    para = doc.add_paragraph()
    para.alignment = align
    set_line_spacing(para, line_spacing)
    if indent:
        set_first_line_indent(para, 2)
    run = para.add_run(text)
    set_font(run, font_name or FONT_BODY, font_size, bold)
    set_pagination(para, keep_next, keep_lines, page_break_before)
    return para


def add_blank(doc, line_spacing=LINE_BODY):
    para = doc.add_paragraph()
    set_line_spacing(para, line_spacing)
    return para


# 附件说明用的序号前缀「1. 」= 半角数字 + 半角点 + 半角空格，约占 1.5 个汉字宽。
# 悬挂缩进要按这个宽度算，否则续行无法与附件名称首字对齐。
_NUM_PREFIX_CHARS = 1.5


def _add_signature_para(doc, text, right_pt=None):
    """发文机关署名：右对齐，缩进由 sign_date_indents 给出（默认右空二字）。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_line_spacing(para, LINE_BODY)
    set_indent_pt(para, right=(2 * SIZE_BODY if right_pt is None else right_pt))
    set_font(para.add_run(text), FONT_BODY, SIZE_BODY)
    return para


def _add_date_para(doc, text, right_pt=None):
    """成文日期：右对齐，缩进由 sign_date_indents 给出（默认右空四字）。

    与署名共用中心线时 right_pt 会比署名小——日期宽、署名窄，日期右探
    出去才是「居中于单位名称」的正确结果。
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_line_spacing(para, LINE_BODY)
    set_indent_pt(para, right=(4 * SIZE_BODY if right_pt is None else right_pt))
    set_font(para.add_run(text), FONT_BODY, SIZE_BODY)
    return para


def _add_attachment_para(doc, prefix, name, left, hanging):
    """附件说明一行。

    参数按 OOXML 语义：left 是**续行**的左边界，hanging 是首行相对续行
    **向左突出**的量（w:hanging 让首行左移）。因此首行起始 = left - hanging。
    想让首行在 S 字处、回行在 C 字处，应取 left=C、hanging=C-S。
    字数按正文三号字折算成磅值写入，避免被文档默认字号带偏。
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_line_spacing(para, LINE_BODY)
    set_indent_pt(para, left=left * SIZE_BODY, hanging=hanging * SIZE_BODY)
    run = para.add_run(prefix + name)
    set_font(run, FONT_BODY, SIZE_BODY)
    return para


# 发文字号：机关代字 + 〔年份〕+ 序号 + 号（GB/T 9704-2012 7.2.5）
DOC_NUMBER_RE = re.compile(r'^.{1,20}〔\d{4}〕\d+号$')

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

    整段 deepcopy w:p，字体、颜色、字号、边框一概照搬。
    红头是各单位的门面，本技能**不重做红头**——重排等于替用户改门面。

    唯一例外是**发文字号**：GB/T 9704-2012 7.2.4 要求它"居中排布"，
    而源文件常有没居中的（居左或随默认）。这是国标明文规定的位置，
    不是审美偏好，所以照国标纠正，并在生成报告里说明改了什么。
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
    centered = []

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

        # 发文字号按国标居中（7.2.4）
        para = Paragraph(new_p, doc)
        if DOC_NUMBER_RE.match(para.text.strip()):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            centered.append(para.text.strip())

    if copied:
        _report_letterhead_fonts(src, indices)
    if centered:
        print(f'✓ 发文字号已按国标居中排布（7.2.4）：'
              + '、'.join(centered))
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
# 默认 none：源文档没有页码就不许凭空加。
# 页码属于源文件"有没有"的事实，不是排版偏好——从无到有是新增内容，
# 与"原文一个字不增不减"冲突。要加必须由使用者显式指定。
PAGE_NUMBER_DEFAULT = 'none'

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


# 输出文件已存在时的处理方式。默认 rename：
# 产出物常被人工二次修改（补内容、调措辞），直接覆盖会让这些修改无声消失，
# 而"文件被覆盖"是不可撤销的。宁可多出一个 _2 文件，也不能让人白改一遍。
#   rename     自动另存为「原名_2.docx」「_3.docx」…（默认）
#   fail       停止生成并提示，交给使用者决定
#   overwrite  覆盖同名文件（仅在明确要求时使用）
ON_EXIST_MODES = ('rename', 'fail', 'overwrite')
ON_EXIST_DEFAULT = 'rename'
ON_EXIST_ALIASES = {
    'rename': 'rename', 'new': 'rename', 'auto': 'rename', '另存': 'rename',
    '另存为': 'rename', '新建': 'rename',
    'fail': 'fail', 'stop': 'fail', 'abort': 'fail', 'error': 'fail',
    '停止': 'fail', '中止': 'fail', '提示': 'fail',
    'overwrite': 'overwrite', 'replace': 'overwrite', 'force': 'overwrite',
    '覆盖': 'overwrite', '替换': 'overwrite',
}


# 保真模式：strict = 原文一个字不增不减（默认）。
# standard 仅用于"本来就要把 Word 自动编号固化成字符"的场景。
FIDELITY_MODES = ('strict', 'standard')
FIDELITY_DEFAULT = 'strict'
FIDELITY_ALIASES = {
    'strict': 'strict', '严格': 'strict', '保真': 'strict',
    'standard': 'standard', '标准': 'standard', '补编号': 'standard',
}


def resolve_fidelity(value):
    """保真模式归一：strict（默认）/ standard；无法识别返回 None。"""
    if value is None:
        return FIDELITY_DEFAULT
    key = str(value).strip().lower()
    key = FIDELITY_ALIASES.get(key, key)
    return key if key in FIDELITY_MODES else None


def resolve_on_exist(value):
    """把配置值归一到 ON_EXIST_MODES 之一；无法识别时返回 None。"""
    if value is None:
        return ON_EXIST_DEFAULT
    key = str(value).strip().lower()
    key = ON_EXIST_ALIASES.get(key, key)
    return key if key in ON_EXIST_MODES else None


def resolve_output_path(path, on_exist):
    """按 on_exist 策略确定最终写入路径。

    rename：找第一个不存在的「原名_N」；fail：抛 FileExistsError；
    overwrite 或文件本不存在：原样返回。
    """
    if on_exist not in ON_EXIST_MODES:
        raise ValueError(f'on_exist 取值无法识别：{on_exist!r}，可选 {ON_EXIST_MODES}')
    if not os.path.exists(path):
        return path, None
    if on_exist == 'overwrite':
        return path, None
    if on_exist == 'fail':
        raise FileExistsError(path)
    stem, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f'{stem}_{n}{ext}'):
        n += 1
        if n > 999:                      # 兜底，避免异常情况下死循环
            raise RuntimeError(f'同名文件过多，无法确定新文件名：{path}')
    return f'{stem}_{n}{ext}', path


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


# 页脚距页底的距离（mm）。
# 国标 7.5：页码一字线的上沿距版心下边缘 7mm。版心下边缘距页底 35mm，
# 故页码应在距页底 28mm 附近；又因页脚内容自页脚顶边向下排、且字形上沿
# 距行顶尚有一段空隙，实测校准后取 23.58mm。
# 校准方法：生成多页样本 → LibreOffice 转 PDF → 按 300dpi 渲染扫描墨迹，
# 量出一字线上沿相对版心下边缘的距离（原取 4.58mm 时实测 26.00mm，偏差 +19.00mm）。
FOOTER_DISTANCE_MM = 23.58


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
        section.footer_distance = Mm(FOOTER_DISTANCE_MM)
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

    # 2) 奇数页页脚（default）：居右空一字。
    #    「一字」按正文三号字（16 磅）算，故用磅值而非字符值
    footer_odd = section.footer
    footer_odd.is_linked_to_previous = False
    _make_footer_paragraph(
        footer_odd, WD_ALIGN_PARAGRAPH.RIGHT,
        lambda p: set_indent_pt(p, right=SIZE_BODY))

    # 3) 偶数页页脚（even）：居左空一字。取不到时降级，仅保留奇数页页码
    footer_even = _get_even_footer(section)
    if footer_even is not None:
        footer_even.is_linked_to_previous = False
        _make_footer_paragraph(
            footer_even, WD_ALIGN_PARAGRAPH.LEFT,
            lambda p: set_indent_pt(p, left=SIZE_BODY))

    # 4) 页脚位置：见 FOOTER_DISTANCE_MM（实测校准，非估算）
    section.footer_distance = Mm(FOOTER_DISTANCE_MM)


def _get_even_footer(section):
    """获取偶数页页脚。python-docx 未公开 even footer，需内部构造；失败则返回 None 降级。"""
    try:
        from docx.enum.section import WD_HEADER_FOOTER
        from docx.section import _Footer
        return _Footer(section._sectPr, section.part, WD_HEADER_FOOTER.EVEN_PAGE)
    except Exception:
        return None


# ── 版记（印发机关 + 印发日期，GB/T 9704-2012 7.4）──────

TYPE_BOTTOM_MM = 262.0        # 版心下边缘 = 297 - 35
VERSION_BLOCK_MM = 40.0       # 版记块（上线条 + 印发行 + 下线条）占的高度
# 版记底线的落点。国标 7.4.2/7.4.3 要版记排在版心最下方、末条分隔线为版记
# 的最后一行，理论上应当贴着版心下边缘（262mm）。这里收在 252mm——留一格
# （正文一行 28 磅 ≈ 9.9mm）余量。原因：成品在本机是用替代字体渲染的，用户
# 在装有国标字体的 Word 里打开时行数可能差一行；贴死到 262 一旦多出一行，
# 版记就会被整体挤到新的一页，出现"末页只有版记"的硬伤。留一格只表现为
# 底部多一点白，不会翻页。
VERSION_BOTTOM_LINE_MM = TYPE_BOTTOM_MM - 10.0
# 印发机关那一行文字的 bbox 底边，比下分隔线低约 0.78mm（实测：
# 文字 y1=109.33，下分隔线 y=108.55）。用来在下分隔线量不到时兜底换算。
VERSION_TEXT_TO_LINE_MM = 0.8


def _spacer_p(before_pt=0.0):
    """一个只有 1 磅行高、可带段前间距的空段落——把版记往下推一格。

    为什么不像早先那样堆一串 28 磅空段落：空段在分页边界上会被渲染器静默
    丢掉几段（实测设 12 段只推进到 183mm、设 14 段落在 217mm），垫多少都不
    确定；而且空段高度还受文档网格影响，不是 28 磅。

    改成一个空段 + 一个 `w:before` 数值后，位移与这个数值是严格 1:1 的
    （实测 before=200 磅 → 版记下移 70.56mm = 200 磅），写多少就是多少。
    """
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    snap = OxmlElement('w:snapToGrid')
    snap.set(qn('w:val'), 'false')
    pPr.append(snap)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(int(round(before_pt * 20))))
    sp.set(qn('w:after'), '0')
    sp.set(qn('w:line'), '20')           # 1 磅固行距
    sp.set(qn('w:lineRule'), 'exact')
    pPr.append(sp)
    rPr = OxmlElement('w:rPr')           # 1 磅字，别让行高被字号撑起来
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '2')
    rPr.append(sz)
    pPr.append(rPr)
    p.append(pPr)
    return p


def add_version_record(doc, vr):
    """把版记整块原样搬到文末，返回 (首元素, 末元素)。

    blocks 是源文档那几段的原始 XML（上分隔线段 + 印发行，下分隔线在印发行
    自身里）。整块搬而不是重排：版记是单位自己排好的，字号、印发机关与印发
    日期之间的间距、两条分隔线都属"不要给变动"的部分——重排一次就变一次样。
    """
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    made = []
    for xml in (vr.get('blocks') or []):
        el = etree.fromstring(xml.encode('utf-8'))
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            body.append(el)
        made.append(el)
    # 版记是"上线条 + 印发行 + 下线条"一个整体，不能被拆到两页去：
    # 不设「与下段同页」时，遇到分页边界会出现上半截在前一页、
    # 印发机关掉到下一页的情况（实测垫到某几段数时位置会突然跳上来）。
    for el in made[:-1]:
        pPr = el.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            el.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            keep = OxmlElement('w:keepNext')
            # 顺序：keepNext 必须排在 spacing / ind / jc 之前
            anchor = None
            for tag in ('w:spacing', 'w:ind', 'w:jc', 'w:rPr'):
                anchor = pPr.find(qn(tag))
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(keep)
            else:
                pPr.append(keep)

    # 版记不吸附文档网格。国标正文是 28 磅网格，而版记是四号字、1.4 倍行距，
    # 吸到 28 磅网格上行会变高，那两条按源文档相对位置画的 VML 分隔线就会
    # 穿到字上（实测印发机关行底部压在下横线上）。关掉吸附，行高回到源样，
    # 线条位置也就回到了源样——这同样是"不要给变动"的一部分。
    for el in made:
        pPr = el.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            el.insert(0, pPr)
        if pPr.find(qn('w:snapToGrid')) is None:
            snap = OxmlElement('w:snapToGrid')
            snap.set(qn('w:val'), 'false')
            # schema 顺序：snapToGrid 在 spacing / ind / jc 之前
            anchor = None
            for tag in ('w:spacing', 'w:ind', 'w:jc', 'w:rPr'):
                anchor = pPr.find(qn(tag))
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(snap)
            else:
                pPr.append(snap)
    return (made[0] if made else None), (made[-1] if made else None)


def _vrecord_head_el(doc, marker):
    """版记块的首元素：印发行，或它前面那条带 VML 线条的上分隔线段。

    沉底垫的空段落全部插在这个元素之前，删的时候也只会删到它为止——
    不会误删正文。
    """
    p = None
    for para in doc.paragraphs:
        if marker and marker in (para.text or ''):
            p = para
            break
    if p is None:
        return None
    el = p._element
    prev = el.getprevious()
    if prev is not None and prev.tag == qn('w:p'):
        txt = ''.join(t.text or '' for t in prev.iter(qn('w:t'))).strip()
        has_shape = any(True for _ in prev.iter('{%s}shape' % VML_NS)) \
            or any(True for _ in prev.iter('{%s}line' % VML_NS))
        if not txt and has_shape:
            return prev
    return el


def _vrecord_probes(marker):
    """版记行的匹配串，从长到短。

    长串用来避免误命中正文（正文里也可能出现「XX师范大学」）；
    短串用来兜底两种情况：PDF 里这一行可能被拆成两块（印发机关一块、
    印发日期一块），以及版记被推到新一页时 PyMuPDF 抽文字会掉字
    （实测「XX师范大学党政办公室」抽成「X师范大学党政办公室X」，
    前 4 个字全废，只有中段的「党政办公室」仍然完整）。
    """
    m = re.sub(r'\s', '', marker or '')
    out = []
    for cand in (m[:10], m[:8], m[:6], '党政办公室', '印发日期', '印发'):
        if len(cand) >= 3 and cand not in out and cand in (m or cand):
            out.append(cand)
    return out


def _measure_vrecord(docx_path, marker):
    """量版记位置，返回 (页数, 印发行文字底边 mm, 下分隔线 mm)。

    分隔线按形状找而不是按文字找：版记上下两条线与版心等宽（156mm，x 从
    28 到 184mm），末页里没有别的横线会这么长，所以「版心宽的横线」就是
    版记的分隔线，取最下面那条即下分隔线；再与文字底边互相校验，防止把
    表格线一类的东西误当分隔线。
    量不出来返回 None（量不出来就是没验，不拿估计值冒充）。
    """
    if to_pdf is None:
        return None, None, None
    try:
        import pymupdf as fitz
    except ImportError:
        return None, None, None
    probes = _vrecord_probes(marker)
    tmp = tempfile.mkdtemp(prefix='vr_measure_')
    try:
        pdf = to_pdf(docx_path, tmp)
        doc = fitz.open(pdf)
        pages = len(doc)
        page = doc[-1]
        px = 72 / 25.4
        text_y1 = None
        for b in page.get_text('blocks'):
            y1_mm = b[3] / px
            # 只看版心以内：页码在 268mm 附近，混进来会把版记量成页脚
            if y1_mm > TYPE_BOTTOM_MM + 0.5:
                continue
            txt = re.sub(r'\s', '', b[4] or '')
            if any(pt in txt for pt in probes):
                text_y1 = y1_mm if text_y1 is None else max(text_y1, y1_mm)
        # 版记底线：版心宽的横线，且与印发行紧邻（上线条在文字上方约 15mm）
        cands = []
        min_w = 140 * px
        for drawing in page.get_drawings():
            for item in drawing.get('items') or []:
                y = None
                if item[0] == 'l':
                    p1, p2 = item[1], item[2]
                    if abs(p1.y - p2.y) < 0.5 and abs(p2.x - p1.x) >= min_w:
                        y = p1.y / px
                elif item[0] == 're':
                    r = item[1]
                    if r.height < 1.5 and r.width >= min_w:
                        y = r.y1 / px
                if y is None:
                    continue
                if text_y1 is not None and not (text_y1 - 30 <= y <= text_y1 + 8):
                    continue
                cands.append(y)
        line_y = round(max(cands), 2) if cands else None
        doc.close()
        return pages, (round(text_y1, 2) if text_y1 is not None else None), line_y
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _vrecord_pos(line_y, text_y1):
    """版记当前底线位置：优先用下分隔线，量不到用文字底边换算。"""
    if line_y is not None:
        return line_y
    if text_y1 is None:
        return None
    return text_y1 - VERSION_TEXT_TO_LINE_MM


def _measure_vrecord_robust(docx_path, marker, n=5):
    """多测几次取**最大值**，返回 (页数, 底线位置 mm)。

    为什么要多测：同一个 docx 连转 8 次，约 1/5 的次数末页内容会整体上移
    3 行（下分隔线 255.3mm → 225.7mm），是 LibreOffice 在分页边界上的偶发
    重排，改 profile、去 keepNext 都没能消除。取最大值就是取"版心排得最满"
    的那个布局——它占了约 4/5 的次数，是渲染器的常态结果；偶发的上移变体
    只会在成品里表现为多留一行空白，不会把版记顶到页外。

    页数也取最大值：宁可认为页数更多，也不要把溢出看成没溢出。
    """
    pages = None
    vals = []
    for _ in range(max(1, n)):
        p, y1, ln = _measure_vrecord(docx_path, marker)
        if p is not None:
            pages = p if pages is None else max(pages, p)
        v = _vrecord_pos(ln, y1)
        if v is not None:
            vals.append(v)
    if not vals:
        return pages, None
    return pages, max(vals)


def sink_version_record(docx_path, marker):
    """把版记沉到末页版心最下方，返回 (是否成功, 说明)。

    国标 7.4.2：版记排在最后一面版心最下方（版记最后一个要素在最后一行）。
    Word 没有"贴底"属性，只能靠留白把版记推下去。这里在版记前插**一个**
    1 磅高的空段落，并给它一个 `w:before` 段前间距——一个 twips 数值表示
    全部留白（见 _spacer_p 的说明）。

    四次踩坑决定了现在的写法：
    1. "还差多少就再补多少"式迭代会失控——版记被挤到新一页后，新页顶部的大片
       空白又变成"还差很多"，越垫越多（实测把 6 页垫成 8 页）。
    2. 空段推进量不可信——渲染器在分页边界上会静默丢空段（设 12 段只到 183mm、
       设 14 段落在 217mm），按段数精算不成立。改成一个空段 + 一个段前间距后，
       位移与数值严格 1:1（实测 before=200 磅 → 下移 70.56mm）。
    3. 早先复用同一个 LibreOffice profile，同一个 docx 连转 5 次末页落点会在
       109.3 / 108.3 / 78.6mm 之间跳；改成每次全新 profile 后不再抖。
    4. 即使 profile 独立，仍有约 1/5 的概率末页整体上移 3 行（去 keepNext、
       换字体都试过，消不掉），所以测量取多次里的最大值，见
       _measure_vrecord_robust。

    也试过给版记加 `w:framePr`（yAlign=bottom 锚定版心底部，本可完全不依赖
    正文流），但 LibreOffice 直接忽略 y 定位属性，没法在本机验证渲染结果，
    故不采用——验证不了的排版不写进成品。
    """
    if find_soffice is None or find_soffice() is None:
        return None, '未找到 LibreOffice，未做沉底校准（版记已排在文末）'
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        return None, '未安装 PyMuPDF，未做沉底校准（版记已排在文末）'
    target = VERSION_BOTTOM_LINE_MM
    pt_per_mm = 72 / 25.4

    tmpdir = tempfile.mkdtemp(prefix='vr_base_')
    base = os.path.join(tmpdir, 'base.docx')
    shutil.copy(docx_path, base)

    def apply(before_pt):
        """从基线重建，插一个带段前间距的空段。"""
        shutil.copy(base, docx_path)
        doc = Document(docx_path)
        head = _vrecord_head_el(doc, marker)
        if head is None:
            return False
        head.addprevious(_spacer_p(before_pt))
        doc.save(docx_path)
        return True

    try:
        pages0, cur0 = _measure_vrecord_robust(docx_path, marker, 5)
        if cur0 is None:
            return None, '成品里定位不到版记行，未做沉底校准'
        if target - cur0 <= 0.8:
            return True, (f'版记已在最下方，未加段前间距'
                          f'（下分隔线 {cur0:.1f}mm，{pages0} 页）')

        before = (target - cur0) * pt_per_mm
        if not apply(before):
            return None, '成品里定位不到版记行，未做沉底校准'
        pages_c, cur = _measure_vrecord_robust(docx_path, marker, 3)
        if cur is None:
            shutil.copy(base, docx_path)
            return None, '沉底后定位不到版记行，已按未沉底保留'

        # 溢出：版记被推过页尾落到新一页（间距在新页页顶会被丢掉，位置反而
        # 更靠上）。此时在 0 与当前值之间二分，取"不溢出里最靠下"的那个。
        lo, hi = 0.0, None
        for _ in range(3):
            if pages_c is None or pages_c <= pages0:
                break
            hi = before
            before = (lo + hi) / 2
            apply(before)
            pages_c, cur = _measure_vrecord_robust(docx_path, marker, 1)
        if pages_c is not None and pages_c > pages0:
            shutil.copy(base, docx_path)
            return False, (f'版记加段前间距会掉到新的一页（{pages_c} 页），'
                           f'已按原位置保留，请目视核对')

        # 毫米级修正一次即可（1:1 位移，实测残差 0.3mm 内）
        if abs(target - cur) > 1.5:
            prev = before
            before = max(0.0, before + (target - cur) * pt_per_mm)
            apply(before)
            pages_v, cur_v = _measure_vrecord_robust(docx_path, marker, 3)
            if cur_v is None or (pages_v is not None and pages_v > pages0):
                # 修正后反而溢出，退回上一版
                before = prev
                apply(before)
            else:
                cur = cur_v

        tw = int(round(before * 20))
        return True, (f'版记贴底：下分隔线 {cur:.1f}mm'
                      f'（版心下边缘 {TYPE_BOTTOM_MM:.0f}mm，'
                      f'距底 {TYPE_BOTTOM_MM - cur:.1f}mm），'
                      f'段前间距 {tw / 20:.1f} 磅，{pages0} 页')
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _measure_block_centers(docx_path, signature, date):
    """量落款署名与成文日期两条文字的中心线（毫米），量不出来返回 (None, None)。"""
    if to_pdf is None:
        return None, None
    try:
        import pymupdf as fitz
    except ImportError:
        return None, None
    want_s = re.sub(r'\s', '', signature or '')
    want_d = re.sub(r'\s', '', date or '')
    tmp = tempfile.mkdtemp(prefix='vr_sign_')
    try:
        pdf = to_pdf(docx_path, tmp)
        doc = fitz.open(pdf)
        px = 72 / 25.4
        got_s = got_d = None
        for page in doc:
            for b in page.get_text('blocks'):
                txt = re.sub(r'\s', '', b[4] or '')
                c = (b[0] + b[2]) / 2 / px
                # 只认整块就是这一条的文字：正文里「XX师范大学家具管理办法」
                # 这种包含关系不能算
                if want_s and txt == want_s and got_s is None:
                    got_s = c
                elif want_d and txt == want_d and got_d is None:
                    got_d = c
            if got_s is not None and got_d is not None:
                break
        doc.close()
        return got_s, got_d
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def align_date_to_signature(docx_path, signature, date, tol_mm=0.8, rounds=2):
    """让成文日期的中心线落在发文机关署名的中心上。

    sign_date_indents 是按估算字宽算缩进的，但中文与数字混排时渲染器会
    额外加间距（实测日期实际比估算宽 12 磅），中心会差 2mm 左右。
    所以最后一步改成量出来再挪：量两条文字的中心，差多少就把日期段落的
    右缩进改多少，最多两轮。
    """
    if not signature or not date:
        return None, '只有一侧有落款，无需居中'
    last = None
    for _ in range(rounds):
        cs, cd = _measure_block_centers(docx_path, signature, date)
        if cs is None or cd is None:
            return None, '渲染里定位不到落款，未做日期居中校准'
        diff = cs - cd                     # 正数=日期偏左，要往右挪
        last = diff
        if abs(diff) <= tol_mm:
            return True, f'日期中心与单位名称中心偏差 {abs(diff):.2f}mm'
        doc = Document(docx_path)
        tgt = re.sub(r'\s', '', date)
        hit = None
        for p in doc.paragraphs:
            if re.sub(r'\s', '', p.text or '') == tgt:
                hit = p
                break
        if hit is None:
            return None, '成品里定位不到成文日期段落'
        pPr = hit._element.find(qn('w:pPr'))
        ind = pPr.find(qn('w:ind')) if pPr is not None else None
        if ind is None:
            return None, '成文日期段落没有缩进设置，未做校准'
        cur = int(ind.get(qn('w:right')) or 0)
        # 中心右移 diff 毫米 = 右缩进减少同样距离（1mm = 56.7 twips）
        new = max(0, cur - int(diff * 72 / 25.4 * 20))
        ind.set(qn('w:right'), str(new))
        doc.save(docx_path)
    cs, cd = _measure_block_centers(docx_path, signature, date)
    if cs is None or cd is None:
        return None, '校准后定位不到落款'
    return abs(cs - cd) <= tol_mm, f'日期与署名中心偏差 {abs(cs - cd):.2f}mm'


# ── 主体构建 ─────────────────────────────────────────────


def validate_content(content):
    """生成前集中校验输入数据，一次性报出全部问题。

    只做标题检查远远不够：body 引用了不存在的表格、tables 里混进数字单元格、
    page_number 拼错，都会产出一份内容缺失或格式不对的文档。
    这里把类型与取值一次查完，报出具体字段和原因，而不是生成到一半才失败。

    返回警告列表（不阻断的问题，如未知 type 降级为正文）。
    发现阻断性问题时抛 ValueError，错误信息含全部问题。
    """
    errors = []
    warns = []
    if not isinstance(content, dict):
        raise TypeError(f'content 必须是 dict，实际是 {type(content).__name__}')

    def _str_field(name, maxlen):
        v = content.get(name)
        if v is None:
            return None
        if not isinstance(v, str):
            errors.append(f'{name}：应为字符串，实际是 {type(v).__name__}'
                          f'（值 {v!r}）')
            return None
        if maxlen and len(v) > maxlen:
            warns.append(f'{name}：长度 {len(v)} 超过 {maxlen}，'
                         f'可能是误把大段文本填进了该字段')
        return v

    _str_field('title', 200)
    _str_field('recipient', 200)
    _str_field('signature', 100)
    _str_field('date', 50)
    _str_field('output', 500)

    tkn = content.get('title_keep_next')
    if tkn is not None and not isinstance(tkn, bool):
        warns.append(f'title_keep_next：应为 true / false，实际是 {tkn!r}')

    tables = content.get('tables')
    if tables is None:
        tables = []
    elif not isinstance(tables, list):
        errors.append(f'tables：应为列表，实际是 {type(tables).__name__}')
        tables = []
    else:
        for ti, t in enumerate(tables):
            if not isinstance(t, dict):
                errors.append(f'tables[{ti}]：应为 dict，实际是 {type(t).__name__}')
                continue
            rows = t.get('rows')
            if rows is None:
                errors.append(f'tables[{ti}]：缺少 rows')
                continue
            if not isinstance(rows, list) or not rows:
                errors.append(f'tables[{ti}].rows：应为非空列表')
                continue
            widths = set()
            for ri, row in enumerate(rows):
                if not isinstance(row, list):
                    errors.append(f'tables[{ti}].rows[{ri}]：应为列表，'
                                  f'实际是 {type(row).__name__}')
                    continue
                widths.add(len(row))
                for ci, cell in enumerate(row):
                    if not isinstance(cell, str):
                        errors.append(
                            f'tables[{ti}].rows[{ri}][{ci}]：应为字符串，'
                            f'实际是 {type(cell).__name__}（值 {cell!r}）；'
                            f'数字请写成 "300" 而不是 300')
            if len(widths) > 1:
                warns.append(f'tables[{ti}]：各行列数不一致 {sorted(widths)}，'
                             f'缺的格子会留空')
            spans = t.get('spans')
            if spans is not None:
                if not isinstance(spans, list) or len(spans) != len(rows):
                    errors.append(
                        f'tables[{ti}].spans：应为与 rows 等长的列表，'
                        f'实际 rows {len(rows)} 行、spans '
                        f'{len(spans) if isinstance(spans, list) else "非列表"}')

    body = content.get('body')
    if body is None:
        body = []
    elif not isinstance(body, list):
        errors.append(f'body：应为列表，实际是 {type(body).__name__}')
        body = []
    else:
        for bi, item in enumerate(body):
            if isinstance(item, str):
                continue
            if not isinstance(item, dict):
                errors.append(f'body[{bi}]：应为字符串或 dict，'
                              f'实际是 {type(item).__name__}')
                continue
            typ = item.get('type')
            if typ == 'table':
                idx = item.get('index', 0)
                if not isinstance(idx, int) or isinstance(idx, bool):
                    errors.append(f'body[{bi}].index：应为整数，实际是 {idx!r}')
                elif not (0 <= idx < len(tables)):
                    errors.append(
                        f'body[{bi}].index={idx}：超出 tables 范围'
                        f'（tables 共 {len(tables)} 个，合法下标 '
                        f'0..{max(len(tables) - 1, 0)}）')
                continue
            if typ is None:
                errors.append(f'body[{bi}]：dict 形式必须带 type 字段')
                continue
            if typ == 'blank':
                # 空行是合法角色（不渲染文字），不参与角色检查、也不强求 text
                continue
            if typ not in ROLE_ALIASES and typ not in ROLE_STYLES \
                    and typ not in SIGN_ROLES:
                warns.append(f'body[{bi}].type={typ!r}：不是已知角色，'
                             f'已按正文处理（可用 '
                             f'{", ".join(sorted(ROLE_STYLES))}）')
            if 'text' not in item:
                errors.append(f'body[{bi}]：缺少 text 字段')
            elif not isinstance(item['text'], str):
                errors.append(f'body[{bi}].text：应为字符串，'
                              f'实际是 {type(item["text"]).__name__}')
            for key in ('keep_next', 'keep_lines', 'page_break_before'):
                if key in item and not isinstance(item[key], bool):
                    warns.append(f'body[{bi}].{key}：应为 true / false，'
                                 f'实际是 {item[key]!r}')

    attachments = content.get('attachments')
    if attachments is not None:
        if not isinstance(attachments, list):
            errors.append(f'attachments：应为列表，'
                          f'实际是 {type(attachments).__name__}')
        else:
            for ai, a in enumerate(attachments):
                if not isinstance(a, str):
                    errors.append(f'attachments[{ai}]：应为字符串，'
                                  f'实际是 {type(a).__name__}（值 {a!r}）')

    pn = content.get('page_number')
    if pn is not None and resolve_page_number_mode(pn) is None:
        errors.append(f'page_number：取值 {pn!r} 无法识别，可选 {PAGE_NUMBER_MODES}')

    oe = content.get('on_exist')
    if oe is not None and resolve_on_exist(oe) is None:
        errors.append(f'on_exist：取值 {oe!r} 无法识别，可选 {ON_EXIST_MODES}')

    header = content.get('header')
    if header is not None:
        if not isinstance(header, dict):
            errors.append(f'header：应为 dict，实际是 {type(header).__name__}')
        else:
            src = header.get('source')
            if not src:
                errors.append('header：缺少 source（源文件路径）')
            elif not os.path.exists(src):
                errors.append(f'header.source：源文件不存在 —— {src}；'
                              f'版头将无法保留')
            idxs = header.get('indices')
            if not idxs:
                errors.append('header：缺少 indices（版头段落下标）')
            elif (not isinstance(idxs, list)
                  or not all(isinstance(i, int) and not isinstance(i, bool)
                             for i in idxs)):
                errors.append(f'header.indices：应为整数列表，实际是 {idxs!r}')

    fid = content.get('fidelity')
    if fid is not None and resolve_fidelity(fid) is None:
        errors.append(f'fidelity：取值 {fid!r} 无法识别，可选 {FIDELITY_MODES}')

    # 严格保真：任何"源文档里没有的字符"都要在生成前说出来
    if resolve_fidelity(fid) == 'strict':
        items = body if isinstance(body, list) else []
        npref = [bi for bi, it in enumerate(items)
                 if isinstance(it, dict) and it.get('num_prefix')]
        if npref:
            warns.append(
                f'严格保真：body 中 {len(npref)} 段带 num_prefix'
                f'（Word 自动编号，源文件里不是字符），未并入正文 —— '
                f'成品中看不到这些编号。需要显示请用 --fidelity standard 重新提取。')
        if attachments:
            warns.append(
                '严格保真：attachments 会由生成器补「附件：」与序号字符，'
                '这些字源文档里没有。如需一字不增，请把附件说明原段放进 body，'
                '并清空 attachments。')

    if errors:
        raise ValueError(
            f'输入数据校验未通过，共 {len(errors)} 处：\n  - '
            + '\n  - '.join(errors))
    return warns


def build_document(content, output_path, on_exist=None):
    """根据 content 字典生成公文 DOCX。

    输出文件已存在时默认**不覆盖**，自动另存为「原名_2.docx」并返回新路径，
    以免覆盖掉人工二次修改的内容。可用 on_exist 改变该行为：
    rename（默认）/ fail（报错停止）/ overwrite（覆盖）。

    content 支持字段：
      title       标题（必填）
      recipient   主送机关，如「XX省人民政府：」（可选）
      body        正文列表，元素可为字符串（兜底识别）或
                  {"type": "heading1".."heading5"|"body"|"numbered_body"
                          |"blank"|"table", "text": "..."}
                  {"type": "table", "index": 0} 为附表占位
      tables      附表列表 [{"rows": [...], "header": true,
                             "spans": [...], "ncol": 3}]
      attachments 附件名称列表（可选）
      signature   发文机关署名（可选）
      date        成文日期（可选）
      page_number 页码样式：none（默认，不添加）/ center（居中）/
                  odd_even（奇偶页不同）
                  有源文件时由 inspect_docx.py 按源文档推断：源无页码→none，
                  源有页码→沿用。从无到有属于新增内容，必须显式指定才加
      fidelity    保真模式：strict（默认，原文一字不增不减，
                  自动编号不并入正文）/ standard（补回自动编号）
      output      输出路径（可选，优先级低于 output_path 参数）
      on_exist    输出文件已存在时的处理：rename（默认，自动另存）/ fail / overwrite
                  （同名字段优先级低于 on_exist 参数）
    """
    title = (content.get('title') or '').strip()
    if not title:
        raise ValueError('缺少 title：公文标题为必填项，'
                         '请传 --title 或在 content 中设置 title')
    # 其余字段一次查完：类型错、取值越界都会产出内容缺失的文档，
    # 与其生成后才发现，不如生成前集中报出
    for w in validate_content(content):
        print(f'警告：{w}')
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
    # 源标题设了「与下段同页」就照写：标题掉在页末是事故，不能因为重排就丢
    set_pagination(title_para,
                   keep_next=bool(content.get('title_keep_next')))

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
    font_by_role = {'hei': FONT_HEI, 'kai': FONT_KAI, 'body': FONT_BODY,
                    'title': FONT_TITLE}
    tables = content.get('tables') or []
    used_tables = set()
    sign_in_body = False                 # 落款是否已在 body 原位出现
    blank_run = 0                        # 连续空段计数

    # 署名与成文日期要共用一条中心线，得先知道两者的文字（不管落在 body
    # 里还是顶层字段），才能在渲染前把两个缩进一次算出来。
    _sig_text = _date_text = None
    for _it in content.get('body', []):
        if isinstance(_it, dict):
            _r = ROLE_ALIASES.get(_it.get('type'), _it.get('type'))
            if _r == 'signature':
                _sig_text = _it.get('text') or ''
            elif _r == 'date':
                _date_text = _it.get('text') or ''
    _sig_text = _sig_text or content.get('signature') or ''
    _date_text = _date_text or content.get('date') or ''
    _sig_right, _date_right = sign_date_indents(_sig_text, _date_text)

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
                # 引用了不存在的表格 = 会交付一份缺表的公文，属内容丢失。
                # 不能只警告后继续，必须停下来让使用者补齐。
                raise ValueError(
                    f'body 引用了表格 index={idx}，但 tables 只有 {len(tables)} 个'
                    f'（合法下标 0..{len(tables) - 1}）。'
                    f'请补全 tables 或删除该引用后再生成——'
                    f'继续生成会产出缺表的文档。')
            continue

        text, role = resolve_role(item)

        # 落款就地输出：源文档把它放在正文中部（通知正文之后、附件之前）
        # 是常态，统一甩到文末会让成文日期跑到最后一页。
        if role == 'signature':
            sign_in_body = True
            blank_run = 0
            _add_signature_para(doc, text, _sig_right)
            continue
        if role == 'date':
            sign_in_body = True
            blank_run = 0
            _add_date_para(doc, text, _date_right)
            continue

        if role == 'blank' or not text.strip():
            # 连续空行压成一个。源文件常在落款前后、附件前后堆好几行空段落，
            # 剥离落款之后它们会连成一片——实测主送机关下面一度出现 5 个空行。
            # 空行不是文字内容，压缩它不违反"一个字不增不减"。
            blank_run += 1
            if blank_run > 1:
                continue
            add_blank(doc)
            continue
        blank_run = 0
        font_role, bold = ROLE_STYLES[role]
        pag = item if isinstance(item, dict) else {}
        # 段落级对齐覆盖：章标题（「第一章」「第二章」）在规章、办法里通常居中，
        # 而国标公文的一级标题是左空二字——两种情况并存，只能按段指定。
        align = resolve_align(pag.get('align'))
        if role in CENTER_ROLES:
            # 附件标题：二号小标宋居中，用标题行距，不首行缩进
            para = doc.add_paragraph()
            para.alignment = align or WD_ALIGN_PARAGRAPH.CENTER
            set_line_spacing(para, LINE_TITLE)
            run = para.add_run(text)
            set_font(run, font_by_role[font_role], SIZE_TITLE)
            set_pagination(para, bool(pag.get('keep_next')),
                           bool(pag.get('keep_lines')),
                           bool(pag.get('page_break_before')))
            continue
        add_paragraph(doc, text, font_name=font_by_role[font_role], bold=bold,
                      align=align or WD_ALIGN_PARAGRAPH.JUSTIFY,
                      keep_next=bool(pag.get('keep_next')),
                      keep_lines=bool(pag.get('keep_lines')),
                      page_break_before=bool(pag.get('page_break_before')))

    # 附件说明（7.3.4）：每个附件单独一段，回行与名称首字对齐
    attachments = [str(a).strip() for a in (content.get('attachments') or [])
                   if str(a).strip()]
    if attachments:
        add_blank(doc)
        if len(attachments) == 1:
            # 首行「附件：」左空二字起排，回行落到名称首字（再右移 3 字）
            _add_attachment_para(doc, '附件：', attachments[0],
                                 left=2 + 3, hanging=3)
        else:
            # 续行统一落在 6.5 字处（名称首字）；序号对齐在 5 字处
            cont = 2 + 3 + _NUM_PREFIX_CHARS
            for i, name in enumerate(attachments, 1):
                if i == 1:
                    _add_attachment_para(doc, f'附件：{i}. ', name,
                                         left=cont, hanging=3 + _NUM_PREFIX_CHARS)
                else:
                    _add_attachment_para(doc, f'{i}. ', name,
                                         left=cont, hanging=_NUM_PREFIX_CHARS)

    # 发文机关署名与成文日期（7.3.5）
    # body 里若已显式放了 signature / date 元素（源文档落款在正文中部，
    # 通知正文之后、附件之前），则不再在文末重复输出——把日期甩到最后一页
    # 是公文里不能犯的错。
    signature = content.get('signature')
    date = content.get('date')
    if (signature or date) and not sign_in_body:
        add_blank(doc)
        if signature:
            _add_signature_para(doc, signature, _sig_right)
        if date:
            _add_date_para(doc, date, _date_right)

    # 未在 body 中显式定位的表格，追加到文末（落款之后）
    for idx, t in enumerate(tables):
        if idx not in used_tables:
            add_blank(doc)
            add_table(doc, t.get('rows', []),
                      header=t.get('header', True),
                      align=t.get('align', 'center'),
                      spans=t.get('spans'), ncol=t.get('ncol'))

    # 版记（印发机关 + 印发日期）：整块原样搬到文末，排在末页最下方。
    # 前面不留空行——版记自带上方分隔线，再垫一个空行就把线顶离正文了；
    # 沉底靠 sink_version_record 在 save 之后实测插入段前间距来完成。
    vr = content.get('version_record') or {}
    vr_marker = (vr.get('text') or '').strip()
    if vr_marker and vr.get('blocks'):
        add_version_record(doc, vr)

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

    mode = resolve_on_exist(on_exist if on_exist is not None
                            else content.get('on_exist'))
    if mode is None:
        raise ValueError(f"on_exist 取值无法识别：{content.get('on_exist')!r}，"
                         f"可选 {ON_EXIST_MODES}")
    out, original = resolve_output_path(out, mode)
    if original:
        print(f'注意：{os.path.basename(original)} 已存在，未覆盖该文件；'
              f'本次另存为 {os.path.basename(out)}')

    out_dir = os.path.dirname(os.path.abspath(out))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    try:
        doc.save(out)
    except PermissionError:
        # 常见于产出物正被 Word/WPS 打开着，此时覆盖会失败且原因不直观
        raise PermissionError(
            f'无法写入 {os.path.abspath(out)}：'
            '文件可能正被 Word/WPS 打开，或没有写入权限。'
            '请关闭该文件或换一个输出路径后重试。')

    changed = {k: v for k, v in font_report.items() if v[0] != v[1]}
    if changed:
        print('字体回退：本机缺少国标字体，已自动替换为等效字体')
        for role, (old, new) in changed.items():
            print(f'  {old} → {new}')
        print('  （在装有国标字体的机器上生成则不会发生替换）')

    # 落款日期居中 + 版记沉底，都要渲染测量，只能在 save 之后做
    if _sig_text and _date_text and not content.get('_no_measure'):
        ok, msg = align_date_to_signature(out, _sig_text, _date_text)
        if msg:
            print(f'日期居中：{msg}')
    if vr_marker and vr.get('blocks') and not content.get('_no_measure'):
        ok, msg = sink_version_record(out, vr_marker)
        if msg:
            print(f'版记沉底：{msg}')
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
    parser.add_argument('--on-exist', dest='on_exist', choices=ON_EXIST_MODES,
                        default=None,
                        help='输出文件已存在时的处理：rename 自动另存为「原名_2」'
                             '（默认）/ fail 停止并提示 / overwrite 覆盖')
    parser.add_argument('--page-number', dest='page_number',
                        choices=PAGE_NUMBER_MODES, default=None,
                        help='页码样式：none 不显示（默认，源文档没有就不加）/ '
                             'center 居中 / odd_even 奇偶页不同')
    parser.add_argument('--fidelity', choices=FIDELITY_MODES, default=None,
                        help='保真模式：strict（默认）原文一字不增不减，'
                             '自动编号只记线索不并入正文；standard 补回自动编号')
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
        if args.fidelity:
            content['fidelity'] = args.fidelity
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
        path = build_document(content, out, on_exist=args.on_exist)
    except ValueError as e:
        # 直接抛 traceback 对用户没有意义，给出一句能照着改的提示
        print(f'错误：{e}', file=sys.stderr)
        return 1
    except FileExistsError as e:
        target = getattr(e, 'filename', None) or out
        print(f'已停止：输出文件已存在，为避免覆盖未生成新文件。', file=sys.stderr)
        print(f'  {os.path.abspath(target)}', file=sys.stderr)
        print('  处理方式：另指定 --output；或加 --on-exist rename 自动另存为'
              '「原名_2」；或加 --on-exist overwrite 明确覆盖。', file=sys.stderr)
        return 2
    except PermissionError as e:
        print(f'错误：{e}', file=sys.stderr)
        return 3
    print(f'公文已生成: {os.path.abspath(path)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())

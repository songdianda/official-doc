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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

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

# 层级识别：编号前缀 -> (字体, 是否加粗)
LEVEL_PATTERNS = [
    (re.compile(r'^[一二三四五六七八九十]+、'), FONT_HEI, False),      # 一、
    (re.compile(r'^（[一二三四五六七八九十]+）'), FONT_KAI, False),    # （一）
    (re.compile(r'^\d+\.'), FONT_BODY, True),                        # 1.
    (re.compile(r'^（\d+）'), FONT_BODY, True),                      # （1）
    (re.compile(r'^[①-⑳]'), FONT_BODY, True),                        # ①
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

    # LEVEL_PATTERNS 在模块加载时已固化旧字体名，原地重建
    LEVEL_PATTERNS[:] = [
        (re.compile(r'^[一二三四五六七八九十]+、'), FONT_HEI, False),      # 一、
        (re.compile(r'^（[一二三四五六七八九十]+）'), FONT_KAI, False),    # （一）
        (re.compile(r'^\d+\.'), FONT_BODY, True),                        # 1.
        (re.compile(r'^（\d+）'), FONT_BODY, True),                      # （1）
        (re.compile(r'^[①-⑳]'), FONT_BODY, True),                        # ①
    ]
    return {r: (before[r], resolved[r]) for r in resolved}


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


def detect_level(text):
    """按编号前缀识别层级，返回 (字体, 是否加粗)。无匹配返回 None。"""
    stripped = text.strip()
    for pattern, font, bold in LEVEL_PATTERNS:
        if pattern.match(stripped):
            return font, bold
    return None


# ── 页码（奇偶页不同） ────────────────────────────────────

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


def add_page_number(section):
    """添加奇偶页不同的页码：单页码居右空一字，双页码居左空一字（7.5）。"""
    # 1) 开启奇偶页不同
    settings = section.part.document.settings.element
    if settings.find(qn('w:evenAndOddHeaders')) is None:
        flag = OxmlElement('w:evenAndOddHeaders')
        settings.insert(0, flag)

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
      body        正文列表，元素可为字符串（自动识别层级）或
                  {"type": "text"|"level1".."level5"|"blank", "text": "..."}
      attachments 附件名称列表（可选）
      signature   发文机关署名（可选）
      date        成文日期（可选）
      output      输出路径（可选，优先级低于 output_path 参数）
    """
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

    # 标题：小标宋 二号 居中 行距32磅
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(title_para, LINE_TITLE)
    title_run = title_para.add_run(content['title'])
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
    for item in content.get('body', []):
        if isinstance(item, str):
            text = item
            level = detect_level(text)
            font, bold = level if level else (FONT_BODY, False)
        else:
            text = item.get('text', '')
            ltype = item.get('type', 'text')
            if ltype == 'blank':
                add_blank(doc)
                continue
            font, bold = {
                'level1': (FONT_HEI, False),
                'level2': (FONT_KAI, False),
                'level3': (FONT_BODY, True),
                'level4': (FONT_BODY, True),
                'level5': (FONT_BODY, True),
            }.get(ltype, (FONT_BODY, False))
        if not text.strip():
            add_blank(doc)
            continue
        add_paragraph(doc, text, font_name=font, bold=bold)

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

    # 页码
    add_page_number(section)

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
        }

    out = args.output or content.get('output')
    if not out:
        parser.error('未指定输出路径，请用 --output 指定')

    path = build_document(content, out)
    print(f'公文已生成: {os.path.abspath(path)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从现有 .docx 提取段落内容 + 角色判定线索，供人工/AI 审阅后交给生成器。

用法：
    python3 inspect_docx.py 原文.docx -o content.json
    python3 inspect_docx.py 原文.docx            # 打印到标准输出

设计意图：
    段落角色（标题 vs 编号正文）是语义问题，不该由排版脚本猜。
    本脚本只做「提取 + 给建议 + 列存疑」，最终 type 由人工/AI 拍板。
    生成的 JSON 可直接喂给 generate_gongwen.py，未确认的段按最安全
    的正文处理。

依赖：pip install python-docx
"""

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

# 编号前缀 -> 可能的层级
NUM_PREFIX = [
    (re.compile(r'^[一二三四五六七八九十]+、'), 1),
    (re.compile(r'^（[一二三四五六七八九十]+）'), 2),
    (re.compile(r'^\d+\.'), 3),
    (re.compile(r'^（\d+）'), 4),
    (re.compile(r'^[①-⑳]'), 5),
]

# 样式名 -> 层级（中英文 Word 常见写法）
STYLE_HEADING = re.compile(r'(?:heading|标题|heading)\s*(\d)', re.I)
TERMINAL_PUNCT = '。；！？'

# 成文日期（7.3.5.4：阿拉伯数字标全年月日）。整段仅为日期时几乎必然是落款日期。
DATE_FULL = re.compile(r'^\d{4}年\d{1,2}月\d{1,2}日$')
# 剥离依据：写出来，才能让使用者判断这次剥离可不可信
STRIP_REASON = {
    'date': '整段仅为「YYYY年M月D日」，符合成文日期格式',
    'signature': '位于成文日期之上最近的非空短段'
                 '（≤40 字且不以句末标点或顿号结尾）',
}
# 附件说明（7.3.4）。
ATTACH_PREFIX = re.compile(r'^附件[:：]')
# 发文字号（7.2.4）：机关代字 + 〔年份〕 + 序号，如「X教发〔2026〕15号」
DOC_NUMBER = re.compile(r'^[\u4e00-\u9fa5]{1,15}[〔\[（]\s*\d{4}\s*[〕\]）]\s*第?\s*\d+\s*号?$')

# 红头常见色值：R 高、G/B 低（FF0000 / C00000 / E60000 等）
def is_red_color(val):
    if not val or val.lower() in ('auto', '000000'):
        return False
    try:
        r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
    except (ValueError, IndexError):
        return False
    return r >= 0xA0 and g <= 0x60 and b <= 0x60


CN_DIGITS = '零一二三四五六七八九十'


def cn_number(n):
    """阿拉伯数字 → 中文数字（1~99），超出返回原数字字符串。"""
    if n <= 10:
        return CN_DIGITS[n]
    if n < 20:
        return '十' + CN_DIGITS[n - 10]
    if n < 100:
        return CN_DIGITS[n // 10] + '十' + (CN_DIGITS[n % 10] if n % 10 else '')
    if n == 100:
        return '一百'
    return str(n)


def roman(n):
    table = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'), (100, 'C'),
             (90, 'XC'), (50, 'L'), (40, 'XL'), (10, 'X'), (9, 'IX'),
             (5, 'V'), (4, 'IV'), (1, 'I')]
    out = ''
    for v, s in table:
        while n >= v:
            out += s
            n -= v
    return out or str(n)


def format_number(value, fmt):
    """按 numFmt 渲染序号。bullet / none 返回 None（由调用方取 lvlText 原样）。"""
    f = (fmt or '').lower()
    if f in ('bullet', 'none'):
        return None
    if f in ('japanesecounting', 'chinesecounting', 'chinesecountingthousand',
             'chineselegalsimplified', 'ideographdigital', 'ideographtraditional'):
        return cn_number(value)
    if f == 'lowerletter':
        return chr(ord('a') + value - 1) if 1 <= value <= 26 else str(value)
    if f == 'upperletter':
        return chr(ord('A') + value - 1) if 1 <= value <= 26 else str(value)
    if f == 'lowerroman':
        return roman(value).lower()
    if f == 'upperroman':
        return roman(value)
    return str(value)                      # decimal 等默认阿拉伯数字


def load_numbering_defs(doc):
    """解析 numbering.xml → {numId: {ilvl: {'fmt','lvlText','start'}}}。

    返回空字典表示文档没有自动编号定义。
    """
    try:
        part = doc.part.numbering_part
    except Exception:
        return {}
    if part is None:
        return {}
    root = part.element

    abstracts = {}
    for a in root.iter(qn('w:abstractNum')):
        aid = a.get(qn('w:abstractNumId'))
        levels = {}
        for lvl in a.findall(qn('w:lvl')):
            fmt_el = lvl.find(qn('w:numFmt'))
            txt_el = lvl.find(qn('w:lvlText'))
            st_el = lvl.find(qn('w:start'))
            levels[lvl.get(qn('w:ilvl'))] = {
                'fmt': fmt_el.get(qn('w:val')) if fmt_el is not None else 'decimal',
                'lvlText': txt_el.get(qn('w:val')) if txt_el is not None else '%1.',
                'start': int(st_el.get(qn('w:val'))) if st_el is not None else 1,
            }
        abstracts[aid] = levels

    nums = {}
    for n in root.iter(qn('w:num')):
        aid_el = n.find(qn('w:abstractNumId'))
        levels = dict(abstracts.get(aid_el.get(qn('w:val')), {})
                      if aid_el is not None else {})
        for ov in n.findall(qn('w:lvlOverride')):
            st = ov.find(qn('w:startOverride'))
            if st is not None:
                ilvl = ov.get(qn('w:ilvl'))
                lv = dict(levels.get(ilvl, {}))
                lv['start'] = int(st.get(qn('w:val')))
                levels[ilvl] = lv
        nums[n.get(qn('w:numId'))] = levels
    return nums


def load_style_numbering(doc):
    """解析 styles.xml → {styleId: numPr元素}（沿 basedOn 继承链向上找）。

    编号规则可能挂在**段落样式**上而不是段落自身（Word 的"多级列表"常见做法），
    只看段落的 w:pPr/w:numPr 会整条丢失编号。
    """
    part = None
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        part = doc.part.part_related_by(RT.STYLES)
    except Exception:
        part = None
    if part is None:
        # 不同版本 python-docx 暴露 styles part 的方式不一致，兜底按名字找
        try:
            for p in doc.part.package.iter_parts():
                if str(p.partname).endswith('/styles.xml'):
                    part = p
                    break
        except Exception:
            part = None
    if part is None:
        return {}
    raw = {}
    for s in part.element.iter(qn('w:style')):
        sid = s.get(qn('w:styleId'))
        if not sid:
            continue
        pPr = s.find(qn('w:pPr'))
        numPr = pPr.find(qn('w:numPr')) if pPr is not None else None
        base_el = pPr.find(qn('w:basedOn')) if pPr is not None else None
        raw[sid] = {'numPr': numPr,
                    'basedOn': base_el.get(qn('w:val')) if base_el is not None else None}

    out = {}
    for sid in raw:
        seen, cur, found = set(), sid, None
        while cur and cur in raw and cur not in seen:
            seen.add(cur)
            if raw[cur]['numPr'] is not None:
                found = raw[cur]['numPr']
                break
            cur = raw[cur]['basedOn']
        if found is not None:
            out[sid] = found
    return out


LEVEL_PATTERNS = [
    ('一、', re.compile(r'^([一二三四五六七八九十]+)、')),
    ('（一）', re.compile(r'^（([一二三四五六七八九十]+)）')),
    # 后接 (?!\d) 排除多级编号：否则「1.」与「1.1.」会被当成同一个 1 而误报重复
    ('1.', re.compile(r'^(\d+)[.、](?!\d)')),
    ('（1）', re.compile(r'^（(\d+)）')),
    ('①', re.compile(r'^([①-⑳])')),
]


def detect_number_conflicts(paras):
    """找出重复出现的同级编号。

    手打编号与 Word 自动编号混用时最容易撞号：自动编号按自己的计数走，
    手打的那条它不知道，于是文档里出现两个「一、」。
    **只报告、不自动改**——重复既可能是原文就错，也可能是两套编号各算一套，
    擅自改会乱层级，交给使用者判断。
    """
    seen = {}
    conflicts = []
    for sig in paras:
        text = (sig.get('text') or '').strip()
        if not text:
            continue
        # strict 下自动编号没有写进正文，措辞要说清，别让人以为已补回
        source = ('自动编号（未并入正文）' if sig.get('autoNumber') else '原文')
        for label, pat in LEVEL_PATTERNS:
            m = pat.match(text)
            if not m:
                continue
            key = (label, m.group(1))
            if key in seen:
                prev_idx, prev_src = seen[key]
                conflicts.append({
                    'level': label, 'value': m.group(1),
                    'first': {'index': prev_idx, 'source': prev_src},
                    'again': {'index': sig['index'], 'source': source},
                    'text': text[:40],
                })
            else:
                seen[key] = (sig['index'], source)
            break
    return conflicts


def resolve_auto_number(p, defs, counters, last_values, style_num=None):
    """算出 Word 自动编号在该段应该显示的文本（如「二、」）。

    python-docx 的 p.text 只取 w:t，自动编号不在文本里而挂在 w:numPr 上，
    直接提取会**整条丢失**——源文件里手打的「一、」在，Word 自动生成的
    「二、」就没了。

    返回 (显示文本, 线索dict)；无自动编号返回 (None, None)。
    """
    if not defs:
        return None, None
    pPr = p._element.pPr
    numPr = pPr.find(qn('w:numPr')) if pPr is not None else None
    if numPr is None and style_num:
        # 编号规则挂在段落样式上时（多级列表常见），段落自身没有 numPr
        sid = p.style.style_id if p.style is not None else None
        numPr = style_num.get(sid) if sid else None
    if numPr is None:
        return None, None
    ilvl_el = numPr.find(qn('w:ilvl'))
    numid_el = numPr.find(qn('w:numId'))
    if ilvl_el is None or numid_el is None:
        return None, None
    ilvl = ilvl_el.get(qn('w:val'))
    numid = numid_el.get(qn('w:val'))

    d = defs.get(numid, {}).get(ilvl)
    if not d:
        return None, None

    key = (numid, ilvl)
    if key in counters:
        counters[key] += 1
    else:
        counters[key] = d.get('start', 1)
    value = counters[key]
    last_values[key] = value

    # 本级计数推进后，所有**更深**的层级回到各自起始值。
    # Word 的行为是进入新的一级标题后，其下的二级标题从 1 重新数；
    # 不重置的话第二章的二级会接着第一章数成「2.3.」而不是「2.1.」。
    try:
        cur = int(ilvl)
    except (TypeError, ValueError):
        cur = 0
    for k in list(counters):
        if k[0] != numid:
            continue
        try:
            if int(k[1]) > cur:
                counters[k] = defs.get(numid, {}).get(k[1], {}).get('start', 1) - 1
                last_values.pop(k, None)
        except (TypeError, ValueError):
            continue

    rendered = format_number(value, d.get('fmt'))
    lvl_text = d.get('lvlText', '%1.')

    if rendered is None:                   # 项目符号：原样保留，不编号
        return lvl_text, {'numId': numid, 'ilvl': ilvl, 'fmt': d.get('fmt'),
                          'value': value, 'text': lvl_text, 'bullet': True}

    def sub(m):
        idx = str(int(m.group(1)) - 1)
        if idx == ilvl:
            return rendered
        if (numid, idx) in last_values:    # 多级编号的父层级
            return str(last_values[(numid, idx)])
        return str(defs.get(numid, {}).get(idx, {}).get('start', 1))

    text = re.sub(r'%(\d+)', sub, lvl_text)
    return text, {'numId': numid, 'ilvl': ilvl, 'fmt': d.get('fmt'),
                  'value': value, 'text': text, 'bullet': False}


FIDELITY_MODES = ('strict', 'standard')
FIDELITY_DEFAULT = 'strict'
FIDELITY_ALIASES = {
    'strict': 'strict', '严格': 'strict', '保真': 'strict', 'strictfidelity': 'strict',
    'standard': 'standard', '标准': 'standard', '补编号': 'standard',
    'fill': 'standard', 'restore': 'standard',
}


def resolve_fidelity(value):
    """保真模式归一：strict（默认）/ standard；无法识别返回 None。

    strict   原文一个字不增不减。Word 自动编号在源文件里不是字符，
              渲染出来补进正文就是"加字"，所以只作线索记录，不并入 text。
    standard 把自动编号补回成文字（旧行为），适用于本来就要把
              自动编号固化成字符的场景。
    """
    if value is None:
        return FIDELITY_DEFAULT
    key = str(value).strip().lower()
    key = FIDELITY_ALIASES.get(key, key)
    return key if key in FIDELITY_MODES else None


# w:jc 的合法取值不止 python-docx 枚举认识的那几个。LibreOffice 转出的
# docx 会写 start / end（OOXML 严格版的写法），直接读 p.alignment 会抛
# ValueError。取值只作线索用，统一归到 familiar 的几个名字上即可。
JC_ALIASES = {'start': 'left', 'end': 'right', 'both': 'justify'}


def safe_alignment(p):
    """读段落对齐，遇到不认识的 w:jc 值也不崩。"""
    pPr = p._element.pPr
    if pPr is None:
        return None
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        return None
    val = jc.get(qn('w:val'))
    if not val:
        return None
    return JC_ALIASES.get(val, val)


def para_signals(p, index, auto_prefix=None, auto_info=None):
    """采集一段的样式线索。auto_prefix 为 Word 自动编号渲染出的文本。

    text      = 原文字符 + 自动编号（判读用，能看出「二、」这类层级）
    raw_text  = 原文字符本身（严格保真输出用）
    两者分开存：判读需要编号才能定层级，输出却不能凭空多出这几个字。
    """
    r = p.runs[0] if p.runs else None
    rPr = r._element.rPr if (r is not None) else None
    rFonts = rPr.find(qn('w:rFonts')) if rPr is not None else None
    pPr = p._element.pPr
    outline = None
    if pPr is not None:
        e = pPr.find(qn('w:outlineLvl'))
        if e is not None:
            try:
                outline = int(e.get(qn('w:val')))
            except (TypeError, ValueError):
                outline = None
    numPr = pPr is not None and pPr.find(qn('w:numPr')) is not None
    raw_text = p.text.strip()
    text = (auto_prefix + raw_text) if auto_prefix else raw_text
    return {
        'index': index,
        'text': text,
        'raw_text': raw_text,
        'auto_prefix': auto_prefix,
        'style': p.style.name if p.style is not None else None,
        'styleId': p.style.style_id if p.style is not None else None,
        'outlineLvl': outline,
        'bold': bool(r.font.bold) if (r is not None and r.font.bold is not None) else False,
        'size': (r.font.size.pt if (r is not None and r.font.size) else None),
        'font': (rFonts.get(qn('w:eastAsia')) if rFonts is not None else None),
        'listPara': numPr,          # Word 自动编号/项目符号
        'autoNumber': auto_info,    # 自动编号线索（由 numbering.xml 算出，非原文字符）
        'alignment': safe_alignment(p),
    }


def suggest(sig):
    """给出建议角色 + 置信度 + 理由。只是建议，最终由人工/AI 决定。

    不使用字数阈值；主要看语法特征（是否成句）和 Word 语义标记
    （outlineLvl / 标题样式），加粗仅作弱参考——原格式本身可能有错。
    """
    text = sig['text']
    if not text:
        return 'blank', 'high', '空段落'

    reasons = []
    # 1) Word 语义标记最强：标题样式 / outlineLvl
    level = None
    if sig['outlineLvl'] is not None and sig['outlineLvl'] >= 0:
        level = sig['outlineLvl'] + 1
        reasons.append(f"Word outlineLvl={sig['outlineLvl']}")
    elif sig['styleId']:
        m = STYLE_HEADING.search(sig['styleId'])
        if m:
            level = int(m.group(1))
            reasons.append(f"标题样式「{sig['style']}」")

    prefix_level = next((lv for pat, lv in NUM_PREFIX if pat.match(text)), None)

    if level:
        role = f'heading{min(level, 5)}'
        return role, 'high', '；'.join(reasons)

    # 2) 无语义标记，看编号 + 句法。
    # 句末标点**只是线索，不是判据**：带问号的标题、省略句号的工作事项都真实存在。
    # 因此带标点时不再给 high，一律要求结合样式与上下文复核。
    if prefix_level:
        last = text.rstrip()[-1] if text.rstrip() else ''
        if last in '。；':
            return ('numbered_body', 'medium',
                    '以句号/分号结尾，倾向完整叙述句（罗列的工作事项）；'
                    '但标点只是线索，若实为标题请改判')
        if last in '？！':
            # 疑问句标题、号召式小标题同样带问号叹号，无法据此判断
            return (f'heading{prefix_level}', 'low',
                    '以问号/叹号结尾——疑问句标题与感叹式事项都可能，'
                    '标点不足以判定，需结合上下文确认')
        if prefix_level <= 2:
            # 汉字编号（一、/（一））通常不会是正文罗列，但仍有例外
            return (f'heading{prefix_level}', 'medium',
                    '汉字编号且不以句末标点结尾，通常是标题；'
                    '未发现 Word 标题样式，属线索判定')
        # 阿拉伯数字编号（1. /（1）/ ①）：国标标题体系到「（一）」为止，更深
        # 一级的阿拉伯编号在公文里绝大多数是正文罗列项，不是标题。即便无句号
        # （如「1.做好xx工作」）也是省略标点的短事项，不该按标题加粗。
        return ('numbered_body', 'medium',
                '阿拉伯数字编号（1./（1）/①）在公文中通常是正文罗列项，'
                '不是标题；若实为标题请改判')
    return 'body', 'high', '无编号，普通正文'


def detect_page_number(doc):
    """检测源文档是否已有页码。

    返回 {'found': bool, 'mode': 'center'|'odd_even'|'none', 'detail': [...]}

    规则（重要）：**源文档没页码就不给加**。排版是重排格式，不是替用户
    决定要不要页码——单页文件、内部传阅件常刻意不编页码，凭空加出来是
    越权。源文档有页码时，沿用其原有样式（居中 / 奇偶页不同）。
    """
    detail = []
    found = False
    has_even = False

    settings = doc.settings.element
    odd_even_on = settings.find(qn('w:evenAndOddHeaders')) is not None

    for si, section in enumerate(doc.sections):
        for tag in ('w:footerReference', 'w:headerReference'):
            for ref in section._sectPr.findall(qn(tag)):
                rid = ref.get(qn('r:id'))
                if rid is None or rid not in section.part.related_parts:
                    continue
                part = section.part.related_parts[rid]
                instr = ' '.join(e.text or ''
                                 for e in part.element.iter(qn('w:instrText')))
                # 用 token 匹配，避免 NUMPAGES / SECTIONPAGES 被当成 PAGE
                tokens = re.findall(r'[A-Za-z]+', instr.upper())
                text = ''.join(t.text or ''
                               for t in part.element.iter(qn('w:t')))
                is_page = 'PAGE' in tokens
                # 兜底：页脚里写死的纯数字（无域代码）也算页码
                if not is_page and text.strip('—-－  \t'):
                    is_page = bool(re.fullmatch(r'[0-9]{1,4}', text.strip('—-－ ')))
                if not is_page:
                    continue
                found = True
                ftype = ref.get(qn('w:type'))
                if ftype == 'even':
                    has_even = True
                detail.append({
                    'section': si,
                    'location': 'footer' if tag.startswith('w:footer') else 'header',
                    'type': ftype,
                    'field': instr.strip() or '(无域，文本页码)',
                    'text': text,
                })

    if not found:
        return {'found': False, 'mode': 'none', 'detail': []}
    mode = 'odd_even' if (has_even and odd_even_on) else 'center'
    return {'found': True, 'mode': mode, 'detail': detail}


def extract_table(tbl):
    """提取表格并保留合并单元格结构。

    python-docx 的 row.cells 会把横向合并的格子按物理列重复展开——一个跨 2 列的
    「300」会被读成 ['300', '300']，照搬进新表就凭空多出一个不存在的数值。
    这里解析 w:gridSpan / w:vMerge：被合并覆盖的位置一律留空，跨度记在 spans 里，
    由生成器真正执行单元格合并。

    返回 (rows, spans, ncol)。spans[i][j] 为 None 或
    {'h': 横向跨列数} / {'v': 'restart'|'continue'}。
    """
    rows, spans = [], []
    in_merge = {}       # 逻辑列 -> 上一行该列是否处于纵向合并中
    for tr in tbl._tbl.findall(qn('w:tr')):
        row, srow = [], []
        ci, recorded = 0, []
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            colspan, vmerge = 1, None
            if tcPr is not None:
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    try:
                        colspan = max(int(gs.get(qn('w:val')) or 1), 1)
                    except (TypeError, ValueError):
                        colspan = 1
                vm = tcPr.find(qn('w:vMerge'))
                if vm is not None:
                    val = vm.get(qn('w:val'))
                    # ST_Merge 的缺省值是 continue：Word 常给后续格写 <w:vMerge/>
                    # 而不带 val，按"非 continue 即 restart"判会把续行误判成新起点。
                    # 仅当上一行同列并未处于合并中时，孤立的省略式才按 restart 处理。
                    if val == 'restart':
                        vmerge = 'restart'
                    elif val == 'continue':
                        vmerge = 'continue'
                    else:
                        vmerge = 'continue' if in_merge.get(ci) else 'restart'
            text = ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip()
            if vmerge == 'continue':
                row.append('')            # 内容已在上一行同列，此处留空
                srow.append({'v': 'continue'})
            else:
                row.append(text)
                span = {}
                if colspan > 1:
                    span['h'] = colspan
                if vmerge == 'restart':
                    span['v'] = 'restart'
                srow.append(span or None)
            recorded.append((ci, vmerge))
            ci += colspan
        rows.append(row)
        spans.append(srow)
        in_merge = {c: True for c, v in recorded if v is not None}

    def width(srow):
        return sum((s or {}).get('h', 1) for s in srow)

    ncol = max((width(srow) for srow in spans), default=0)
    return rows, spans, ncol


def _para_max_size(el):
    """段落内最大字号（磅）。w:sz 存的是半磅值。"""
    sizes = []
    for sz in el.iter(qn('w:sz')):
        try:
            sizes.append(int(sz.get(qn('w:val'))) / 2.0)
        except (TypeError, ValueError):
            pass
    return max(sizes) if sizes else 0


def _para_is_red(el):
    return any(is_red_color(c.get(qn('w:val'))) for c in el.iter(qn('w:color')))


def _para_pagination(el):
    """段落的分页控制属性（Word「换行和分页」里的两个开关）。

    w:keepNext         = 与下段同页 —— 标题不孤零零留在页末，这是最常见的设置
    w:keepLines        = 段中不分页 —— 一段文字不跨页断开
    w:pageBreakBefore  = 段前分页 —— 这一段另起一页

    源文档设了而成品不写，分页就会变。这不是字体字号那类"不够好看"的偏差，
    是标题掉在页底、正文被切开的排版事故，必须跟内容一样带走。

    注意：手动分页符（run 里的 w:br type="page"）不在此列，本技能不保留。
    """
    pPr = el.find(qn('w:pPr'))
    if pPr is None:
        return {}
    out = {}
    for tag, key in (('w:keepNext', 'keep_next'),
                     ('w:keepLines', 'keep_lines'),
                     ('w:pageBreakBefore', 'page_break_before')):
        node = pPr.find(qn(tag))
        if node is not None and (node.get(qn('w:val')) or 'true').lower() != 'false':
            out[key] = True
    return out


def _para_has_border(el):
    pPr = el.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:pBdr')) is not None


def _para_has_image(el):
    """段落里是否有图片。很多单位的红头是图片而不是文字，必须一并保留。"""
    return (any(True for _ in el.iter(qn('w:drawing')))
            or any(True for _ in el.iter(qn('w:pict'))))


def _textbox_text(el):
    """取文本框（w:txbxContent）里的文字。

    文本框内的 w:t 不在段落的直接 w:r 下，python-docx 的 p.text 读不到——
    文本框版红头因此会显示成"空段落"，仅靠图片判定捞它并不可靠。
    """
    parts = []
    for box in el.iter(qn('w:txbxContent')):
        for t in box.iter(qn('w:t')):
            parts.append(t.text or '')
    return ''.join(parts).strip()


# 数学公式命名空间（m:），python-docx 的 nsmap 里没有，需手写
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

# VML（老式图形 w:pict 里的 v:line / v:shape / v:imagedata），同样不在 nsmap。
# 直接写 qn('v:line') 会抛 KeyError: 'v'。
VML_NS = 'urn:schemas-microsoft-com:vml'


def _m(tag):
    return '{%s}%s' % (M_NS, tag)


def _vml(tag):
    return '{%s}%s' % (VML_NS, tag)


def _qname(tag):
    """把 'w:xxx' / 'v:xxx' / 'oMathxxx' 形式的标签名转成 Clark 记法。"""
    if tag.startswith('oMath'):
        return _m(tag)
    if tag.startswith('v:'):
        return _vml(tag[2:])
    return qn(tag)


def _has(el, *tags):
    return any(True for _ in el.iter(tags[0])) if len(tags) == 1 else \
        any(any(True for _ in el.iter(t)) for t in tags)


# 正文里本技能**没有完整验证过**的内容。检测到必须说出来——
# 静默漏掉等于交付一份缺内容的文档，使用者还以为排好了。
UNSUPPORTED = [
    # 只认真正引用了位图的元素。红头红线、版记线是 v:line / v:shape，
    # 没有位图文件，混进来会把「一条分隔线」报成「一张图片」。
    ('图片', ('a:blip', 'v:imagedata'),
     '正文图片不参与重排，尺寸与位置可能变化，需生成后核对'),
    ('图形', ('v:line', 'v:shape', 'v:rect', 'v:oval', 'v:roundrect'),
     '线条、形状等矢量图形不保留（红头分隔线、版记线多属此类）'),
    ('文本框', ('w:txbxContent',),
     '文本框内的文字会当普通段落排，框体本身不保留'),
    ('脚注', ('w:footnoteReference',),
     '脚注内容不在正文段落里，不会被提取'),
    ('尾注', ('w:endnoteReference',),
     '尾注内容不在正文段落里，不会被提取'),
    ('批注', ('w:commentRangeStart', 'w:commentReference'),
     '批注不会被提取'),
    ('修订标记', ('w:ins', 'w:del', 'w:delText'),
     '按最终显示文本提取；若含未接受的修订，结果可能与预期不符'),
    ('嵌入对象', ('w:object',),
     '嵌入对象（图表 / OLE）不会被提取'),
    ('公式', ('oMath', 'oMathPara'),
     '公式不会被提取'),
]


def scan_unsupported(elements, skip):
    """扫描正文里未完整验证的内容，只报告、不擅自处理。

    skip 为版头段落下标——它们已整段原样保留，其中的图片 / 文本框不算问题。
    """
    found = {}
    for i, el in enumerate(elements):
        if i in skip:
            continue
        for name, tags, hint in UNSUPPORTED:
            probe = [_qname(t) for t in tags]
            if _has(el, *probe):
                rec = found.setdefault(
                    name, {'count': 0, 'index': [], 'hint': hint})
                rec['count'] += 1
                rec['index'].append(i)
    return [{'kind': k, 'count': v['count'], 'index': v['index'],
             'hint': v['hint']} for k, v in found.items()]


def detect_letterhead(paras, elements):
    """识别红头（发文机关标志）、发文字号、红色分隔线。

    只扫描文档开头那几段：正文里引述「X教发〔2026〕15号」不该被当成版头删掉。
    打分判定，遇到第一个不像版头的段落即停止，保证取到的是开头整块而非散落段落。
    识别到发文字号之后，只再接受分隔线——否则正文标题可能被误纳入。

    红头与发文字号之间常有空行（实测样本：红头在第 1 段、文号在第 4 段，
    中间隔两个空行）。早期版本遇到空行就停，结果文号检不出来，反而因为
    「首个非空非版头段落」的规则被当成了标题——红头文件最不能错的一处。
    现在允许跨过有限个空行继续找文号，拿到文号后才收紧。

    返回 (索引列表, 明细列表)。
    """
    picked, detail = [], []
    got_number = False
    blank_run = 0               # 已识别版头后连续遇到的空行数
    MAX_BLANK_RUN = 3           # 空行太多就说明版头到此为止了
    for i in range(min(len(paras), 10)):
        text = paras[i]['text'].strip()
        # 文本框里的字不在段落的直接 w:r 下，p.text 读不到，单独取出来参与判定
        box = _textbox_text(elements[i]) if not text else ''
        probe = text or box
        if not probe:
            if _para_has_border(elements[i]) or _para_has_image(elements[i]):
                picked.append(i)
                has_img = _para_has_image(elements[i])
                detail.append({'index': i,
                               'text': '(红头图片)' if has_img else '(分隔线)',
                               'reason': ('无文字且含图形（红头常为图片或分隔线）'
                                          if has_img else '无文字且带段落边框')})
                blank_run = 0
                continue
            if not picked:
                continue          # 版头之前的空行，跳过
            # 已拿到红头：文号还没着落时跨过空行继续找，找到了就只等分隔线
            if not got_number:
                blank_run += 1
                if blank_run > MAX_BLANK_RUN:
                    break
                continue
            continue              # 文号之后的空行，属于分隔线区，继续但不计数
        blank_run = 0
        if got_number:
            break                 # 文号之后只剩分隔线，有文字就该是标题了

        score, reasons = 0, []
        if _para_is_red(elements[i]):
            score += 2
            reasons.append('红色文字')
        if _para_max_size(elements[i]) >= 22:
            score += 1
            reasons.append('字号≥二号')
        if DOC_NUMBER.match(probe):
            score += 3
            reasons.append('发文字号格式')
            got_number = True
        if probe.endswith('文件') and i <= 2:
            score += 2
            reasons.append('以"文件"结尾')
        if box and i <= 2:
            score += 1            # 文本框本身是做版头的常见做法
            reasons.append('文本框内文字')
        if score >= 2:
            picked.append(i)
            detail.append({'index': i, 'text': probe,
                           'reason': '、'.join(reasons)})
            continue
        break
    return picked, detail


def detect_version_record(elements, body):
    """识别文末版记（印发机关 + 印发日期），整块原样带走、不重排。

    GB/T 9704-2012 7.4：版记排在公文最后一面版心最下方，印发机关左空一字、
    印发日期右空一字，上下各一条与版心等宽的分隔线。

    做法是**不重排**：把源文档那两段（分隔线段 + 印发行）的 XML 整块搬过去。
    版记的字号、机关与日期之间的间距、两条线都是单位自己排好的，重排一次
    就变一次样——这正是"不要给变动"的意思。位置由生成器保证：排在文末，
    并做沉底校准（生成器里做，见 generate_gongwen.sink_version_record）。

    只在 body 末尾找：正文中间引述「……印发」的文件名不该被当成版记搬走。

    返回 (版记 dict 或 None, 余下的 body)。
    """
    pos = None
    for i in range(len(body) - 1, -1, -1):
        b = body[i]
        if not isinstance(b, dict) or b.get('type') in ('signature', 'date',
                                                       'table'):
            continue
        t = (b.get('text') or '').strip()
        if not t:
            continue
        if len(t) <= 60 and t.endswith('印发'):
            pos = i
        break
    if pos is None:
        return None, body

    src_idx = ((body[pos].get('_src') or {}).get('index'))
    if src_idx is None or src_idx >= len(elements):
        return None, body

    idxs = [src_idx]
    # 上分隔线：紧邻其前的空段，且带 VML 线条 / 形状
    j = src_idx - 1
    if j >= 0:
        el = elements[j]
        txt = ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()
        if not txt and any(_has(el, _vml(k))
                           for k in ('line', 'shape', 'rect')):
            idxs.insert(0, j)

    blocks = []
    for i in idxs:
        el = elements[i]
        # 带图片引用的段落搬 XML 会丢 rId，退回当正文处理并让验收去报
        if el.find('.//' + qn('a:blip')) is not None \
                or el.find('.//' + _vml('imagedata')) is not None:
            return None, body
        blocks.append(etree.tostring(el, encoding='unicode'))

    # body 里连带它前面的空段一起移除：版记自带上方分隔线段，再留一个空行
    # 就把分隔线顶离正文了
    rest = body[:pos]
    k = len(rest) - 1
    while k >= 0 and isinstance(rest[k], dict) \
            and not (rest[k].get('text') or '').strip():
        k -= 1
    rest = rest[:k + 1]
    rest += body[pos + 1:]
    return ({'text': body[pos]['text'], 'blocks': blocks,
             'source_index': src_idx}, rest)


def build(path, strip_sign=True, fidelity=FIDELITY_DEFAULT):
    """提取内容与角色建议（不修改原文）。

    strip_sign=False 时不剥离署名 / 成文日期，两段都留在 body。
    自动剥离靠的是格式特征，在名单式、表格化落款等特殊文档里可能判错，
    此时应由使用者自己决定怎么处理。

    fidelity='strict'（默认）：原文一个字不增不减。
    源文件里 Word 自动编号不是字符而是 w:numPr，把它渲染成「二、」写进
    正文就是凭空加字——strict 下只记录线索，原样输出段落自身的字符。
    """
    fidelity = resolve_fidelity(fidelity)
    if fidelity is None:
        raise ValueError(f'fidelity 取值无法识别：{fidelity!r}，'
                         f'可选 {FIDELITY_MODES}')
    strict = fidelity == 'strict'
    doc = Document(path)

    # 按文档流顺序遍历，否则表格相对段落的位置会丢失
    from docx.oxml.ns import qn as _qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    flow = []
    for child in doc.element.body.iterchildren():
        if child.tag == _qn('w:p'):
            flow.append(('p', Paragraph(child, doc)))
        elif child.tag == _qn('w:tbl'):
            flow.append(('t', Table(child, doc)))

    paras = []
    elements = []
    tables = []
    auto_numbered = []
    # 占位：body 中记录 ('para', sig) / ('table', index)
    order = []

    # Word 自动编号（w:numPr）不在 w:t 里，必须单独解析，否则整条编号丢失
    num_defs = load_numbering_defs(doc)
    # 编号规则可能挂在段落样式上（多级列表常见），段落自身没有 numPr
    style_num = load_style_numbering(doc)
    counters, last_values = {}, {}

    # 文档以表格开头时，那个表格很可能是表格式版头（红头 + 文号）。
    # 本技能不会把它当版头原样保留，必须说出来，不能默默按普通表格排掉。
    warnings = []
    if flow and flow[0][0] == 't':
        warnings.append(
            '文档以表格开头：若该表格是表格式版头（红头 + 发文字号），'
            '本技能按普通表格处理，它会被排在标题之后。'
            '如需原样保留为版头，请先在源文档中改为文字 / 图片 / 文本框版红头，'
            '或在生成后手工调整。')

    for kind, obj in flow:
        if kind == 'p':
            prefix, info = resolve_auto_number(obj, num_defs, counters,
                                               last_values, style_num)
            sig = para_signals(obj, len(paras), prefix, info)
            paras.append(sig)
            elements.append(obj._element)
            if info:
                auto_numbered.append({'index': sig['index'], 'text': info['text'],
                                      'fmt': info['fmt'], 'value': info['value'],
                                      'bullet': info['bullet'],
                                      # strict 下没有写进正文，只是线索
                                      'applied': not strict})
            order.append(('para', len(paras) - 1))
        else:
            rows, spans, ncol = extract_table(obj)
            entry = {'rows': rows, 'header': True, 'ncol': ncol}
            # 只有真存在合并单元格时才写 spans，保持 JSON 简洁
            if any(s for srow in spans for s in srow):
                entry['spans'] = spans
            tables.append(entry)
            order.append(('table', len(tables) - 1))

    # ── 结构性要素 ──────────────────────────────────────────
    # 标题、署名、成文日期由生成器单独输出，留在 body 里会出现两次。
    # 版头（红头 / 发文字号 / 分隔线）不属于正文，识别出来原样保留，不参与重排。
    # 不识别的话红头会被当成标题排成二号小标宋，真正的标题反而沦为正文段落。
    head_idx, head_detail = detect_letterhead(paras, elements)
    head_set = set(head_idx)

    # 正文里未完整验证的内容（图片 / 文本框 / 脚注 / 修订 …）：
    # 不能静默漏掉，必须列出来让使用者知道哪些东西没被处理
    unsupported = scan_unsupported(elements, head_set)

    # 标题取**首个非空且非版头**的段落：公文标题上方常有空行或文号行，
    # 写死首段会取到空串或红头，还会把真正的标题当正文排一遍。
    title_idx = next((i for i, s in enumerate(paras)
                      if s['text'].strip() and i not in head_set), None)

    # 成文日期：整段仅为「2026年3月17日」时几乎必然是落款日期
    date_idx = None
    if strip_sign:
        for i in range(len(paras) - 1, -1, -1):
            if i != title_idx and DATE_FULL.match(paras[i]['text'].strip()):
                date_idx = i
                break
    # 署名：落款日期之上最近的非空短段（机关名一般不以句号结尾）
    sig_idx = None
    if date_idx is not None:
        for j in range(date_idx - 1, -1, -1):
            t = paras[j]['text'].strip()
            if not t or j == title_idx:
                continue
            if len(t) <= 40 and t[-1] not in TERMINAL_PUNCT + '，,、':
                sig_idx = j
            break
    # 署名与日期**不**从 body 删除，而是原地换成 signature / date 元素。
    # 通知类公文的落款本来就在正文之后、附件之前，把它剥离到顶层字段、
    # 由生成器统一挪到文末，会让成文日期掉到最后一页——这是公文里不能犯的错。
    # 顶层字段照旧写入，供不读 body 的调用方使用。
    dropped = {i for i in (title_idx,) if i is not None}
    dropped |= head_set
    # 标题之前的空段落是版头与标题之间的间距，不是正文内容。
    # 留着它们，正文会凭空多出一段空白——实测源文件主送机关与正文紧邻，
    # 搬完这几个空行后中间就隔了一行。
    if title_idx is not None:
        for j in range(0, title_idx):
            if not paras[j]['text'].strip():
                dropped.add(j)

    # 主送机关 / 附件说明：误判代价高（删了就丢内容），**只提候选不自动剥离**
    body = []
    uncertain = []
    candidates = []
    seen_first = False

    for kind, i in order:
        if kind == 'table':
            body.append({'type': 'table', 'index': i})
            continue
        sig = paras[i]
        if i in dropped:
            continue
        # 落款就地占位：保留源文档里"正文之后、附件之前"这个位置
        if strip_sign and i == sig_idx:
            body.append({'type': 'signature',
                         'text': (sig['raw_text'] if strict
                                  else sig['text']).strip()})
            continue
        if strip_sign and i == date_idx:
            body.append({'type': 'date',
                         'text': (sig['raw_text'] if strict
                                  else sig['text']).strip()})
            continue
        # disp 只用于判读（含自动编号，能看出「二、」这类层级）；
        # 输出用 text：strict 下是原文字符本身，一个字都不加
        disp = sig['text']
        text = sig['raw_text'] if strict else sig['text']
        if (not seen_first and disp.strip() and title_idx is not None
                and i > title_idx):
            seen_first = True
            if disp.rstrip().endswith('：') and len(disp) <= 40:
                candidates.append({
                    'field': 'recipient', 'index': i, 'text': disp,
                    'note': '疑似主送机关；确认后填入 recipient，并删除 body 中该段'})
        elif ATTACH_PREFIX.match(disp.strip()):
            candidates.append({
                'field': 'attachments', 'index': i, 'text': disp,
                'note': '疑似附件说明；确认后填入 attachments，并删除 body 中该段'})

        role, conf, reason = suggest(sig)
        item = {'type': role, 'text': text}
        # strict 下自动编号不并入正文，但要把"这段其实有编号"记下来：
        # 使用者据此决定要不要手工补，脚本不替他加
        if strict and sig.get('auto_prefix'):
            item['num_prefix'] = sig['auto_prefix']
        # 分页控制原样带走：源段落设了「与下段同页」，成品也要设，否则标题会掉页底
        item.update(_para_pagination(elements[i]))
        item['_src'] = {
            'index': sig['index'],
            'style': sig['style'],
            'outlineLvl': sig['outlineLvl'],
            'bold': sig['bold'],
            'size': sig['size'],
            'autoNumber': sig.get('autoNumber'),
            'confidence': conf,
            'reason': reason,
        }
        body.append(item)
        if conf != 'high':
            uncertain.append({'index': sig['index'], 'text': disp[:60],
                              'suggested': role, 'reason': reason})

    # 去掉 body 开头的连续空段。标题与正文之间的留白由生成器统一给一个，
    # 照搬源文档那几个空行会挪错位置——主送机关单独渲染在 body 之外，
    # 源文件"标题→空→主送→正文"会被排成"标题→空→主送→空→正文"。
    k = 0
    while k < len(body) and isinstance(body[k], dict) \
            and body[k].get('type') in (None, 'blank', 'body') \
            and not (body[k].get('text') or '').strip():
        k += 1
    if k:
        body = body[k:]

    # 版记（印发机关 + 印发日期）：整块原样搬到文末，不参与重排
    version_record, body = detect_version_record(elements, body)
    if version_record:
        # 版记里的线条是原样搬走的，不该再报"图形不保留"
        vset = {version_record['source_index']}
        if len(version_record.get('blocks') or []) > 1:
            vset.add(version_record['source_index'] - 1)
        for u in unsupported:
            u['index'] = [i for i in u['index'] if i not in vset]
            u['count'] = len(u['index'])
        unsupported = [u for u in unsupported if u['count']]

    if title_idx is None:
        title = ''
    else:
        title = (paras[title_idx]['raw_text'] if strict
                 else paras[title_idx]['text']).strip()
    # 标题的「与下段同页」：公文标题几乎都设了，丢了就是标题落在页末
    title_pag = (_para_pagination(elements[title_idx])
                 if title_idx is not None else {})

    # 页码：源文档没有就不加，有则沿用原样式
    pn = detect_page_number(doc)

    result = {
        'title': title,
        'page_number': pn['mode'],
        'fidelity': fidelity,
        'body': body,
        '_uncertain': uncertain,
        '_note': ('_src / _uncertain / _candidates 为判读线索，可删除；'
                  'type 定稿后交给 generate_gongwen.py。'
                  '表格保留在 tables，body 中的 {"type":"table"} 决定其位置。'
                  'page_number 由源文档页码自动推断：'
                  '源文件无页码→none，有→沿用 center/odd_even。'
                  '标题取首个非空段落；落款署名与成文日期已从 body 剥离'
                  '（见 signature / date / _stripped），避免重复输出。'
                  '版头（红头/发文字号/分隔线）不参与重排，由 header 记录'
                  '并原样复制到新文档——不要替用户重做红头。'
                  'keep_next / keep_lines 是源段落的分页控制'
                  '（与下段同页 / 段中不分页），原样带走到成品；'
                  '不需要可删除。'),
    }
    if strict:
        result['_note'] += (
            '【严格保真】fidelity=strict：正文每段只输出源文件里真实存在的字符。'
            'Word 自动编号（w:numPr）不并入正文——它在源文件里不是字符，'
            '补进去就是加字；带自动编号的段落另给 num_prefix 字段记录待补内容，'
            '是否补由使用者决定。需要补回时用 --fidelity standard 重新提取。')
    if title_pag.get('keep_next'):
        # 只在确实设了的时候写，没设就不写——保持 JSON 干净
        result['title_keep_next'] = True
    if version_record:
        result['version_record'] = version_record
    if sig_idx is not None:
        result['signature'] = (paras[sig_idx]['raw_text'] if strict
                               else paras[sig_idx]['text']).strip()
    if date_idx is not None:
        result['date'] = (paras[date_idx]['raw_text'] if strict
                          else paras[date_idx]['text']).strip()
    stripped = []
    for i, field in ((sig_idx, 'signature'), (date_idx, 'date')):
        if i is None:
            continue
        # 剥离是"猜"出来的，依据必须一并给出，供生成前核对：
        # 猜错了就是丢一段内容，不能让使用者无从判断
        rec = {
            'index': i,
            'field': field,
            'text': (paras[i]['raw_text'] if strict
                     else paras[i]['text']).strip(),
            'reason': (STRIP_REASON['date'] if field == 'date'
                       else STRIP_REASON['signature']),
            'confidence': 'high' if field == 'date' else 'medium',
            'context': {
                'prev': paras[i - 1]['text'].strip() if i > 0 else None,
                'next': (paras[i + 1]['text'].strip()
                         if i + 1 < len(paras) else None),
            },
            'note': '若判断有误，请把该段放回 body 并清空对应字段'
                    '（或使用 --no-strip 不剥离）',
        }
        stripped.append(rec)
    if stripped:
        result['_stripped'] = stripped
    if head_idx:
        result['header'] = {
            'source': os.path.abspath(path),
            'indices': head_idx,
            # 文本框内的字 p.text 取不到，用判定时的 probe 才能看到实际内容
            'texts': [next((d['text'] for d in head_detail if d['index'] == i),
                           paras[i]['text'].strip()) for i in head_idx],
        }
        result['_head_detail'] = head_detail
    if candidates:
        result['_candidates'] = candidates
    if pn['detail']:
        result['_page_number'] = pn['detail']
    if auto_numbered:
        result['_auto_number'] = auto_numbered
    for u in unsupported:
        idx = u['index'][:5]
        more = '…' if u['count'] > 5 else ''
        warnings.append(f'正文含「{u["kind"]}」（段落 {idx}{more}）：{u["hint"]}')
    conflicts = detect_number_conflicts(paras)
    if conflicts:
        result['_number_conflicts'] = conflicts
    if unsupported:
        result['_unsupported'] = unsupported
    if warnings:
        result['_warnings'] = warnings
    if tables:
        result['tables'] = tables
    return result


def main():
    ap = argparse.ArgumentParser(
        description='提取 docx 段落并给出角色建议（不修改原文）')
    ap.add_argument('docx', help='源 .docx 路径')
    ap.add_argument('-o', '--output', help='输出 JSON 路径，省略则打印到标准输出')
    ap.add_argument('--no-meta', action='store_true',
                    help='不输出 _src 线索字段，只保留 type/text')
    ap.add_argument('--no-strip', action='store_true',
                    help='不剥离署名与成文日期，两段都留在 body'
                         '（自动剥离在特殊文档中可能判错时使用）')
    ap.add_argument('--fidelity', choices=FIDELITY_MODES, default=FIDELITY_DEFAULT,
                    help='strict（默认）：原文一个字不增不减，Word 自动编号只记线索、'
                         '不并入正文；standard：把自动编号补回成文字')
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        print(f'文件不存在: {args.docx}', file=sys.stderr)
        return 1

    result = build(args.docx, strip_sign=not args.no_strip,
                   fidelity=args.fidelity)
    if args.no_meta:
        for item in result['body']:
            item.pop('_src', None)

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        out_dir = os.path.dirname(os.path.abspath(args.output))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(text)
        n = len(result['body'])
        u = len(result['_uncertain'])
        print(f'已提取 {n} 段 → {args.output}')
        print(f'其中 {u} 段存疑，请审阅 _uncertain 列表后确认 type')
        mode = result['page_number']
        if mode == 'none':
            print('页码：源文件未发现页码 → 设为 none（不添加页码）')
        else:
            print(f'页码：源文件已有页码 → 沿用 {mode}')
        an = result.get('_auto_number') or []
        if an:
            strict = result.get('fidelity') == 'strict'
            if strict:
                print(f'自动编号：{len(an)} 处由 Word 生成（源文件里不是字符）。'
                      f'【严格保真】未并入正文，编号写在对应段的 num_prefix 字段，'
                      f'是否补由你决定；要补请用 --fidelity standard 重新提取：')
            else:
                print(f'自动编号：{len(an)} 处由 Word 生成（不在原文字符里），已补回文本：')
            for a in an:
                print(f'   第{a["index"]}段 → 「{a["text"]}」'
                      f'（{a["fmt"]}，值={a["value"]}）')
        st = result.get('_stripped') or []
        if st:
            print('落款：已从正文剥离（生成器会重新输出，留在正文里会重复）。'
                  '这是按格式特征推断的，请核对判断是否属实：')
            for s in st:
                print(f"   第{s['index']}段 → {s['field']}「{s['text']}」"
                      f"（{s['reason']}，置信度 {s['confidence']}）")
            print('   若判错：把该段放回 body 并清空对应字段，或用 --no-strip 重新提取')
        cand = result.get('_candidates') or []
        if cand:
            print(f'候选要素 {len(cand)} 处（误判代价高，未自动剥离，请确认）：')
            for c in cand:
                print(f"   第{c['index']}段 → {c['field']}：{c['text'][:30]}")
        head = result.get('header')
        if head:
            print(f"版头：检测到 {len(head['indices'])} 段（红头/发文字号/分隔线）"
                  f" → 原样保留，不参与重排；明细见 _head_detail")
            for t in head['texts']:
                if t:
                    print(f'   {t}')
        conf = result.get('_number_conflicts') or []
        if conf:
            print(f'编号冲突：{len(conf)} 处同级编号重复，请核对（可能是手打编号与'
                  f' Word 自动编号各算一套）：')
            for c in conf:
                print(f"   「{c['value']}」在第{c['first']['index']}段"
                      f"（{c['first']['source']}）与第{c['again']['index']}段"
                      f"（{c['again']['source']}）重复：{c['text']}")
        vr = result.get('version_record')
        if vr:
            print(f"版记：检测到印发机关与印发日期 → 整块原样搬到文末，不参与重排"
                  f"（含 {len(vr.get('blocks') or [])} 段，含分隔线）")
            print(f'   {vr["text"][:60]}')
        for w in result.get('_warnings') or []:
            print(f'提示：{w}')
        tabs = result.get('tables') or []
        if tabs:
            merged = sum(1 for t in tabs if t.get('spans'))
            extra = f'，其中 {merged} 个含合并单元格（已保留合并结构）' if merged else ''
            print(f'表格：{len(tabs)} 个{extra}')
    else:
        print(text)
    return 0


if __name__ == '__main__':
    sys.exit(main())

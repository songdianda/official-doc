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
# 附件说明（7.3.4）。
ATTACH_PREFIX = re.compile(r'^附件[:：]')


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
    if f in ('japanesecounting', 'chinesecounting', 'chineselegalsimplified',
             'ideographdigital', 'ideographtraditional'):
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


def resolve_auto_number(p, defs, counters, last_values):
    """算出 Word 自动编号在该段应该显示的文本（如「二、」）。

    python-docx 的 p.text 只取 w:t，自动编号不在文本里而挂在 w:numPr 上，
    直接提取会**整条丢失**——源文件里手打的「一、」在，Word 自动生成的
    「二、」就没了。

    返回 (显示文本, 线索dict)；无自动编号返回 (None, None)。
    """
    if not defs:
        return None, None
    pPr = p._element.pPr
    if pPr is None:
        return None, None
    numPr = pPr.find(qn('w:numPr'))
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


def para_signals(p, index, auto_prefix=None, auto_info=None):
    """采集一段的样式线索。auto_prefix 为 Word 自动编号渲染出的文本。"""
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
    text = p.text.strip()
    if auto_prefix:
        text = auto_prefix + text
    return {
        'index': index,
        'text': text,
        'style': p.style.name if p.style is not None else None,
        'styleId': p.style.style_id if p.style is not None else None,
        'outlineLvl': outline,
        'bold': bool(r.font.bold) if (r is not None and r.font.bold is not None) else False,
        'size': (r.font.size.pt if (r is not None and r.font.size) else None),
        'font': (rFonts.get(qn('w:eastAsia')) if rFonts is not None else None),
        'listPara': numPr,          # Word 自动编号/项目符号
        'autoNumber': auto_info,    # 自动编号线索（由 numbering.xml 算出，非原文字符）
        'alignment': str(p.alignment) if p.alignment is not None else None,
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

    # 2) 无语义标记，看编号 + 句法
    if prefix_level:
        if text.rstrip()[-1] in TERMINAL_PUNCT:
            return ('numbered_body', 'high',
                    '以句号结尾的完整叙述句，是罗列的工作事项而非标题')
        if prefix_level <= 2:
            # 汉字编号（一、/（一））不会是正文罗列，无句号即标题
            return (f'heading{prefix_level}', 'high',
                    '汉字编号且不以句号结尾，是标题')
        # 阿拉伯数字编号 + 无句号：可能标题，也可能省略标点的短事项 → 存疑
        return (f'heading{prefix_level}', 'low',
                '阿拉伯数字编号且不以句号结尾，'
                '可能是标题，也可能是省略标点的短事项')
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


def build(path):
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
    tables = []
    auto_numbered = []
    # 占位：body 中记录 ('para', sig) / ('table', index)
    order = []

    # Word 自动编号（w:numPr）不在 w:t 里，必须单独解析，否则整条编号丢失
    num_defs = load_numbering_defs(doc)
    counters, last_values = {}, {}

    for kind, obj in flow:
        if kind == 'p':
            prefix, info = resolve_auto_number(obj, num_defs, counters, last_values)
            sig = para_signals(obj, len(paras), prefix, info)
            paras.append(sig)
            if info:
                auto_numbered.append({'index': sig['index'], 'text': info['text'],
                                      'fmt': info['fmt'], 'value': info['value'],
                                      'bullet': info['bullet']})
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
    # 标题取**首个非空段落**：公文标题上方常有空行或文号行，写死首段会取到空串，
    # 还会把真正的标题当正文排一遍。
    title_idx = next((i for i, s in enumerate(paras) if s['text'].strip()), None)

    # 成文日期：整段仅为「2026年3月17日」时几乎必然是落款日期
    date_idx = None
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
    dropped = {i for i in (title_idx, sig_idx, date_idx) if i is not None}
    # 落款块前后的空段落一并清掉：源文件里那些用来把落款"顶下去"的空行，
    # 生成器会在落款前后各补一个空行，留着就是一堆空段落噪音
    anchor = sig_idx if sig_idx is not None else date_idx
    if anchor is not None:
        j = anchor - 1
        while j >= 0 and j != title_idx and not paras[j]['text'].strip():
            dropped.add(j)
            j -= 1
        for j in range((date_idx if date_idx is not None else anchor) + 1, len(paras)):
            if not paras[j]['text'].strip():
                dropped.add(j)
            else:
                break

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
        text = sig['text']
        if (not seen_first and text.strip() and title_idx is not None
                and i > title_idx):
            seen_first = True
            if text.rstrip().endswith('：') and len(text) <= 40:
                candidates.append({
                    'field': 'recipient', 'index': i, 'text': text,
                    'note': '疑似主送机关；确认后填入 recipient，并删除 body 中该段'})
        elif ATTACH_PREFIX.match(text.strip()):
            candidates.append({
                'field': 'attachments', 'index': i, 'text': text,
                'note': '疑似附件说明；确认后填入 attachments，并删除 body 中该段'})

        role, conf, reason = suggest(sig)
        item = {'type': role, 'text': text}
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
            uncertain.append({'index': sig['index'], 'text': sig['text'][:60],
                              'suggested': role, 'reason': reason})

    title = paras[title_idx]['text'].strip() if title_idx is not None else ''

    # 页码：源文档没有就不加，有则沿用原样式
    pn = detect_page_number(doc)

    result = {
        'title': title,
        'page_number': pn['mode'],
        'body': body,
        '_uncertain': uncertain,
        '_note': ('_src / _uncertain / _candidates 为判读线索，可删除；'
                  'type 定稿后交给 generate_gongwen.py。'
                  '表格保留在 tables，body 中的 {"type":"table"} 决定其位置。'
                  'page_number 由源文档页码自动推断：'
                  '源文件无页码→none，有→沿用 center/odd_even。'
                  '标题取首个非空段落；落款署名与成文日期已从 body 剥离'
                  '（见 signature / date / _stripped），避免重复输出。'),
    }
    if sig_idx is not None:
        result['signature'] = paras[sig_idx]['text'].strip()
    if date_idx is not None:
        result['date'] = paras[date_idx]['text'].strip()
    stripped = []
    for i, field in ((sig_idx, 'signature'), (date_idx, 'date')):
        if i is not None:
            stripped.append({'index': i, 'field': field,
                             'text': paras[i]['text'].strip()})
    if stripped:
        result['_stripped'] = stripped
    if candidates:
        result['_candidates'] = candidates
    if pn['detail']:
        result['_page_number'] = pn['detail']
    if auto_numbered:
        result['_auto_number'] = auto_numbered
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
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        print(f'文件不存在: {args.docx}', file=sys.stderr)
        return 1

    result = build(args.docx)
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
            print(f'自动编号：{len(an)} 处由 Word 生成（不在原文字符里），已补回文本：')
            for a in an:
                print(f'   第{a["index"]}段 → 「{a["text"]}」'
                      f'（{a["fmt"]}，值={a["value"]}）')
        if result.get('signature') or result.get('date'):
            print('落款：已从正文剥离（生成器会重新输出，留在正文里会重复） → '
                  f"署名「{result.get('signature') or '（无）'}」"
                  f" 日期「{result.get('date') or '（无）'}」，见 _stripped")
        cand = result.get('_candidates') or []
        if cand:
            print(f'候选要素 {len(cand)} 处（误判代价高，未自动剥离，请确认）：')
            for c in cand:
                print(f"   第{c['index']}段 → {c['field']}：{c['text'][:30]}")
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

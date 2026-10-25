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
TERMINAL_PUNCT = '。；；！？'


def para_signals(p, index):
    """采集一段的样式线索。"""
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
    numPr = None
    if pPr is not None and pPr.find(qn('w:numPr')) is not None:
        numPr = True
    return {
        'index': index,
        'text': p.text.strip(),
        'style': p.style.name if p.style is not None else None,
        'styleId': p.style.style_id if p.style is not None else None,
        'outlineLvl': outline,
        'bold': bool(r.font.bold) if (r is not None and r.font.bold is not None) else False,
        'size': (r.font.size.pt if (r is not None and r.font.size) else None),
        'font': (rFonts.get(qn('w:eastAsia')) if rFonts is not None else None),
        'listPara': numPr,          # Word 自动编号/项目符号
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
    # 占位：body 中记录 ('para', sig) / ('table', index)
    order = []
    for kind, obj in flow:
        if kind == 'p':
            sig = para_signals(obj, len(paras))
            paras.append(sig)
            order.append(('para', len(paras) - 1))
        else:
            rows = [[c.text.strip() for c in row.cells] for row in obj.rows]
            tables.append({'rows': rows, 'header': True})
            order.append(('table', len(tables) - 1))

    # 约定：首段为标题，只进 title 字段，不进 body（否则会重复输出一次）
    body = []
    uncertain = []
    for kind, i in order:
        if kind == 'table':
            body.append({'type': 'table', 'index': i})
            continue
        sig = paras[i]
        if sig['index'] == 0:
            continue        # 首段已作 title
        role, conf, reason = suggest(sig)
        item = {'type': role, 'text': sig['text']}
        item['_src'] = {
            'index': sig['index'],
            'style': sig['style'],
            'outlineLvl': sig['outlineLvl'],
            'bold': sig['bold'],
            'size': sig['size'],
            'confidence': conf,
            'reason': reason,
        }
        body.append(item)
        if conf != 'high':
            uncertain.append({'index': sig['index'], 'text': sig['text'][:60],
                              'suggested': role, 'reason': reason})

    title = paras[0]['text'] if paras and paras[0]['text'] else ''

    # 页码：源文档没有就不加，有则沿用原样式
    pn = detect_page_number(doc)

    result = {
        'title': title,
        'page_number': pn['mode'],
        'body': body,
        '_uncertain': uncertain,
        '_note': ('_src 与 _uncertain 为判读线索，可删除；'
                  'type 定稿后交给 generate_gongwen.py。'
                  '表格保留在 tables，body 中的 {"type":"table"} 决定其位置。'
                  'page_number 由源文档页码自动推断：'
                  '源文件无页码→none，有→沿用 center/odd_even。'),
    }
    if pn['detail']:
        result['_page_number'] = pn['detail']
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
    else:
        print(text)
    return 0


if __name__ == '__main__':
    sys.exit(main())

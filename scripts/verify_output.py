#!/usr/bin/env python3
"""生成前后内容验收：比对段落文字、表格与图片，优先拦住漏字、漏表、漏图。

用法：
    python3 scripts/verify_output.py 源文件.docx 成品.docx [提取结果.json]

这个脚本**不检查版式**。版式（字体、行距、页码位置）要另外逐页看。
它只回答一个问题：源文档里的东西，成品里还在不在。

判定方式
  段落文字  多重集合比对：源里出现 N 次，成品里至少要出现 N 次。
            少了就是漏字，多出来的会在报告里列出供确认。
  表格      先比数量，再逐个比单元格文字（合并格按还原后的内容比）。
  图片      比数量。正文图片本技能不保留，会明确报出来而不是静默。

有意调整的部分（不是缺陷，会在报告里单独说明）
  - 落款署名 / 成文日期被剥离到 signature / date，成品中位置改变但内容仍在
  - Word 自动编号被补回成文字
  - 版头整段原样复制（文本框内的字无法用文本比对，需目视确认）
  - 落款前后用于「顶下去」的空段落被清理
"""

import argparse
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                      # noqa: E402
from docx.oxml.ns import qn                    # noqa: E402


def _count_images(el):
    """段落里的图片数量（w:drawing / 老式 w:pict）。"""
    return (sum(1 for _ in el.iter(qn('w:drawing')))
            + sum(1 for _ in el.iter(qn('w:pict'))))


def doc_images(path):
    """文档中的图片总数，含表格与页眉页脚里的。"""
    doc = Document(path)
    n = 0
    for child in doc.element.body.iterchildren():
        n += _count_images(child)
    return n


def doc_paragraphs_and_tables(path):
    """取出成品的段落文字、表格单元格文字。"""
    doc = Document(path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for t in doc.tables:
        rows = []
        for row in t.rows:
            # 合并单元格会被 python-docx 重复展开，取出现过的文本集合更稳
            rows.append([c.text.strip() for c in row.cells])
        tables.append(rows)
    return paras, tables


def flatten(rows):
    """把二维表格摊成一维，且对合并格去重后再计数。

    python-docx 的 row.cells 会把跨列合并的格子读成多份相同内容，
    直接比对会凭空多出格子，这里先按行去重。
    """
    out = []
    for row in rows:
        seen, uniq = set(), []
        for c in row:
            if c in seen:
                continue
            seen.add(c)
            uniq.append(c)
        out.extend(x for x in uniq if x)
    return out


def load_source(path, meta):
    """源文档期望清单：段落文字、表格、图片数。"""
    if meta:
        items = []
        if meta.get('title'):
            items.append(meta['title'].strip())
        for it in meta.get('body', []):
            txt = (it.get('text') or '').strip() if isinstance(it, dict) else ''
            if txt:
                items.append(txt)
        for key in ('signature', 'date'):
            if meta.get(key):
                items.append(str(meta[key]).strip())
        head = (meta.get('header') or {}).get('indices') or []
        tables = [t.get('rows', []) for t in (meta.get('tables') or [])]
    else:
        # 没有提取结果时降级：直接读源文档文本（不含自动编号补回）
        doc = Document(path)
        items = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        head = []
        tables = [[[c.text.strip() for c in row.cells] for row in t.rows]
                  for t in doc.tables]
    return items, tables, doc_images(path), head


def main():
    ap = argparse.ArgumentParser(
        description='比对源文档与成品的内容，检查是否有漏字、漏表、漏图')
    ap.add_argument('source', help='源 .docx')
    ap.add_argument('target', help='生成的 .docx')
    ap.add_argument('meta', nargs='?',
                    help='inspect_docx.py 输出的 JSON（可选，用了更准）')
    args = ap.parse_args()

    for p in (args.source, args.target):
        if not os.path.exists(p):
            print(f'文件不存在: {p}', file=sys.stderr)
            return 1

    meta = None
    if args.meta:
        import json
        if not os.path.exists(args.meta):
            print(f'提取结果不存在: {args.meta}', file=sys.stderr)
            return 1
        meta = json.load(open(args.meta, encoding='utf-8'))

    src_items, src_tables, src_img, head_idx = load_source(args.source, meta)
    dst_items, dst_tables = doc_paragraphs_and_tables(args.target)
    dst_img = doc_images(args.target)

    ok = True
    print('──────── 内容验收报告 ────────')
    print(f'源文件：{os.path.basename(args.source)}')
    print(f'成品　：{os.path.basename(args.target)}')

    # 段落文字
    cs, cd = Counter(src_items), Counter(dst_items)
    missing = cs - cd
    extra = cd - cs
    print(f'\n【段落文字】源 {len(src_items)} 段（去重后 {len(cs)} 种），'
          f'成品 {len(dst_items)} 段')
    if missing:
        ok = False
        print(f'  缺失 {sum(missing.values())} 处（源有、成品没有或次数不足）：')
        for t, n in list(missing.items())[:20]:
            print(f'    ×{n}  {t[:60]}')
        if len(missing) > 20:
            print(f'    …… 另有 {len(missing) - 20} 种')
    else:
        print('  源文件所有段落文字均已在成品中出现')
    if extra:
        print(f'  成品多出 {sum(extra.values())} 处（多为页码、空行或补回的编号）：')
        for t, n in list(extra.items())[:10]:
            print(f'    +{n}  {t[:60]}')
        if len(extra) > 10:
            print(f'    …… 另有 {len(extra) - 10} 种')

    # 表格
    print(f'\n【表格】源 {len(src_tables)} 个，成品 {len(dst_tables)} 个')
    if len(src_tables) != len(dst_tables):
        ok = False
        print(f'  数量不一致——成品可能缺表')
    for i, (s, d) in enumerate(zip(src_tables, dst_tables)):
        fs, fd = Counter(flatten(s)), Counter(flatten(d))
        miss, ext = fs - fd, fd - fs
        if miss or ext:
            ok = False
            print(f'  表{i + 1}：源 {len(s)} 行 / 成品 {len(d)} 行，内容有差异')
            for t, n in list(miss.items())[:8]:
                print(f'      源有成品无 ×{n}：{t[:40]}')
            for t, n in list(ext.items())[:8]:
                print(f'      成品有源无 +{n}：{t[:40]}')
        else:
            print(f'  表{i + 1}：逐格一致（{len(s)} 行，共 {sum(fs.values())} 个有内容的格）')

    # 图片
    print(f'\n【图片】源 {src_img} 个，成品 {dst_img} 个')
    if src_img != dst_img:
        print(f'  数量不一致。若源文档正文含图片：本技能不保留正文图片，'
              f'需手工补回')
        # 版头图片会被复制，只有正文图片才算真丢失
        if meta and meta.get('header'):
            print('  （源文件含版头，版头中的图片已原样复制）')
        ok = False
    else:
        print('  一致')

    # 有意调整的部分
    notes = []
    if meta:
        st = meta.get('_stripped') or []
        if st:
            notes.append('落款已剥离后由生成器重新输出：'
                         + '、'.join(f"{s['field']}「{s['text']}」" for s in st))
        an = meta.get('_auto_number') or []
        if an:
            notes.append(f'Word 自动编号 {len(an)} 处已补回为文字（源文件里不是字符）')
        if meta.get('header'):
            n = len(meta['header']['indices'])
            notes.append(f'版头 {n} 段整段原样复制；'
                         f'若含文本框，其中文字无法用文本比对，需目视确认')
        uns = meta.get('_unsupported') or []
        for u in uns:
            notes.append(f"源文档含「{u['kind']}」{u['count']} 处：{u['hint']}")
    if notes:
        print('\n【有意调整 / 需留意】')
        for n in notes:
            print(f'  - {n}')

    print('\n结论：' + ('内容一致，未发现漏字、漏表、漏图' if ok
                    else '存在内容差异，请按上述条目核对后再交付'))
    print('注意：本脚本只比对内容，不能证明版式正确，版式仍需逐页查看。')
    return 0 if ok else 2


if __name__ == '__main__':
    sys.exit(main())

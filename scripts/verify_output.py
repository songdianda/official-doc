#!/usr/bin/env python3
"""生成前后内容验收：比对段落文字、表格与图片，优先拦住漏字、漏表、漏图。

用法：
    python3 scripts/verify_output.py 源文件.docx 成品.docx [提取结果.json]

这个脚本**不检查版式**。版式（字体、行距、页码位置）要另外逐页看。
它只回答一个问题：源文档里的东西，成品里还在不在。

三道检查，各自独立、互不替代
  1. 原文覆盖    **直接从源文档读**，不经过提取结果。
                 源里每一段（含表格格、页眉页脚）都必须在成品里找得到，
                 找不到就是丢了。这一道专门防「提取 JSON 和成品同时漏掉同一段」
                 ——那种情况若只拿 JSON 当基准，会得出"一致"的假结论。
  2. 表格逐格    按 (行, 列) 坐标比对，不是比内容集合。
                 集合比对抓不到"两格内容互换"和"重复值多一个少一个"。
                 合并单元格先按 w:gridSpan / w:vMerge 还原成逻辑格再比。
  3. 图片        按部位分别计数：正文 / 页眉页脚 / 脚注尾注批注。
                 只看 body 会把页眉里的图片数成 0，页眉图丢了也报"一致"。

有意调整的部分（不是缺陷，会在报告里单独说明）
  - 落款署名 / 成文日期被剥离到 signature / date，成品中位置改变但内容仍在
  - Word 自动编号被补回成文字（源文件里它不是字符）
  - 版头整段原样复制（文本框内的字需目视确认）
  - 落款前后用于「顶下去」的空段落被清理
"""

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                      # noqa: E402
from docx.oxml.ns import qn                    # noqa: E402

# 只复用纯 XML 工具函数，不复用提取逻辑——验收必须独立于被验收的对象
from inspect_docx import extract_table         # noqa: E402


def _norm(s):
    """去掉所有空白后比较，避免 Word 换行、空格差异造成误报。"""
    return re.sub(r'\s+', '', s or '')


def _all_text(el):
    """元素内全部 w:t 文字（含文本框内的）。

    文本框里的 w:t 嵌在 w:pict 深处，python-docx 的 p.text 用相对 xpath
    只取直接子节点，取不到——红头是文本框时会读成空段，必须自己遍历。
    """
    return ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()


# ── 图片计数 ─────────────────────────────────────────────

def _count_images(el):
    """XML 元素里的图片引用数（w:drawing / 老式 w:pict）。"""
    return (sum(1 for _ in el.iter(qn('w:drawing')))
            + sum(1 for _ in el.iter(qn('w:pict'))))


def _classify(doc):
    """把包内部件按部位分类。

    按 partname 分类而不按 section.header / section.footer 取：
    首页页眉、偶数页页眉不一定挂在默认属性上，逐个属性取会漏；
    直接扫部件则只要是包里的页眉页脚就都在。
    """
    zones = {'body': [], 'header_footer': [], 'other': []}
    for part in doc.part.package.parts:
        name = str(part.partname)
        if not name.startswith('/word/') or not hasattr(part, 'element'):
            continue
        base = name.rsplit('/', 1)[-1]
        if base == 'document.xml':
            zones['body'].append(part.element)
        elif base.startswith('header') or base.startswith('footer'):
            zones['header_footer'].append(part.element)
        elif base in ('footnotes.xml', 'endnotes.xml', 'comments.xml'):
            zones['other'].append(part.element)
    return zones


def image_stats(path):
    """按部位统计图片：正文 / 页眉页脚 / 脚注尾注批注。

    只数 body 是不够的：红头图片常在页眉里，页眉图丢了却报「双方 0 张」，
    是最容易被忽略的一类内容丢失。
    """
    doc = Document(path)
    zones = _classify(doc)
    stats = {k: sum(_count_images(el) for el in v) for k, v in zones.items()}
    # 去重后的图片文件数（同一张图在多个页眉被引用时只算一份）
    stats['files'] = len({str(p.partname) for p in doc.part.package.parts
                          if '/media/' in str(p.partname)})
    return stats


# ── 直接从文档读原文（不依赖任何提取结果）─────────────────

def raw_source(path):
    """源文档里"应该被带走的全部文字"，逐条带位置。

    返回 list of {'text', 'where', 'zone'}，zone ∈ 正文 / 表格 / 页眉页脚。
    这一份清单**不经过** inspect_docx 的角色判定与落款剥离，
    是验收的独立基准：提取环节漏掉的东西，在这里仍然看得到。
    """
    doc = Document(path)
    items = []

    for i, p in enumerate(doc.paragraphs):
        text = _all_text(p._element)
        if text:
            items.append({'text': text, 'where': f'正文段落[{i}]',
                          'zone': '正文'})

    for ti, tbl in enumerate(doc.tables):
        rows, _spans, _ncol = extract_table(tbl)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                if val.strip():
                    items.append({'text': val.strip(),
                                  'where': f'表{ti + 1}({r + 1}行{c + 1}列)',
                                  'zone': '表格'})

    for el in _classify(doc)['header_footer']:
        text = _all_text(el)
        if text:
            items.append({'text': text, 'where': '页眉/页脚',
                          'zone': '页眉页脚'})
    for el in _classify(doc)['other']:
        text = _all_text(el)
        if text:
            items.append({'text': text, 'where': '脚注/尾注/批注',
                          'zone': '其他'})
    return items


def target_texts(path):
    """成品里全部可比对文字：正文段落 + 表格格 + 页眉页脚。"""
    doc = Document(path)
    out = []
    for p in doc.paragraphs:
        text = _all_text(p._element)
        if text:
            out.append(text)
    for tbl in doc.tables:
        rows, _spans, _ncol = extract_table(tbl)
        for row in rows:
            out.extend(v.strip() for v in row if v.strip())
    zones = _classify(doc)
    for key in ('header_footer', 'other'):
        for el in zones[key]:
            text = _all_text(el)
            if text:
                out.append(text)
    return out


def match_one(text, pool):
    """在 pool 中找一条能覆盖 text 的记录（互为子串即算命中）。

    命中后把该条从 pool 里消耗掉：源里有两段相同的文字，成品只有一段时，
    第二段会找不到配对，从而暴露"少了一段"——只比集合是发现不了的。
    """
    n = _norm(text)
    if not n:
        return True
    for i, cand in enumerate(pool):
        if n in _norm(cand) or _norm(cand) in n:
            pool.pop(i)
            return True
    return False


# ── 表格：按坐标逐格比对 ──────────────────────────────────

def table_grid(path):
    """读文档的表格，还原合并后返回 [(行数, 列数, {(r,c): 文本})]。"""
    doc = Document(path)
    grids = []
    for tbl in doc.tables:
        rows, spans, ncol = extract_table(tbl)
        cells = {}
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                cells[(r, c)] = val.strip()
        grids.append((len(rows), ncol, cells))
    return grids


def compare_tables(src_grids, dst_grids):
    """逐格比对，返回 (是否一致, 差异明细)。"""
    diffs = []
    if len(src_grids) != len(dst_grids):
        diffs.append(f'表格数量不一致：源 {len(src_grids)} 个，'
                     f'成品 {len(dst_grids)} 个')
    for i, (s, d) in enumerate(zip(src_grids, dst_grids)):
        srow, scol, scells = s
        drow, dcol, dcells = d
        if (srow, scol) != (drow, dcol):
            diffs.append(f'表{i + 1} 结构不同：源 {srow}行×{scol}列，'
                         f'成品 {drow}行×{dcol}列')
        keys = sorted(set(scells) | set(dcells))
        for (r, c) in keys:
            a, b = scells.get((r, c), ''), dcells.get((r, c), '')
            if a != b:
                diffs.append(f'表{i + 1}({r + 1}行{c + 1}列)：'
                             f'源「{a[:30]}」→ 成品「{b[:30]}」')
    return (not diffs), diffs


def json_items(meta):
    """提取结果（JSON）里的文字清单，用于判断漏在哪个环节。"""
    items = []
    if not meta:
        return items
    if meta.get('title'):
        items.append(str(meta['title']).strip())
    for it in meta.get('body', []):
        txt = (it.get('text') or '').strip() if isinstance(it, dict) else ''
        if txt:
            items.append(txt)
    for key in ('signature', 'date'):
        if meta.get(key):
            items.append(str(meta[key]).strip())
    for t in (meta.get('tables') or []):
        for row in t.get('rows', []):
            items.extend(str(v).strip() for v in row if str(v).strip())
    for a in (meta.get('attachments') or []):
        items.append(str(a).strip())
    return [x for x in items if x]


def main():
    ap = argparse.ArgumentParser(
        description='比对源文档与成品的内容，检查是否有漏字、漏表、漏图')
    ap.add_argument('source', help='源 .docx')
    ap.add_argument('target', help='生成的 .docx')
    ap.add_argument('meta', nargs='?',
                    help='inspect_docx.py 输出的 JSON（可选，用于定位漏在哪个环节）')
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

    ok = True
    warns = []
    print('──────── 内容验收报告 ────────')
    print(f'源文件：{os.path.basename(args.source)}')
    print(f'成品　：{os.path.basename(args.target)}')

    # ── 1. 原文覆盖（基准来自源文档本身）────────────────
    src_items = raw_source(args.source)
    pool = target_texts(args.target)
    jpool = json_items(meta)
    missing = []
    for it in src_items:
        if not match_one(it['text'], pool):
            missing.append(it)
    print(f'\n【原文覆盖】源文档 {len(src_items)} 条文字，逐条在成品中查找')
    if missing:
        hard = [m for m in missing if m['zone'] in ('正文', '表格')]
        soft = [m for m in missing if m['zone'] not in ('正文', '表格')]
        if hard:
            ok = False
            print(f'  成品中找不到 {len(hard)} 条（内容丢失）：')
            for m in hard[:20]:
                mark = ''
                if meta and not match_one(m['text'], jpool):
                    mark = '　[提取结果里也没有 → 漏在提取环节，不是排版环节]'
                print(f'    × {m["where"]}  {m["text"][:60]}{mark}')
            if len(hard) > 20:
                print(f'    …… 另有 {len(hard) - 20} 条')
        if not hard:
            print('  正文与表格文字均可在成品中找到'
                  '（页眉页脚 / 脚注等缺失见文末提示，不计入失败）')
        for zone in ('页眉页脚', '其他'):
            group = [m for m in soft if m['zone'] == zone]
            if group:
                warns.append(
                    f'{zone}文字 {len(group)} 条未进入成品'
                    f'（本技能不复制{zone}内容，需手工补回）：'
                    + '、'.join(m['text'][:20] for m in group[:5]))
    else:
        print('  源文档全部文字均可在成品中找到')

    # ── 2. 表格逐格（带位置）────────────────────────────
    src_grids = table_grid(args.source)
    dst_grids = table_grid(args.target)
    same, diffs = compare_tables(src_grids, dst_grids)
    total_cells = sum(len(g[2]) for g in src_grids)
    print(f'\n【表格逐格】源 {len(src_grids)} 个表 / 成品 {len(dst_grids)} 个表，'
          f'源共 {total_cells} 个格位')
    if not src_grids and not dst_grids:
        print('  双方均无表格')
    elif same:
        print(f'  逐格一致：位置与内容均相同（含合并单元格还原）')
    else:
        ok = False
        print(f'  发现 {len(diffs)} 处差异（位置不同也算差异）：')
        for d in diffs[:20]:
            print(f'    × {d}')
        if len(diffs) > 20:
            print(f'    …… 另有 {len(diffs) - 20} 处')

    # ── 3. 图片（分部位）───────────────────────────────
    si, di = image_stats(args.source), image_stats(args.target)
    print(f'\n【图片】源 正文 {si["body"]} / 页眉页脚 {si["header_footer"]}'
          f' / 其他 {si["other"]}（去重后 {si["files"]} 个图片文件）')
    print(f'　　　　成品 正文 {di["body"]} / 页眉页脚 {di["header_footer"]}'
          f' / 其他 {di["other"]}（去重后 {di["files"]} 个图片文件）')
    if si['body'] != di['body']:
        ok = False
        if si['body'] > di['body']:
            print(f'  正文图片少 {si["body"] - di["body"]} 个：'
                  f'本技能不重排正文图片，需手工补回')
        else:
            print(f'  成品正文图片比源多 {di["body"] - si["body"]} 个，请核对')
    if si['header_footer'] != di['header_footer']:
        ok = False
        if si['header_footer'] > di['header_footer']:
            print(f'  页眉页脚图片少 '
                  f'{si["header_footer"] - di["header_footer"]} 个：'
                  f'若源文档的版头在页眉里，需手工补回')
        else:
            print(f'  成品页眉页脚图片比源多 '
                  f'{di["header_footer"] - si["header_footer"]} 个，请核对')
    if si['other'] != di['other']:
        warns.append(f'脚注/尾注/批注中的图片：源 {si["other"]}、'
                     f'成品 {di["other"]}（本技能不处理这些内容）')
    if (si['body'], si['header_footer']) == (di['body'], di['header_footer']):
        print('  正文与页眉页脚图片数量一致')

    # ── 有意调整 ───────────────────────────────────────
    notes = []
    if meta:
        st = meta.get('_stripped') or []
        if st:
            notes.append('落款已剥离后由生成器重新输出：'
                         + '、'.join(f"{s['field']}「{s['text']}」" for s in st))
        an = meta.get('_auto_number') or []
        if an:
            notes.append(f'Word 自动编号 {len(an)} 处已补回为文字'
                         f'（源文件里不是字符）')
        if meta.get('header'):
            n = len(meta['header']['indices'])
            notes.append(f'版头 {n} 段整段原样复制；'
                         f'若含文本框，其中文字已按全文 w:t 参与比对，'
                         f'框体与位置仍需目视确认')
        for u in (meta.get('_unsupported') or []):
            notes.append(f"源文档含「{u['kind']}」{u['count']} 处：{u['hint']}")
    notes.extend(warns)
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

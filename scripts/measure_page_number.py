#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""页码位置实测：转 PDF → 渲染 → 扫描墨迹，量出真实位置。

    python3 scripts/measure_page_number.py 成品.docx

为什么要有这个脚本
  页脚距页底设多少，不等于页码落在哪里。字体、行距、字形上沿留白都会影响，
  靠推算会错（本技能早期按 4.58mm 设置，实测页码距版心下边缘 26mm，
  国标要求 7mm，偏差 19mm）。只有把成品渲染出来量，才知道对不对。

做法
  成品 → LibreOffice 转 PDF → 按 300dpi 渲染 → 在页脚区域扫描非白像素，
  取第一行墨迹的上沿，换算成「距版心下边缘多少毫米」。
  墨迹比文字 bbox 的上沿更贴近视觉位置，两者约差 0.8mm。

依赖
  LibreOffice（soffice）与 PyMuPDF。缺任何一个都会明确报错并退出，
  不会假装通过——量不出来就是没验，不能当合格。

判定
  GB/T 9704-2012 7.5：页码一字线上沿距版心下边缘 7mm。
  默认容差 0.5mm（--tolerance 可调）。
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

# 版式常量（GB/T 9704-2012）
PAGE_H_MM = 297.0
PAGE_W_MM = 210.0
MARGIN_BOTTOM_MM = 35.0
MARGIN_LEFT_MM = 28.0
MARGIN_RIGHT_MM = 26.0
TYPE_BOTTOM_MM = PAGE_H_MM - MARGIN_BOTTOM_MM      # 版心下边缘 262mm
TYPE_LEFT_MM = MARGIN_LEFT_MM
TYPE_RIGHT_MM = PAGE_W_MM - MARGIN_RIGHT_MM        # 184mm
TYPE_CENTER_MM = (TYPE_LEFT_MM + TYPE_RIGHT_MM) / 2
TARGET_MM = 7.0                                    # 国标要求
ONE_CHAR_MM = 16 * 25.4 / 72                       # 三号字「一字」≈ 5.64mm


def find_soffice():
    for cand in ('soffice', 'libreoffice',
                 '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                 '/opt/homebrew/bin/soffice', '/usr/bin/soffice'):
        p = shutil.which(cand) if os.path.basename(cand) == cand else cand
        if p and os.path.exists(p):
            return p
    return None


def to_pdf(docx_path, outdir):
    soffice = find_soffice()
    if soffice is None:
        raise SystemExit(
            '未找到 LibreOffice（soffice），无法渲染测量。\n'
            '  macOS: brew install --cask libreoffice\n'
            '  或指定路径后重试。量不出来就是没验，不能用估算代替。')
    # 独立 profile：默认配置被并发/强制退出破坏后会报 DeploymentException
    subprocess.run(
        [soffice, '-env:UserInstallation=file:///tmp/lo_profile_measure',
         '--headless', '--convert-to', 'pdf', '--outdir', outdir, docx_path],
        check=True, capture_output=True)
    pdf = os.path.join(outdir,
                       os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
    if not os.path.exists(pdf):
        raise SystemExit(f'转换失败，未生成 PDF：{pdf}')
    return pdf


def ink_rows(samples, stride, width, y0, y1, n):
    """在 y0..y1 行区间内找出有墨迹的行号（纯 Python，不需要 numpy）。"""
    rows = []
    for y in range(y0, min(y1, len(samples) // stride)):
        start = y * stride
        row = samples[start:start + width * n]
        if row and min(row) < 200:
            rows.append(y)
    return rows


def measure(pdf_path, dpi=300, bottom=MARGIN_BOTTOM_MM):
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz
        except ImportError:
            raise SystemExit(
                '缺少 PyMuPDF（pip install pymupdf），无法渲染测量。\n'
                '量不出来就是没验，不能用估算代替。')
    type_bottom = PAGE_H_MM - bottom
    doc = fitz.open(pdf_path)
    results = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        samples = pix.samples
        n, stride = pix.n, pix.stride
        width = pix.width
        px = dpi / 25.4                      # 每毫米像素数
        y0 = int(type_bottom * px)
        y1 = int(PAGE_H_MM * px)
        rows = ink_rows(samples, stride, width, y0, y1, n)
        if not rows:
            results.append({'page': i + 1, 'found': False})
            continue
        # 墨迹水平范围（只取页码所在的首行附近，避免正文残墨干扰）
        band0, band1 = rows[0], min(rows[0] + int(6 * px), y1)
        left_px, right_px = None, None
        for y in range(band0, band1):
            start = y * stride
            row = samples[start:start + width * n]
            for x in range(width):
                if min(row[x * n:(x + 1) * n]) < 200:
                    left_px = x if left_px is None else min(left_px, x)
                    right_px = x if right_px is None else max(right_px, x)
        top_mm = rows[0] / px
        rec = {
            'page': i + 1,
            'found': True,
            'distance_mm': round(top_mm - type_bottom, 2),
        }
        if left_px is not None:
            left_mm = left_px / px
            right_mm = (right_px + 1) / px
            center = (left_mm + right_mm) / 2
            rec['left_mm'] = round(left_mm, 2)
            rec['right_mm'] = round(right_mm, 2)
            rec['center_mm'] = round(center, 2)
            rec['h_style'] = ('center' if abs(center - TYPE_CENTER_MM) < 1.0
                              else ('left' if center < TYPE_CENTER_MM
                                    else 'right'))
            rec['gap_mm'] = (round(TYPE_RIGHT_MM - right_mm, 2)
                             if rec['h_style'] == 'right'
                             else round(left_mm - TYPE_LEFT_MM, 2)
                             if rec['h_style'] == 'left' else None)
        results.append(rec)
    doc.close()
    return results


def main():
    ap = argparse.ArgumentParser(
        description='实测成品页码位置（转 PDF 渲染后扫描墨迹）')
    ap.add_argument('docx', help='生成的 .docx 成品')
    ap.add_argument('--dpi', type=int, default=300, help='渲染分辨率，默认 300')
    ap.add_argument('--tolerance', type=float, default=0.5,
                    help='与 7mm 的容差（毫米），默认 0.5')
    ap.add_argument('--bottom', type=float, default=MARGIN_BOTTOM_MM,
                    help='下页边距（毫米），默认 35')
    ap.add_argument('--json', dest='json_out', help='把结果写成 JSON 文件')
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        print(f'文件不存在: {args.docx}', file=sys.stderr)
        return 1

    tmp = tempfile.mkdtemp(prefix='pn_measure_')
    try:
        pdf = to_pdf(args.docx, tmp)
        results = measure(pdf, dpi=args.dpi, bottom=args.bottom)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    found = [r for r in results if r.get('found')]
    print('──────── 页码位置实测 ────────')
    print(f'成品　：{os.path.basename(args.docx)}（{len(results)} 页）')
    print(f'标准　：一字线上沿距版心下边缘 {TARGET_MM:.2f}mm'
          f'（GB/T 9704-2012 7.5）')

    if not found:
        print('\n页脚区域未发现墨迹：该文档可能未添加页码'
              '（源文档无页码时本技能不添加）。')
        return 0

    print()
    for r in found:
        line = (f'  第{r["page"]:>2}页：距版心下边缘 {r["distance_mm"]:5.2f}mm'
                f'（偏差 {r["distance_mm"] - TARGET_MM:+.2f}）')
        if r.get('h_style') == 'center':
            line += (f'　居中偏差 {r["center_mm"] - TYPE_CENTER_MM:+.2f}mm')
        elif r.get('gap_mm') is not None:
            side = '右空' if r['h_style'] == 'right' else '左空'
            line += (f'　{side} {r["gap_mm"]:5.2f}mm'
                     f'（一字 {ONE_CHAR_MM:.2f}mm，'
                     f'偏差 {r["gap_mm"] - ONE_CHAR_MM:+.2f}）')
        print(line)

    worst = max(abs(r['distance_mm'] - TARGET_MM) for r in found)
    print(f'\n  共 {len(found)} 页，与 {TARGET_MM:.0f}mm 的最大偏差 '
          f'{worst:.2f}mm（容差 {args.tolerance:.2f}mm）')
    ok = worst <= args.tolerance
    print('  结论：' + ('页码位置合格' if ok
                    else '页码位置超差，请核查页脚距与字体字号'))

    if args.json_out:
        json.dump({'target_mm': TARGET_MM, 'tolerance_mm': args.tolerance,
                   'worst_deviation_mm': round(worst, 2), 'ok': ok,
                   'pages': results},
                  open(args.json_out, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=2)
        print(f'  结果已写入 {args.json_out}')
    return 0 if ok else 2


if __name__ == '__main__':
    sys.exit(main())

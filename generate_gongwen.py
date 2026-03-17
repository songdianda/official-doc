#!/usr/bin/env python3
"""根据 official-doc skill 格式规范生成公文 DOCX"""

import os

from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── 辅助函数 ──────────────────────────────────────────────

def set_font(run, font_name, font_size, bold=False):
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
    pPr = paragraph._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(int(pt_value * 20)))
    spacing.set(qn('w:lineRule'), 'exact')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')


def set_first_line_indent(paragraph, chars=2):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(chars * 100))


def add_body_paragraph(doc, text, font_name='仿宋_GB2312', font_size=16,
                       bold=False, indent=True, line_spacing=28):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_line_spacing(para, line_spacing)
    if indent:
        set_first_line_indent(para, 2)
    run = para.add_run(text)
    set_font(run, font_name, font_size, bold)
    return para


def add_level1(doc, text):
    """第一层：一、 → 黑体，不加粗"""
    return add_body_paragraph(doc, text, font_name='黑体', bold=False)


def add_level2(doc, text):
    """第二层：（一） → 楷体_GB2312，不加粗"""
    return add_body_paragraph(doc, text, font_name='楷体_GB2312', bold=False)


def add_blank(doc, line_spacing=28):
    p = doc.add_paragraph()
    set_line_spacing(p, line_spacing)
    return p


def add_page_number(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(footer_para, 14)

    run1 = footer_para.add_run('— ')
    set_font(run1, '宋体', 14)

    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run_field = footer_para.add_run()
    run_field._element.append(fldChar1)

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    run_instr = footer_para.add_run()
    set_font(run_instr, '宋体', 14)
    run_instr._element.append(instrText)

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run_end = footer_para.add_run()
    run_end._element.append(fldChar2)

    run3 = footer_para.add_run(' —')
    set_font(run3, '宋体', 14)


# ── 生成公文 ──────────────────────────────────────────────

def generate():
    doc = Document()

    # ── 页面设置 ──
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(37)
    section.bottom_margin = Mm(35)
    section.left_margin = Mm(28)
    section.right_margin = Mm(26)

    # ── 标题（方正小标宋简体，二号22磅，居中，行距32磅） ──
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(title_para, 32)
    title_run = title_para.add_run(
        '关于举办"汉东省高端装备制造\n产业招商推介会"的请示'
    )
    set_font(title_run, '方正小标宋简体', 22)

    # ── 标题后空一行 ──
    add_blank(doc, 28)

    # ── 主送机关（仿宋_GB2312，三号16磅，顶格，全角冒号） ──
    recipient_para = doc.add_paragraph()
    recipient_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_line_spacing(recipient_para, 28)
    recipient_run = recipient_para.add_run('省人民政府：')
    set_font(recipient_run, '仿宋_GB2312', 16)

    # ── 正文（仿宋_GB2312，三号16磅，首行缩进二字，行距28磅） ──

    add_body_paragraph(
        doc,
        '为深入贯彻落实省委、省政府关于"双招双引"工作的决策部署，'
        '助推我省高端装备制造产业高质量发展，经前期与北京市相关行业协会'
        '及多家央企沟通，我办拟于近期在京举办"汉东省（北京）高端装备制造'
        '产业招商推介会"。现将有关事项请示如下：'
    )

    # ── 一、活动概况 ──
    add_level1(doc, '一、活动概况')

    add_level2(doc, '（一）时间地点')
    add_body_paragraph(
        doc,
        '拟定于2026年4月1日（星期四）下午14:30，在北京汉东酒店举行。'
    )

    add_level2(doc, '（二）活动主题')
    add_body_paragraph(doc, '深化京汉合作·共创制造未来。')

    add_level2(doc, '（三）参会规模')
    add_body_paragraph(
        doc,
        '拟邀请国家部委相关司局领导、北京市有关部门负责人、'
        '央企及知名民企代表约150人。'
    )

    # ── 二、主要议程 ──
    add_level1(doc, '二、主要议程')

    add_level2(doc, '（一）省领导致辞；')
    add_level2(doc, '（二）推介会')
    add_body_paragraph(doc, '我省产业政策及重点项目推介。')
    add_level2(doc, '（三）签约仪式')
    add_body_paragraph(
        doc,
        '重点项目集中签约仪式（拟签约项目3个，总投资额约2亿元）。'
    )
    add_level2(doc, '（四）政企座谈交流。')

    # ── 三、需请省政府支持事项 ──
    add_level1(doc, '三、需请省政府支持事项')

    add_level2(doc, '（一）恳请省领导出席')
    add_body_paragraph(
        doc,
        '拟恳请省政府分管工业的副省长出席活动并致辞，'
        '以提升活动规格，增强客商信心。'
    )

    add_level2(doc, '（二）协调部门参会')
    add_body_paragraph(
        doc,
        '恳请省政府办公厅协调省发改委、省工信厅、省商务厅各派一名负责'
        '同志随团参会，现场解答政策咨询。'
    )

    add_level2(doc, '（三）经费保障')
    add_body_paragraph(
        doc,
        '本次活动预计经费15万元，建议从我省年度招商引资专项经费中列支。'
    )

    # ── 四、下一步工作 ──
    add_level1(doc, '四、下一步工作')

    add_body_paragraph(
        doc,
        '如蒙获准，我办将立即成立筹备工作组，细化活动方案，严密组织会务，'
        '并加强与省直相关部门的对接，确保活动取得实效。'
    )

    add_body_paragraph(doc, '以上请示，妥否，请批示。')

    # ── 落款 ──
    add_blank(doc, 28)
    add_blank(doc, 28)

    unit_para = doc.add_paragraph()
    unit_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_line_spacing(unit_para, 28)
    unit_run = unit_para.add_run('汉东省人民政府驻北京办事处')
    set_font(unit_run, '仿宋_GB2312', 16)

    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_line_spacing(date_para, 28)
    date_run = date_para.add_run('2026年3月17日')
    set_font(date_run, '仿宋_GB2312', 16)

    # ── 页码 ──
    add_page_number(section)

    # ── 保存 ──
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    output_path = os.path.join(desktop, '公文模板-汉东-正式版.docx')
    doc.save(output_path)
    print(f'公文已生成: {output_path}')


if __name__ == '__main__':
    generate()

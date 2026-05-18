---
name: official-doc
description: 根据党政机关公文格式规范（GB/T 9704-2012），使用 python-docx 生成标准公文 DOCX 文件。支持从文本或 .docx 文件输入，自动检查编号规范，按国标排版输出。
user_invocable: true
---

# official-doc

党政机关公文格式排版 Skill，基于国家标准 GB/T 9704-2012，使用 python-docx 生成标准公文 DOCX 文件。

---

## 概述

| 项目 | 说明 |
|------|------|
| **名称** | official-doc |
| **用途** | 生成符合党政机关公文格式国家标准的 .docx 文件 |
| **依赖** | python-docx、方正小标宋简体、仿宋_GB2312、楷体_GB2312 |
| **输入** | 文本内容或 .docx 文件路径 |
| **输出** | 排版后的 .docx 公文文件 |

---

## 工作流程

### 输入方式

**方式一：提供 .docx 文件路径**

先用 python-docx 读取并提取文本内容，展示给用户确认，再按规范重新排版。

```python
from docx import Document

def read_docx(file_path):
    doc = Document(file_path)
    content = []
    for para in doc.paragraphs:
        if para.text.strip():
            content.append(para.text.strip())
    return '\n'.join(content)
```

**方式二：直接提供文本内容**

用户直接给出标题、主送机关、正文等信息。

### 生成步骤

1. 读取/提取文本内容，展示给用户确认
2. 检查编号规范（见下方规则），列出错误并告知修正方案
3. 确认：标题、主送机关、正文、附件、输出文件名
4. 编写 Python 脚本生成公文
5. 执行脚本输出 .docx 文件

---

## 编号检查规则

### 层级格式对照表

| 层级 | 合法格式 | 常见错误 |
|------|---------|---------|
| 第一层 | `一、` `二、` `三、` | `1.` `（一）` `第一,` |
| 第二层 | `（一）` `（二）` `（三）` | `1、` `1.` `(1)` `（1）` `A.` |
| 第三层 | `1.` `2.` `3.` | `1、` `1）` `（1）` `①` |
| 第四层 | `（1）` `（2）` `（3）` | `1)` `(1)` `①` `a.` |
| 第五层 | `①` `②` `③` | `a.` `A)` |

### 检查要点

1. **层级格式必须严格对应**，不得跨层使用编号
2. **层级嵌套必须连续**，第一层下必须是第二层，不能跳层
3. **同层编号必须连续且格式统一**，不得跳号或混用格式
4. **顿号"、"不得用于数字编号**，`1、` 应为 `1.`
5. **括号必须使用全角中文括号**，`(1)` 应为 `（1）`

### 输出格式

```
发现以下编号问题：

| 段落位置 | 原文编号 | 问题说明 | 修正为 |
|---------|---------|---------|--------|
| 第X段   | 1、     | 第二层应使用（一）格式 | （一） |
```

---

## 格式规范

### 字体文件

| 字体 | 文件路径 |
|------|---------|
| 方正小标宋简体 | `font/方正小标宋简体.TTF` |
| 仿宋_GB2312 | `font/仿宋_GB2312.ttf` |
| 楷体_GB2312 | `font/楷体_GB2312.ttf` |

> python-docx 只能设置字体名称，无法嵌入字体。需通过 XML 操作设置 `w:eastAsia` 属性确保中文字体生效。

### 页面设置

| 项目 | 值 |
|------|-----|
| 纸张 | A4（210mm × 297mm） |
| 上边距 | 37mm |
| 下边距 | 35mm |
| 左边距 | 28mm |
| 右边距 | 26mm |
| 版心尺寸 | 156mm × 225mm |

### 各元素格式

#### 标题

| 属性 | 值 |
|------|-----|
| 字体 | 方正小标宋简体 |
| 字号 | 二号（22磅） |
| 对齐 | 居中 |
| 行距 | 固定值 32 磅 |
| 段前段后 | 0 |

- 由"发文单位+关于+事由+文种"组成
- 回行时词意完整，排列为梯形或菱形

#### 主送机关

| 属性 | 值 |
|------|-----|
| 字体 | 仿宋_GB2312 |
| 字号 | 三号（16磅） |
| 对齐 | 居左顶格 |
| 行距 | 固定值 28 磅 |

- 编排于标题下空一行
- 最后一个机关名称后标全角冒号

#### 正文

| 属性 | 值 |
|------|-----|
| 字体 | 仿宋_GB2312 |
| 字号 | 三号（16磅） |
| 缩进 | 首行缩进 2 字符 |
| 行距 | 固定值 28 磅 |
| 段前段后 | 0 |

结构层次字体：

| 层级 | 编号格式 | 字体 | 加粗 |
|------|---------|------|------|
| 第一层 | `一、` | 黑体 | 否 |
| 第二层 | `（一）` | 楷体_GB2312 | 否 |
| 第三层 | `1.` | 仿宋_GB2312 | 是 |
| 第四层 | `（1）` | 仿宋_GB2312 | 是 |
| 第五层 | `①` | 仿宋_GB2312 | 是 |

- 第一、二层标题末尾不加标点，单独成段

#### 附件说明

- 正文下空一行，左空二字符编排"附件"二字
- 后标全角冒号和附件名称
- 多个附件用阿拉伯数字标注顺序号
- 附件名称后不加标点符号
- 回行时与上一行附件名称首字对齐

#### 页码

| 属性 | 值 |
|------|-----|
| 字体 | 宋体 |
| 字号 | 四号（14磅） |
| 格式 | 半角阿拉伯数字，左右各一条一字线（`— 1 —`） |
| 位置 | 版心下边缘之下 4.58mm |
| 对齐 | 单页码右对齐，双页码左对齐 |

---

## Python 实现

```python
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# === 页面设置 ===
section = doc.sections[0]
section.page_width = Mm(210)
section.page_height = Mm(297)
section.top_margin = Mm(37)
section.bottom_margin = Mm(35)
section.left_margin = Mm(28)
section.right_margin = Mm(26)

# === 设置中文字体 ===
def set_font(run, font_name, font_size, bold=False):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

# === 设置固定行距 ===
def set_line_spacing_fixed(paragraph, pt_value):
    pPr = paragraph._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(int(pt_value * 20)))  # 磅转twip
    spacing.set(qn('w:lineRule'), 'exact')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')

# === 设置首行缩进（2字符）===
# 注意：python-docx 的 Normal 样式自带 firstLineChars=200，
# 必须先清除所有已有缩进属性，再同时设置 firstLineChars 和 firstLine，
# 否则会与默认值叠加导致缩进翻倍。
def set_paragraph_indent(paragraph, chars=2):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    # 清除所有已有的缩进属性，避免与 Normal 样式叠加
    for attr in ['w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars',
                 'w:hanging', 'w:hangingChars', 'w:start', 'w:end']:
        if ind.get(qn(attr)):
            del ind.attrib[qn(attr)]
    # 同时设置字符单位和绝对值（Word 以字符单位优先）
    ind.set(qn('w:firstLineChars'), str(chars * 100))
    ind.set(qn('w:firstLine'), str(chars * 320))

# === 使用示例 ===
# 标题：行距 32 磅
set_line_spacing_fixed(title_paragraph, 32)

# 正文/主送机关：行距 28 磅，首行缩进 2 字符
set_line_spacing_fixed(body_paragraph, 28)
set_paragraph_indent(body_paragraph, 2)
```

---

## 输出要求

- 输出文件格式为 `.docx`
- 默认输出到当前工作目录
- 文件名由用户指定，默认使用公文标题作为文件名

---

## 参考标准

- 《党政机关公文处理工作条例》（中办发〔2012〕14号）
- 《党政机关公文格式》（GB/T 9704-2012）

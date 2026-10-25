---
name: official-doc
display_name: 公文格式排版
display_name_en: Official Document Formatter
description: 依据 GB/T 9704-2012 党政机关公文格式国标生成规范排版的 DOCX 公文，支持从文本或现有 docx 输入，并自动检查结构层次编号错误
description_zh: 按党政机关公文格式国标(GB/T 9704-2012)把内容重排为标准 DOCX 公文，自动检查并修正"一、（一）1.（1）①"层级编号错误
description_en: Format Chinese official documents into standard DOCX per GB/T 9704-2012, with automatic structural numbering checks
examples_zh:
  - "把这份会议纪要按公文格式排版成 Word"
  - "帮我检查这份公文的编号层级有没有错"
  - "把这段文字排成一份标准请示，主送机关是XX厅"
examples_en:
  - "Format this meeting minutes into a standard official Word document"
  - "Check whether the numbering hierarchy in this document is correct"
  - "Turn this text into a standard request document addressed to XX Department"
version: 1.1.0
author: songdianda
---

# 公文格式排版（GB/T 9704-2012）

把公文内容重排为符合《党政机关公文格式》国家标准的 `.docx` 文件。

## 何时使用

用户要求：生成/排版/重排公文、通知、请示、报告、函、纪要；检查公文编号层级；把 Word 或文本改成标准公文格式。

## 工作流程

1. **获取内容**：用户提供 `.docx` 路径时，先用 python-docx 提取全文并展示确认；直接给文本时跳过此步。
2. **编号检查**：逐段扫描层级编号，对照 `references/numbering-rules.md`。有问题则用表格列出「段落位置 / 原文编号 / 问题 / 修正为」，并征询用户是否自动修正。**不擅自改动用户原意。**
3. **确认要素**：标题、主送机关、正文、附件、发文机关署名、成文日期、输出路径。用户未提供的项目逐项询问，**不编造**。
4. **生成文件**：调用 `scripts/generate_gongwen.py`，格式参数见 `references/format-spec.md`。
5. **交付**：告知输出路径，并提示字体是否安装（未装则格式会走样）。

## 调用脚本

```bash
# 准备依赖（首次）
pip install python-docx

# 方式一：JSON 配置（推荐，支持附件/署名/日期）
python3 scripts/generate_gongwen.py --config content.json --output 公文.docx

# 方式二：纯文本正文（每行一段，自动识别层级）
python3 scripts/generate_gongwen.py --title "关于XXX的请示" --recipient "XX厅" \
    --body-file body.txt --signature "XXX办公室" --date "2026年3月17日" --output 公文.docx

# 检查运行环境字体
python3 scripts/generate_gongwen.py --check-fonts
```

`content.json` 结构：

```json
{
  "title": "关于XXX的请示",
  "recipient": "XX厅",
  "body": ["正文段落（自动识别层级）", "一、第一层", "（一）第二层", "1. 第三层"],
  "attachments": ["附件名称"],
  "signature": "发文机关署名",
  "date": "2026年3月17日",
  "output": "公文.docx"
}
```

`body` 元素也可写作 `{"type": "level1"|"level2"|"level3"|"level4"|"level5"|"text"|"blank", "text": "..."}` 显式指定层级。

## 关键格式参数

详见 `references/format-spec.md`，速查：

| 元素 | 字体 | 字号 | 行距 |
|------|------|------|------|
| 标题 | 方正小标宋简体 | 二号 22磅 | 32磅 |
| 主送机关 | 仿宋_GB2312 | 三号 16磅 | 28磅 |
| 正文 | 仿宋_GB2312 | 三号 16磅 | 28磅 |
| 第一层 `一、` | 黑体 | 三号 | 28磅 |
| 第二层 `（一）` | 楷体_GB2312 | 三号 | 28磅 |
| 三/四/五层 | 仿宋_GB2312 加粗 | 三号 | 28磅 |
| 页码 | 宋体 | 四号 14磅 | 单页右空一字 / 双页左空一字 |

页面：A4，上 37mm、下 35mm、左 28mm、右 26mm。正文每自然段首行缩进 2 字符。

## 约束

- **不编造内容**：单位名称、文号、日期、金额、人名等一律向用户确认，缺失就问，不填充占位式臆测数据。
- **不擅自改意**：编号修正只改编号格式，不改文字表述；拿不准的先问。
- **字体不打包，但会自动回退**：本技能不含字体文件（字体受著作权保护，商用需自行取得授权）。脚本内置回退链，本机缺少国标字体时自动选用等效字体并打印替换情况（`方正小标宋简体→宋体-简`、`仿宋_GB2312→方正仿宋_GBK` 等）。生成后须把实际使用的字体告知用户；若标题被回退到普通宋体，应提示"标题非小标宋，装在装了国标字体的机器上效果更好"。
- **只输出 .docx**：不生成 PDF、不生成图片版式。
- **输出位置**：默认当前工作目录，文件名由用户指定。

## 参考

- `references/format-spec.md` —— 完整格式规范（含国标条款号）
- `references/numbering-rules.md` —— 编号层级检查规则
- 《党政机关公文格式》GB/T 9704-2012
- 《党政机关公文处理工作条例》（中办发〔2012〕14号）

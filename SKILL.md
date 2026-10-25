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
version: 1.3.0
author: songdianda
---

# 公文格式排版（GB/T 9704-2012）

把公文内容重排为符合《党政机关公文格式》国家标准的 `.docx` 文件。

## 何时使用

用户要求：生成/排版/重排公文、通知、请示、报告、函、纪要；检查公文编号层级；把 Word 或文本改成标准公文格式。

## 工作流程

1. **获取内容**：用户提供 `.docx` 路径时，用 `scripts/inspect_docx.py` 提取全文、样式线索与角色建议；直接给文本时跳过此步。
2. **判定段落角色**（关键）：逐段确定是标题还是编号正文，规则见 `references/paragraph-roles.md`。
   - 「1.」既可能是三级标题，也可能只是一条编号的工作事项，**不能只看编号**。
   - 依据优先级：Word 的 `outlineLvl`/标题样式 > 上下文结构 > 是否成句 > 原文加粗（**加粗仅作弱参考，原格式可能有误**）。
   - **不用字数阈值**判断。
   - 拿不准时默认按 `numbered_body`（保留编号、不加粗），并把疑问集中列表请用户确认。
   - **不自动改动编号**；编号错误只在用户要求时按 `numbering-rules.md` 修正并单独报告。
3. **编号检查**：对照 `references/numbering-rules.md` 检查跳层、跳号、半角括号等。有问题用表格列出「段落位置 / 原文编号 / 问题 / 修正为」，征询后再改。**不擅自改动用户原意。**
4. **确认要素**：标题、主送机关、正文、附件、发文机关署名、成文日期、输出路径。用户未提供的项目逐项询问，**不编造**。
5. **生成文件**：调用 `scripts/generate_gongwen.py`，格式参数见 `references/format-spec.md`。
6. **交付**：告知输出路径、实际使用的字体（字体回退情况），以及存疑段落的处理结果。

## 调用脚本

```bash
# 准备依赖（首次）
pip install python-docx

# 推荐流程：提取 → 定角色 → 渲染
python3 scripts/inspect_docx.py 原文.docx -o content.json   # 提取 + 角色建议 + 存疑清单
# （审阅 content.json 的 _uncertain，定稿 type）
python3 scripts/generate_gongwen.py --config content.json --output 公文.docx

# 直接给文本（无源文件）
python3 scripts/generate_gongwen.py --title "关于XXX的请示" --recipient "XX厅" \
    --body-file body.txt --signature "XXX办公室" --date "2026年3月17日" --output 公文.docx

# 检查运行环境字体
python3 scripts/generate_gongwen.py --check-fonts
```

`content.json` 结构（**角色必须显式指定**）：

```json
{
  "title": "关于XXX的请示",
  "recipient": "XX厅",
  "body": [
    {"type": "body", "text": "总起段落……"},
    {"type": "heading1", "text": "一、第一层标题"},
    {"type": "heading2", "text": "（一）第二层标题"},
    {"type": "numbered_body", "text": "1. 完成了……，开展了……。"},
    {"type": "table", "index": 0},
    {"type": "blank", "text": ""}
  ],
  "tables": [{"rows": [["列1", "列2"], ["值1", "值2"]], "header": true}],
  "attachments": ["附件名称"],
  "signature": "发文机关署名",
  "date": "2026年3月17日",
  "output": "公文.docx"
}
```

`type` 取值：`heading1`~`heading5`、`body`、`numbered_body`、`blank`、`table`。

**附表**：`inspect_docx.py` 会把源文档表格提取到 `tables`，并在 `body` 对应位置插 `{"type":"table","index":0}`。
排版为居中 + 全线边框 + 表头 12 磅加粗 + 数据 10.5 磅。**表格数据必须与原文逐格一致，不得改动数值。**
其中 **`numbered_body`** 表示"带编号的工作事项"——保留编号、按正文排版、**不加粗**。
工作总结/汇报中「1. 完成了……」这类罗列几乎都是它，误判成 `heading3` 会导致整段被加粗。

纯字符串元素（不写 `type`）时脚本按保守策略兜底：`一、`/`（一）` 判为标题，`1.`/`（1）`/`①` 一律按编号正文（不加粗）。

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
- `references/paragraph-roles.md` —— 段落角色判定规则（标题 vs 编号正文）
- `references/numbering-rules.md` —— 编号层级检查规则
- 《党政机关公文格式》GB/T 9704-2012
- 《党政机关公文处理工作条例》（中办发〔2012〕14号）

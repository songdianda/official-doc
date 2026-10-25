---
name: official-doc
display_name: 公文格式排版
display_name_en: Official Document Formatter
description: 依据 GB/T 9704-2012 党政机关公文格式国标，把内容或现有 docx 重排为标准公文，源文件带红头与发文字号时原样保留、只重排正文，保留编号与表格，自动检查结构层次编号错误。触发词：公文排版、公文格式、重排公文、红头、发文字号、编号层级检查、GB/T 9704
description_zh: 按党政机关公文格式国标(GB/T 9704-2012)把内容或现有 Word 重排为标准 DOCX 公文；源文件带红头、发文字号时原样保留、只重排正文，保留编号与表格（含合并单元格），自动检查并修正"一、（一）1.（1）①"层级编号错误。触发词：公文排版、公文格式、重排公文、红头、发文字号、编号层级检查、GB/T 9704
description_en: "Reformat content or an existing Word file into a standard Chinese official DOCX per GB/T 9704-2012; keeps existing letterhead and document number intact, preserving numbering and tables, with automatic structural numbering checks. Triggers - official document formatting, GB/T 9704, letterhead, document number, heading hierarchy check"
examples_zh:
  - "把这份会议纪要按公文格式排版成 Word"
  - "帮我检查这份公文的编号层级有没有错"
  - "把这段文字排成一份标准请示，主送机关是XX厅"
  - "把这份带附表的请示重排成标准格式，页码跟原文保持一致"
examples_en:
  - "Format this meeting minutes into a standard official Word document"
  - "Check whether the numbering hierarchy in this document is correct"
  - "Turn this text into a standard request document addressed to XX Department"
  - "Reformat this request with an attached table, keeping the original page numbering"
version: 1.9.1
author: songdianda
user-invocable: true
---

# 公文格式排版（GB/T 9704-2012）

把公文内容重排为符合《党政机关公文格式》国家标准的 `.docx` 文件。

## 何时使用

用户要求：生成/排版/重排公文、通知、请示、报告、函、纪要；检查公文编号层级；把 Word 或文本改成标准公文格式。

## 工作流程

1. **获取内容**：用户提供 `.docx` 路径时，用 `scripts/inspect_docx.py` 提取全文、样式线索与角色建议；脚本会**自动检测源文档是否已有页码**，**识别并剥离版头**（红头/发文字号/分隔线，原样保留不重排），**把落款署名与成文日期剥离到 `signature` / `date`**（生成器会重新输出，留在正文里会重复），并保留表格的合并单元格结构。
   - **版头若不识别，红头会被当成标题排成二号小标宋，真正的标题反而沦为普通正文**——这一步不能跳。
   - 主送机关、附件说明**误判代价高**（删了就丢内容），脚本只列候选 `_candidates`，由你确认后填入对应字段并删除 body 中的该段。
2. **判定段落角色**（关键）：逐段确定是标题还是编号正文，规则见 `references/paragraph-roles.md`。
   - 「1.」既可能是三级标题，也可能只是一条编号的工作事项，**不能只看编号**。
   - 依据优先级：Word 的 `outlineLvl`/标题样式 > 上下文结构 > 是否成句 > 原文加粗（**加粗仅作弱参考，原格式可能有误**）。
   - **不用字数阈值**判断。
   - 拿不准时默认按 `numbered_body`（保留编号、不加粗），并把疑问集中列表请用户确认。
   - **不自动改动编号**；编号错误只在用户要求时按 `references/numbering-rules.md` 修正并单独报告。
3. **编号检查**：对照 `references/numbering-rules.md` 检查跳层、跳号、半角括号等。有问题用表格列出「段落位置 / 原文编号 / 问题 / 修正为」，征询后再改。**不擅自改动用户原意。**
   - **先确认自动编号**：Word 自动编号不在 `p.text` 里，`scripts/inspect_docx.py` 会从 `numbering.xml` 算出并补回（结果见 `_auto_number`）。若发现源文档混用了手打编号与自动编号，交付时明确说明——这是最容易"看起来少了一级"的坑。
4. **确认要素**：标题、主送机关、正文、附件、发文机关署名、成文日期、输出路径。用户未提供的项目逐项询问，**不编造**。
   - **页码不在此列**：有源文件时以源文件为准（无页码就不加），无源文件时默认居中，交付时说明采用的值。
5. **生成文件**：调用 `scripts/generate_gongwen.py`，格式参数见 `references/format-spec.md`。
6. **交付**：告知输出路径、实际使用的字体（字体回退情况）、存疑段落的处理结果、
   **版头是否保留**（识别到几段、保留的是什么），
   **并主动说明本技能不生成的要素**（版记、印章，以及源文件无版头时不生成红头），
   避免用户误以为拿到的是完整公文。完整说明见 `README.md`。

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
  "page_number": "center",
  "output": "公文.docx"
}
```

`page_number` 控制页码样式，**由使用者选择**：

| 取值 | 效果 |
|------|------|
| `center` | 页码**居中** |
| `odd_even` | 单页居右空一字、双页居左空一字（GB/T 9704-2012 7.5） |
| `none` | 不显示页码 |

### 页码判定规则（重要）

**源文档没有页码，就不要替它加。** 排版是重排格式，不是替用户决定要不要页码——单页文件、内部传阅件常刻意不编页码，凭空加出来是越权。

- 用 `scripts/inspect_docx.py` 处理现有 docx 时，脚本**自动检测**并写入 `page_number`：
  - 源文档无页码 → `none`
  - 有页码 → 沿用其原样式（居中→`center`；奇偶页不同→`odd_even`）
- 由纯文本新建公文（无源文件可参照）时**默认 `center` 居中**，交付时说明。
- `--page-number center|odd_even|none` 可覆盖上述推断。

居中 vs 奇偶页：7.5 的左右错位是为**双面装订**设计的——翻开书时页码靠外侧才好看。单面打印或电子传阅时页码忽左忽右反而凌乱。

`type` 取值：`heading1`~`heading5`、`body`、`numbered_body`、`blank`、`table`。

**附表**：`scripts/inspect_docx.py` 会把源文档表格提取到 `tables`，并在 `body` 对应位置插 `{"type":"table","index":0}`。
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
| 页码 | 宋体 | 四号 14磅 | **无源文件时**默认居中；有源文件以源文件为准；可选 `odd_even` / `none` |

页面：A4，上 37mm、下 35mm、左 28mm、右 26mm。正文每自然段首行缩进 2 字符。

## 约束

- **不编造内容**：单位名称、文号、日期、金额、人名等一律向用户确认，缺失就问，不填充占位式臆测数据。
- **不擅自改意**：编号修正只改编号格式，不改文字表述；拿不准的先问。
- **字体不打包，但会自动回退**：本技能不含字体文件（字体受著作权保护，商用需自行取得授权）。脚本内置回退链，本机缺少国标字体时自动选用等效字体并打印替换情况（`方正小标宋简体→宋体-简`、`仿宋_GB2312→方正仿宋_GBK` 等）。生成后须把实际使用的字体告知用户；若标题被回退到普通宋体，应提示"标题非小标宋，在装了国标字体的机器上效果更好"。
- **只输出 .docx**：不生成 PDF、不生成图片版式。
- **输出位置**：默认当前工作目录，文件名由用户指定。

## 不适用与能力边界

### 版头：原样保留，不重做

**源文件带红头和发文字号时，一律原样保留，不参与重排。** 红头是各单位的门面，替用户重排等于改他的门面。

`scripts/inspect_docx.py` 会自动识别文档开头的红头、发文字号、红色分隔线（文字或图片都可），
记录到 `header`，由生成器**整段原样复制**到新文档开头——字体、颜色、字号、对齐、边框全保留。

识别规则（详见 `references/format-spec.md` 第 8 节）：

- 只扫描**开头连续**的那几段：正文里引述「X教发〔2026〕15号」不会被当成版头删掉
- 红色文字、字号 ≥ 二号、以「文件」结尾、符合发文字号格式——打分判定
- 无文字但带段落边框（分隔线）或含图片（图片版红头）同样纳入
- 识别到发文字号之后只再接受分隔线，避免把正文标题误纳入

**识别不出、或源文件根本没有版头时，不凭空生成一个**——那需要套用单位红头模板。

**版头字体：连文档默认字体一起保留。** 红头段落若没有逐字设字体，它显示的是
源文档的**文档级默认字体**（`docDefaults`）。生成器会把源文档的默认字体同步到新文档，
否则这些段落会掉回新模板的主题字体，表现为"红头换了张脸"。
正文、页码、表格均由脚本显式设字体，不受此项影响。

版头**不做字体回退**：本机缺字体时保留原字体名由 Word 替换，只在输出里提示缺哪个字体。
这与正文策略不同（正文缺国标字体时回退到本机等效字体），
所以同一份文档里红头与正文可能不是同一张字体——这是有意为之，不是缺陷。

### 不生成的要素

以下要素本技能不生成，交付时应向用户说明，避免被当成"完整公文"：

| 不生成 | 说明 |
|--------|------|
| 红头、发文字号、分隔线 | **源文件没有时不生成**（有则原样保留，见上） |
| 签发人、份号、密级、紧急程度 | 国标 7.2 版头要素，本技能不处理 |
| 版记（抄送机关、印发机关和印发日期） | 国标 7.4，需按单位版记格式另行排版 |
| 印章位置、骑缝章 | 与用印流程绑定，不在排版范围 |

其他边界：

- **仅适用于党政机关公文**：合同、协议、论文、商务文书、宣传材料不适用本格式。
- **输入仅限 `.docx`**：`.doc`、PDF、扫描件需先转成可编辑的 `.docx`；本技能不做格式转换。
- **只输出 `.docx`**：不生成 PDF，不生成图片版式。

## 参考

- `README.md` —— **面向使用者的说明**：能做什么 / 不能做什么 / 怎么用 / 常见问题
- `references/format-spec.md` —— 完整格式规范（含国标条款号）
- `references/paragraph-roles.md` —— 段落角色判定规则（标题 vs 编号正文）
- `references/numbering-rules.md` —— 编号层级检查规则
- 《党政机关公文格式》GB/T 9704-2012
- 《党政机关公文处理工作条例》（中办发〔2012〕14号）

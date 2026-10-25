# official-doc

按《党政机关公文格式》国家标准 **GB/T 9704-2012** 把内容排版成规范的 `.docx` 公文。

## 功能

- 标准排版：标题、主送机关、正文、附件说明、发文机关署名、成文日期、页码（奇偶页不同）
- **角色驱动排版**：先定段落角色再渲染，脚本不靠编号猜语义，避免整段误加粗
- 编号检查：自动识别并校验「一、 →（一）→ 1. →（1）→ ①」五层结构
- 双输入：JSON 配置，或纯文本正文
- 零硬编码：标题、主送机关、正文、附件、署名、日期、输出路径全部可传参
- **字体自动回退**：本机缺少国标字体时，自动选用等效字体，不再让 Word 静默替换

## 安装

```bash
pip install python-docx
```

**字体需自备**（本仓库不含字体文件，字体受著作权保护，商用请自行取得授权）。

国标字体名在各系统上并不一致：Windows 常见「仿宋_GB2312」，macOS 通常只有「方正仿宋_GBK」「华文仿宋」，Linux 多用「FangSong」。若写死国标名而本机没有，Word 会静默回退到默认字体，格式即走样。

脚本内置**字体回退链**，按「国标名 → 等效字体 → 通用兜底」逐项匹配，取本机第一个可用的写入文档，并在生成时打印替换情况：

| 用途 | 国标名 | 回退链（依次匹配） |
|------|--------|------------------|
| 标题 | 方正小标宋简体 | 小标宋体 → 方正小标宋_GBK → 华文中宋 → STZhongsong → 宋体-简 → 宋体 |
| 正文、主送机关 | 仿宋_GB2312 | 方正仿宋_GBK → 方正仿宋 → 华文仿宋 → 仿宋 → FangSong → 宋体 |
| 第一层 | 黑体 | 黑体-简 → 华文黑体 → Hei → Heiti SC → SimHei |
| 第二层 | 楷体_GB2312 | 楷体-简 → 华文楷体 → 楷体 → Kai → Kaiti SC → 宋体 |
| 页码 | 宋体 | 宋体-简 → 华文宋体 → SimSun → Songti SC |

自检（显示每个用途实际会用到哪个字体）：

```bash
python3 scripts/generate_gongwen.py --check-fonts
```

```
↷ 正文/主送机关：实际使用「方正仿宋_GBK」
    国标名「仿宋_GB2312」本机没有，已回退
✓ 页码：实际使用「宋体」
```

在装有国标字体的机器上不会发生任何替换。

## 段落角色（重要）

`"1."` 既可能是三级标题，也可能只是一条编号的工作事项。脚本**不猜语义**，
由调用方在 `type` 字段显式指定角色：

| type | 含义 | 加粗 |
|------|------|------|
| `heading1` ~ `heading5` | 一、/（一）/1./（1）/① 作为**标题** | 三~五层加粗 |
| `body` | 普通正文（无编号） | 否 |
| `numbered_body` | **编号正文**：带编号的工作事项，如「1. 完成了……」 | **否** |
| `blank` | 空行 | — |

工作总结、情况汇报里「1. 完成了……，开展了……」这类罗列几乎都是 `numbered_body`。
误判成 `heading3` 会导致整段被加粗——**漏加粗只是不够醒目，误加粗是整段噪音**。

纯字符串（不写 `type`）时按保守策略兜底：`一、`/`（一）` 判为标题，
`1.`/`（1）`/`①` 一律按编号正文（不加粗）。

详见 `references/paragraph-roles.md`。

## 使用

```bash
# 推荐：提取 → 定角色 → 渲染
python3 scripts/inspect_docx.py 原文.docx -o content.json
# 审阅 content.json 的 _uncertain 存疑清单，定稿 type
python3 scripts/generate_gongwen.py --config content.json --output 公文.docx

# 直接生成（已有 content.json）
python3 scripts/generate_gongwen.py --config content.json --output 公文.docx

# 纯文本正文（每行一段）
python3 scripts/generate_gongwen.py --title "关于XXX的请示" --recipient "XX厅" \
    --body-file body.txt --signature "XXX办公室" --date "2026年3月17日" --output 公文.docx
```

输出路径由 `--output` 指定，默认不落盘到桌面等固定位置。

也可作为模块调用：

```python
from scripts.generate_gongwen import build_document
build_document(content_dict, '公文.docx')
```

## 目录结构

```
official-doc/
├── SKILL.md                      # 技能定义（触发条件、工作流程、约束）
├── README.md
├── LICENSE
├── references/
│   ├── format-spec.md            # 完整格式规范（标注国标条款号）
│   ├── paragraph-roles.md        # 段落角色判定规则（标题 vs 编号正文）
│   └── numbering-rules.md        # 编号层级检查规则
└── scripts/
    ├── generate_gongwen.py       # 公文生成器（按角色渲染）
    └── inspect_docx.py           # 提取原文 + 样式线索 + 角色建议 + 存疑清单
```

## 格式速查

| 元素 | 字体 | 字号 | 行距 |
|------|------|------|------|
| 标题 | 方正小标宋简体 | 二号 22磅 | 32磅 |
| 主送机关 / 正文 | 仿宋_GB2312 | 三号 16磅 | 28磅 |
| 第一层 `一、` | 黑体 | 三号 | 28磅 |
| 第二层 `（一）` | 楷体_GB2312 | 三号 | 28磅 |
| 三/四/五层 | 仿宋_GB2312 加粗 | 三号 | 28磅 |
| 页码 | 宋体 | 四号 14磅 | — |

页面 A4：上 37mm、下 35mm、左 28mm、右 26mm，版心 156mm × 225mm。

## 说明

- 国标只规定到第四层编号（7.3.3），第五层「①」为实务扩展；第三层及以下「加粗」亦为通行做法，非国标强制。
- 页码国标原文为「一字线上距版心下边缘 7mm」，脚本按 Word 页脚换算值 4.58mm 设置 `footer_distance`，两者不冲突。

## 许可证

[MIT](LICENSE)

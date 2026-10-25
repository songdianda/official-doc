# official-doc

按《党政机关公文格式》国家标准 **GB/T 9704-2012** 把内容排版成规范的 `.docx` 公文。

## 功能

- 标准排版：标题、主送机关、正文、附件说明、发文机关署名、成文日期、页码（奇偶页不同）
- 编号检查：自动识别并校验「一、 →（一）→ 1. →（1）→ ①」五层结构
- 双输入：JSON 配置，或纯文本正文（自动识别层级）
- 零硬编码：标题、主送机关、正文、附件、署名、日期、输出路径全部可传参

## 安装

```bash
pip install python-docx
```

**字体需自备**（本仓库不含字体文件，字体受著作权保护，商用请自行取得授权）：

| 用途 | 字体 |
|------|------|
| 标题 | 方正小标宋简体 / 小标宋体 |
| 正文、主送机关 | 仿宋_GB2312 或 仿宋 |
| 第二层标题 | 楷体_GB2312 或 楷体 |
| 第一层标题 | 黑体 |
| 页码 | 宋体 |

python-docx 只能设置字体名称、无法嵌入字体文件，未安装时 Word 会用默认字体替代。可先自检：

```bash
python3 scripts/generate_gongwen.py --check-fonts
```

## 使用

```bash
# JSON 配置
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
│   └── numbering-rules.md        # 编号层级检查规则
└── scripts/
    └── generate_gongwen.py       # 公文生成器
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

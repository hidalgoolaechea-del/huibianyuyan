# report.config.json 字段说明

`scripts/generate_report.py` 读取单个 JSON 配置文件，输出一份 `.docx` 报告。
所有中文内容直接用 UTF-8 书写；路径建议用 `/`（Windows 也接受）。

## 顶层字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `output` | 是 | 输出 `.docx` 绝对路径，可用命令行 `-o` 覆盖 |
| `school_header` | 否 | 报告大标题，默认 `计电学院《汇编语言》实验报告` |
| `project` | 是 | 实验项目名，如 `编程、编译、连接、跟踪` |
| `date` | 是 | 实验日期，按模板写成 `2026年11月11号` |
| `lab_reference` | 否 | 教材与页码，用于自动生成「实验设备与实验环境」第 3 条 |
| `class_name` / `name` / `student_id` / `score` | 否 | **留空字符串**即渲染为带下划线的填空位；默认全部留空，禁止替用户代填 |
| `images_dir` | 是（有图片时） | 截图目录；绝对路径，或相对「配置文件所在目录」的路径 |
| `image_width_cm` | 否 | 图片统一宽度，默认 `11.8`（页宽 14.6cm） |
| `goals` | 是 | 实验目的，字符串数组，建议 3～4 条 |
| `sections` | 是 | 实验内容和原理的小节数组，结构见下 |
| `environment` | 否 | 覆盖默认的 3 条实验环境说明 |
| `thoughts` | 是 | 实验心得，字符串数组，建议 2 段 |

## sections / blocks

`sections` 中每个小节：

```json
{
  "title": "（2）编译、连接",
  "blocks": [ ... ]
}
```

`blocks` 按顺序渲染，支持 4 种类型：

| type | 字段 | 渲染效果 |
|------|------|----------|
| `text` | `value`, 可选 `bold` / `before` / `after` | 正文段落（首行缩进 2 字符，行距 18pt） |
| `code` | `lines`: 字符串数组 | 汇编源代码，Consolas 9pt，左缩进 30pt，行距 11.5pt |
| `data` | `rows`: `[[指令, 结果], ...]` | 指令 → 结果对照表（等宽对齐，用于「写出每步寄存器内容」类题目） |
| `image` | `file`, 可选 `width_cm` | 居中插入截图，**不生成任何图注** |

示例：

```json
{ "type": "image", "file": "s3_asm.png", "width_cm": 11.0 }
```

## 硬性约定

1. **个人信息留空**：`class_name` / `name` / `student_id` / `score` 一律填 `""`，生成的是空白下划线，由使用者手写/自行填写。
2. **图片无图注**：`image` 区块只插入图片本身；不要另建 `text` 区块写「图 1 …」之类的说明。
3. **精简优先**：所有报告默认按精简版式产出（见 `references/workflow.md` 的「精简写作规则」）。

## 最小可运行配置

```json
{
  "output": "./output/实验报告.docx",
  "project": "实验名称",
  "date": "2026年11月11号",
  "images_dir": "./shots",
  "goals": ["（1）…", "（2）…", "（3）…"],
  "sections": [
    { "title": "（1）实验步骤", "blocks": [
      { "type": "text", "value": "按教材步骤完成…" },
      { "type": "image", "file": "shot1.png" }
    ] }
  ],
  "thoughts": ["心得第 1 段…", "心得第 2 段…"]
}
```

完整可参考 `assets/report.config.example.json`（实验三真实内容）。

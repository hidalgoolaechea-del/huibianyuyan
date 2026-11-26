---
name: assembly-lab-report
description: 生成《汇编语言》（王爽，第 2 版）课程实验过程报告 Word 文档（.docx），按实验报告模板版式排版，支持从教材 PDF 读取实验任务、用 DOSBox + MASM + DEBUG 采集运行截图、核对真实寄存器/内存数据并插入报告。当用户要求「做汇编实验报告 / 写实验二报告 / 填写汇编实验报告 / 把 DEBUG 截图插到报告里」，或提供教材 PDF、实验报告模板 docx、.asm 源程序并希望产出实验报告时应使用本技能。触发词：汇编实验报告、汇编语言实验、实验报告模板、DEBUG 截图、MASM、LINK、DOSBox、王爽《汇编语言》、实验一～实验五。适用于 Windows + DOSBox 0.74 + MASM 5.0 环境。This skill should be used when a user asks to produce or fill in an assembly-language lab report with real DEBUG screenshots.
agent_created: true
---

# 汇编语言实验过程报告生成器

## 概览

把「做一次汇编实验报告」的完整流程封装为可复用的配置驱动流水线：读取教材任务 → 编写/运行程序并采集 DOSBox 截图 → 核对截图中的真实数据 → 填充配置 JSON → 生成符合模板版式的 `.docx` → 检查排版 → 交付。
版式沿用学生提供的「实验一/实验二」实验报告模板：标题栏、信息栏、单列边框表格（实验目的 / 实验内容和原理 / 实验设备与实验环境）、实验心得表格。

## 何时使用

在以下任一情况触发：

- 用户要求制作、填写或继续做《汇编语言》实验一～实验五（或同类）实验报告；
- 用户给出教材 PDF、实验报告模板 `.docx`、截图目录或 `.asm` 源程序，并希望生成报告；
- 用户要求把 DOSBox / DEBUG 的运行截图插入实验报告对应位置；
- 用户要求按已有实验报告模板改日期、换实验内容再产出一份新报告。

不要在用户只询问某条汇编指令含义、或只想调试程序而不产出报告时使用。

## 固定约定（必须遵守，不得因用户未提及而擅自改变）

1. **个人信息一律留空**：班级、姓名、学号、综合成绩只保留带下划线的填空位；不要向用户询问姓名学号，更不要沿用任何历史报告中的姓名/学号/班级。配置中这四个字段固定为 `""`。
2. **截图下方不加图注**：只插入图片本身，不生成「图 1 …」「图 2 …」等说明文字；也不把图注改写成正文段落。
3. **默认精简**：页数 3～4 页、截图 2～6 张、实验目的 3～4 条、心得 2 段、环境 3 条。用户说「内容不要太多」时进一步压缩文字，但保留核心实验数据。

## 快速开始（多数请求走这条路径）

当实验任务明确、且已有截图或只需按模板重排内容时：

1. 复制 `assets/report.config.example.json` 到当前工作目录，改名如 `report.config.json`；
2. 按本次实验改写配置：`output`、`project`、`date`、`lab_reference`、`images_dir`、`goals`、`sections`、`thoughts`；
3. 运行生成脚本：

   ```bash
   python scripts/generate_report.py report.config.json
   ```

4. 检查排版：

   ```bash
   python scripts/verify_layout.py "<输出报告.docx>" --out sheet.png
   ```

   用 `ViewImage` 查看 `sheet.png`，确认页数、边框、图片、空白页都正常；
5. 用 `present_files` 呈现 `.docx`，并在回复中说明页数、截图数、日期和关键数据。

## 完整工作流

执行前读 `references/workflow.md`；配置字段读 `references/config_schema.md`；环境安装与踩坑读 `references/environment.md`。

| 步骤 | 动作 | 用到的资源 |
|------|------|-----------|
| 0 | 解析用户输入：实验编号/名称、日期、教材路径、模板、截图目录；缺失项按默认值处理 | `references/workflow.md` Step 0 |
| 1 | 从扫描版教材 PDF 定位并读取实验任务（先渲染目录、再渲染单页高清图） | `pymupdf`；页码偏移规则见 workflow |
| 2 | 写 `.asm` → MASM/LINK 编译 → DEBUG 采集截图（已有截图则跳过） | `scripts/capture_dosbox.ps1`、`scripts/WinCap.cs`、`assets/sessions.example.json` |
| 3 | 预处理截图（裁黑边 + 2 倍 NEAREST 放大） | `scripts/prepare_images.py` |
| 4 | 逐张查看截图，把真实寄存器值、内存数据、编译输出写进配置 | `ViewImage` |
| 5 | 生成 `.docx` | `scripts/generate_report.py` |
| 6 | Word 导出 PDF → 页数统计 → 缩略拼图目视检查 | `scripts/verify_layout.py` |
| 7 | 交付并汇总 | `present_files` |

### 关键脚本用法

```bash
# 截图预处理
python scripts/prepare_images.py <原图目录> <输出目录> --scale 2

# DOSBox 批量截图（Windows PowerShell）
powershell -ExecutionPolicy Bypass -File scripts/capture_dosbox.ps1 \
  -SessionsJson sessions.json -OutDir shots \
  -DosBoxExe "C:\DOSBox\DOSBox.exe" -Conf "C:\masm\dosbox.conf"

# 只重拍一张
powershell -ExecutionPolicy Bypass -File scripts/capture_dosbox.ps1 \
  -SessionsJson sessions.json -OutDir shots -Only s3_trace -NoRestart

# 生成报告 / 覆盖输出路径
python scripts/generate_report.py report.config.json -o "D:\out\报告.docx"

# 一次检查多份报告
python scripts/verify_layout.py 报告3.docx 报告4.docx 报告5.docx --out sheet.png
```

## 内容写作规则

- 每个小节由「标题 + 说明文字 + 可选源码/数据表 + 截图」组成，说明文字只写「做了什么、看到什么、说明什么」；
- 教材要求「写出每一步执行后寄存器内容」时，用 `data` 区块列出 `指令 → 结果`，不要写成长段落；
- 源程序用 `code` 区块（Consolas 9pt）列出，保持等宽对齐；
- 报告中出现的所有数值必须来自实际截图，不得凭教材示例或记忆填写；
- 心得写 2 段，围绕本次实验最意外/最有收获的现象，不写空话。

## 典型用户请求（用于理解触发意图）

- 「用这个技能帮我做实验三的报告，日期 11 月 11，教材在桌面，截图我已经放在 shots 文件夹了」
- 「按我发的实验报告一模板做实验四，内容不要太多，姓名学号先空着」
- 「继续做实验二、三、四、五，日期分别是 11/28、11/11、11/25、12/9」
- 「把这几张 DEBUG 截图插到报告对应位置，不需要图注」

## 资源清单

- `scripts/generate_report.py` — 配置驱动的 `.docx` 生成器（核心）
- `scripts/prepare_images.py` — 截图裁边放大
- `scripts/verify_layout.py` — Word→PDF 页数与拼图检查
- `scripts/capture_dosbox.ps1` + `scripts/WinCap.cs` — DOSBox 客户区截图自动化
- `assets/report.config.example.json` — 完整示例配置（实验三真实内容）
- `assets/sessions.example.json` — 截图会话示例
- `references/config_schema.md` — 配置字段与区块类型
- `references/workflow.md` — 端到端流程、教材定位、精简写作规则
- `references/environment.md` — 依赖安装与 Windows/MASM/DEBUG 踩坑清单

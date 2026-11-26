# 端到端工作流

本文件描述从一个「做实验 X 报告」的请求到交付 `.docx` 的完整过程。
环境准备与踩坑见 `references/environment.md`，配置字段见 `references/config_schema.md`。

## Step 0 · 解析用户输入

用户通常提供（缺项按默认值处理）：

| 输入 | 缺失时的处理 |
|------|--------------|
| 实验编号 / 名称 | 必问，或从教材目录推断后确认 |
| 实验日期 | 必问；按模板惯例写 `2026年11月11号` |
| 教材 PDF 路径 | 默认桌面 `《汇编语言_第2版》.pdf` |
| 报告模板 docx | 默认沿用本 skill 的固定版式（与实验一模板一致） |
| 截图目录 | 已有截图则直接用；否则进入 Step 2 采集 |
| 内容多少 | 默认**精简**（页数 3～4、截图 2～6 张） |
| 姓名 / 学号 | **不要问、不要填**，保持空白 |

## Step 1 · 读取教材中的实验任务

教材常见为**纯扫描版 PDF**（文字层只有水印），必须渲染成图片再 OCR/目视读取：

```python
import pymupdf, os
doc = pymupdf.open(r'...\book.pdf')          # 若中文路径报错，先复制到纯 ASCII 路径
for pno in range(99, 113):
    doc[pno-1].get_pixmap(dpi=100).save('big%03d.png' % pno)
```

定位技巧：
- 先渲染前 12 页目录确认印刷页码；本教材存在 **PDF 页 = 印刷页 + 12** 的偏移；
- 目录不清晰时，按 40 页一段生成缩略拼图定位章节，再渲染单页高清图；
- 已验证的教材位置：实验二 PDF 83–87 页、实验三约 106 页、实验四 133–134 页、实验五 145–148 页。

从任务描述中提取：源程序、要回答的问题、需要记录的数据表、需要验证的内存区间。

## Step 2 · 编写程序并采集运行截图

若用户已提供截图，跳过本步。否则：

1. 在 MASM 目录（如 `C:\masm`）写 `.asm` 源文件；
2. 用 `capture_dosbox.ps1` + `assets/sessions.example.json` 定义每个截图会话；
   `sessions` 每项含 `name`（截图名）、`start`（如 `debug t1.exe`）、`lines`（DEBUG 命令序列）；
3. 运行采集脚本，输出到 `shots/`；
4. 用 `prepare_images.py <shots> <final> --scale 2` 裁黑边并放大，得到 `final/` 供报告引用。

常用 DEBUG 命令：`r` `d` `e` `a` `u` `t` `p` `g` `q`；
段寄存器形式 `d ds:0`、`e ds:0`、`a`、`u ds:0`；长度参数写 `l`（`d 2200:f0 l 10`）。

## Step 3 · 核对截图中的真实数据

**必须**逐张查看截图（放大后更易读），把真实数值写进报告，不要用猜的：
- 编译：`0 Warning Errors / 0 Severe Errors`、`LINK : warning L4021: no stack segment`、EXE 字节数；
- 加载：寄存器初值、`CX`（程序长度）、`CS:IP`、`DS`（指向 PSP）、PSP 头 `CD 20`；
- 跟踪：每一步 `AX/BX/SS/SP/IP` 的变化、`mov ss,ax` 后下一条指令被紧接着执行的现象；
- 结果：`d` 命令显示的内存内容（如 `00 01 02 … 3F`）。

多次 `ViewImage` 同一张图会被判定为重复；需要重看时先裁剪成上下半或换个文件名。

## Step 4 · 编写配置 JSON

复制 `assets/report.config.example.json` 到工作目录，按本次实验改写：
- `project` / `date` / `lab_reference` / `output`；
- `images_dir` 指向 Step 2 的 `final/`；
- `goals` 3～4 条、`thoughts` 2 段；
- `sections` 按教材任务分小节；每节用 `text` / `code` / `data` / `image` 区块组织；
- `class_name` / `name` / `student_id` / `score` 保持 `""`。

## Step 5 · 生成报告

```bash
python scripts/generate_report.py report.config.json
```

脚本输出目标 `.docx`，并在截图文件缺失时明确报错。

## Step 6 · 排版检查

```bash
python scripts/verify_layout.py "<报告.docx>" --out sheet.png
```

检查：页数是否 3～4 页、表格边框是否完整、图片有无超出页面或被截断、有无空白页或孤行。
用 `ViewImage` 查看 `sheet.png` 确认后再交付。

## Step 7 · 交付

用 `present_files` 呈现 `.docx`，并在回复中说明页数、截图数、日期与关键实验数据。

## 精简写作规则（默认）

- 页数控制在 **3～4 页**，截图 **2～6 张**；
- 实验目的 3～4 条、心得 2 段、环境 3 条；
- 每个小节：标题 + 1 段说明文字（必要时再加源码/数据表）+ 1 张截图；
- 说明文字只写「做了什么、看到什么、说明什么」，不抄教材原文、不展开原理长篇；
- 需要填写结果表时，用 `data` 区块紧凑列出，不要写成大段文字；
- 图片下方**不加图注**。

## 固定版式参数

A4；页边距 左右 3.17cm、上下 2.54cm；正文宋体 10.5pt、行距 18pt；
标题 16pt 加粗居中；信息栏 12pt 带下划线填空；单列边框表格；
单元格小标题 14pt；图片宽 11.8cm 居中。

# assembly-lab-report

> WorkBuddy Skill · 一键生成《汇编语言》（王爽，第 2 版）课程实验过程报告（.docx）
>
> 仓库地址：https://github.com/hidalgoolaechea-del/huibianyuyan

把「做一份汇编实验报告」的完整流程封装为可复用的配置驱动流水线：读取教材任务 → DOSBox / MASM / DEBUG 采集运行截图 → 核对真实寄存器与内存数据 → 填充配置 JSON → 生成符合实验报告模板版式的 Word 文档 → 检查排版 → 交付。

版式沿用课程实验报告模板：标题栏、信息栏、单列边框表格（实验目的 / 实验内容和原理 / 实验设备与实验环境）、实验心得表格。

## ✨ 特性

- **配置驱动**：一个 JSON 文件描述整份报告，改内容不用改代码
- **个人信息留空**：班级 / 姓名 / 学号 / 综合成绩只保留下划线填空位，不会带出任何人的信息
- **截图无图注**：图片只插入本身，不生成「图 1 …」之类说明
- **默认精简**：3～4 页、2～6 张截图、目的 3～4 条、心得 2 段
- **真实数据**：报告里的数值来自实际 DEBUG 运行截图，不是教材示例抄写
- **全流程工具**：教材扫描页渲染、DOSBox 自动截图、截图放大、Word 排版检查脚本都附带

## 📁 目录结构

```
assembly-lab-report/
├── SKILL.md                          # Skill 入口（触发条件、工作流、固定约定）
├── assets/
│   ├── report.config.example.json    # 完整示例配置（实验三真实内容）
│   └── sessions.example.json         # DOSBox 截图会话示例
├── references/
│   ├── config_schema.md              # 配置字段说明
│   ├── workflow.md                   # 端到端流程、教材定位、精简写作规则
│   └── environment.md                # 依赖安装与 Windows/MASM/DEBUG 踩坑清单
├── scripts/
│   ├── generate_report.py            # 核心：配置 → .docx
│   ├── prepare_images.py             # 截图裁黑边 + 整数倍放大
│   ├── verify_layout.py              # Word→PDF 页数与缩略拼图检查
│   ├── capture_dosbox.ps1            # DOSBox 客户区截图自动化
│   └── WinCap.cs                     # 窗口聚焦 / 截屏 P/Invoke 辅助类
└── dist/
    └── assembly-lab-report.zip       # 可直接分享的安装包
```

## 🔧 安装

**方式 A · 下载 zip（推荐给不熟 git 的同学）**

下载仓库里的 `dist/assembly-lab-report.zip`（直链：https://github.com/hidalgoolaechea-del/huibianyuyan/raw/main/dist/assembly-lab-report.zip），解压后把整个 `assembly-lab-report` 文件夹放到：

```
C:\Users\<你的用户名>\.workbuddy\skills\assembly-lab-report\
```

**方式 B · clone 仓库**

```bash
git clone https://github.com/hidalgoolaechea-del/huibianyuyan.git \
  "C:/Users/<你的用户名>/.workbuddy/skills/assembly-lab-report"
```

放好后**新开一个 WorkBuddy 对话**即可生效（或重启 WorkBuddy）。

> 也可以放到某个项目的 `.workbuddy/skills/` 下，只在该项目里共享给团队成员。

## 💬 怎么用（在 WorkBuddy 对话里输入）

**点名调用（最稳）**

```
用 assembly-lab-report 技能帮我做汇编语言实验三的报告，日期 2026年11月11号，
教材PDF在桌面，截图放在 C:\shots，内容不要太多
```

**斜杠指令**

```
/assembly-lab-report 按实验报告模板做实验四报告，日期 11月25，姓名学号先空着，图片不用加图注
```

**直接说需求也会自动命中**

```
帮我写汇编实验五的实验报告，日期 12月9，把这几张 DEBUG 截图插到对应位置
继续做实验二、三、四、五，日期分别是 11/28、11/11、11/25、12/9，注意内容不要太多
```

一次说清这 4 项即可：① 实验编号 / 名称 ② 日期 ③ 教材 PDF 或实验内容 ④ 截图目录。
没有截图时，Skill 会走 DOSBox 采集流程；其余（留空个人信息、无图注、精简、版式、排版检查、交付）自动完成。

## 🖥️ 手动运行脚本

```bash
# 依赖（Windows，建议用腾讯云镜像）
pip install python-docx pillow pymupdf pywin32 \
  -i https://mirrors.cloud.tencent.com/pypi/simple \
  --trusted-host mirrors.cloud.tencent.com

# 截图预处理
python scripts/prepare_images.py <原图目录> <输出目录> --scale 2

# 由配置生成报告
python scripts/generate_report.py report.config.json -o "D:\out\报告.docx"

# 检查排版（Word → PDF → 缩略拼图）
python scripts/verify_layout.py 报告.docx --out sheet.png
```

DOSBox 批量截图（Windows PowerShell）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/capture_dosbox.ps1 `
  -SessionsJson sessions.json -OutDir shots `
  -DosBoxExe "C:\DOSBox\DOSBox.exe" -Conf "C:\masm\dosbox.conf"
```

## 🧩 配置速览

```json
{
  "output": "./output/实验报告.docx",
  "project": "编程、编译、连接、跟踪",
  "date": "2026年11月11号",
  "images_dir": "./shots",
  "goals": ["（1）…", "（2）…", "（3）…"],
  "sections": [
    {
      "title": "（1）实验步骤",
      "blocks": [
        { "type": "text",  "value": "按教材步骤完成…" },
        { "type": "code",  "lines": ["assume cs:codesg", "codesg segment"] },
        { "type": "data",  "rows": [["mov ax,2000H", "AX=2000H"]] },
        { "type": "image", "file": "shot1.png" }
      ]
    }
  ],
  "thoughts": ["心得第 1 段…", "心得第 2 段…"]
}
```

完整字段说明见 [`references/config_schema.md`](references/config_schema.md)，可直接复制 [`assets/report.config.example.json`](assets/report.config.example.json) 修改。

## 📌 固定约定

1. 个人信息一律留空，不询问、不代填
2. 截图下方不加图注
3. 默认精简，用户说「内容不要太多」时进一步压缩文字，但保留核心实验数据

## ⚠️ 环境要求

Windows 11 + DOSBox 0.74 + MASM 5.0（`MASM.EXE` / `LINK.EXE` / `DEBUG.EXE`）+ Python 3.13。
若只需按已有截图排版，不运行 DOSBox 也可以。

本仓库**不包含教材扫描件**，也不包含任何学生个人信息；示例数据均来自 DOSBox 实际运行截图。

## 📄 License

[MIT](LICENSE)

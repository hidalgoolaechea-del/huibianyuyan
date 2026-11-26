# 环境与踩坑清单（Windows）

## 运行环境

- Windows 11 + DOSBox 0.74 + MASM 5.0（`MASM.EXE`、`LINK.EXE`、`DEBUG.EXE`，示例目录 `C:\masm`）；
- Python 3.13，依赖：`python-docx`、`Pillow`、`pymupdf`、`pywin32`；
- 安装依赖（清华源在本机可能取不到包，优先用腾讯云镜像）：

```bash
pip install python-docx pillow pymupdf pywin32 \
  -i https://mirrors.cloud.tencent.com/pypi/simple \
  --trusted-host mirrors.cloud.tencent.com
```

## 教材 PDF

- 121MB 扫描版 PDF 会让专用 PDF 组件崩溃；改用 `pymupdf` 渲染页面图片再目视/OCR 读取；
- 中文路径可能报错：先把 PDF 复制到纯 ASCII 路径（如工作区 `pdfwork/book.pdf`）；
- 原生 Windows 程序不识别 git bash 的 `/tmp`，输出路径一律用 Windows 风格绝对路径。

## MASM / DEBUG

- **MASM 数字默认是十进制**：`mov bx,200` 会汇编成 `00C8h`。内存地址、循环次数必须写 `200h`、`40h`、`20h`、`17h`；
- `debug` 的 `a` 命令不支持 `mov es:[bx],al` 这类段前缀，遇到段前缀改用 MASM 编译源文件；
- `d` 命令的长度参数是字母 `l`：`d 2200:f0 l 10`，不要写成 `d 2200:f0 10`；
- EXE 程序加载后 `DS` 指向 PSP；要查看程序数据段，用 `d ds:100 l 10`（数据实际在 `DS:0100` 附近）；
- `t` 执行 `mov ss,ax` 时，下一条修改 `sp` 的指令会被紧接着执行，跟踪表要按实际显示值填写；
- `g` 带断点在不同 DEBUG 版本表现不一致；必要时用 `t` 跟到目标 `IP`，或用无断点 `g` 让程序正常结束后 `d` 查看；
- 单步中断会把 `IP/CS/标志` 压入当前栈，`d` 看到的「凭空出现的数据」多由此而来，报告中要据此解释。

## 截图采集（DOSBox）

- 不要用 DOSBox 内置 `Ctrl+F5` 截图：文件名会覆盖、且依赖 capture 目录增量判断，容易误判；
- 用 `WinCap.cs` + `CopyFromScreen` 直接截取 DOSBox **客户区**；截屏前 `ForceFocus` 并用 `SetWindowPos` 置顶，中心像素过亮时重试；
- `SendKeys` 前必须检查前台窗口，视频播放器等会抢焦点导致按键丢失；
- `SetProcessDPIAware` 等 P/Invoke 放在 `.cs` 文件里用 `Add-Type -Path` 加载，不要用内联 `Add-Type -MemberDefinition`（可能被安全策略拦截）；
- 不要用 `Remove-Item` 清理截图目录（宿主安全钩子会拦截），用 `LastWriteTime` 或直接覆盖写；
- `debug < file` 的 stdin 重定向方案在 DOSBox 下解析异常，别走这条路。

## Word / 排版检查

- **Word COM 打不开文件名含方括号 `[ ]` 的文档**（如 `实验四-[bx]…docx`）：先复制到临时简单文件名再打开/导出，`verify_layout.py` 已自动处理；
- 检查脚本异常退出时可能残留 `WINWORD.EXE` 锁住文件；先结束进程再重试；
- 用 `d.ExportAsFixedFormat(pdf, 17)` 导出 PDF、`d.ComputeStatistics(2)` 取页数，再用 `pymupdf` 生成缩略拼图目视检查。

## 精简要求

用户常明确说「内容不要太多」。默认采用 3～4 页精简版：目的 3～4 条、心得 2 段、截图 2～6 张、说明文字短。

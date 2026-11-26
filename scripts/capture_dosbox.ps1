<#
.SYNOPSIS
    通用 DOSBox + MASM/LINK/DEBUG 截图采集器（Windows）。

.DESCRIPTION
    按 sessions JSON 的定义，自动在 DOSBox 中依次执行命令并截取客户区画面。
    采用 SendKeys 注入 + 前台窗口强制聚焦 + CopyFromScreen 直接截取客户区，
    不依赖 DOSBox 内置 Ctrl+F5 截图（文件名混乱、易被其他窗口遮挡）。

.PARAMETER SessionsJson
    sessions 定义文件（JSON 数组），字段：
      name  截图文件名（不含 .png）
      start DOS 提示符下先执行的命令，如 "debug t1.exe"；"none" 表示只停在提示符
      lines 之后逐行执行的命令数组，如 ["r","d ds:0 l 2","u 0 l 6"]
.PARAMETER OutDir
    截图输出目录。
.PARAMETER DosBoxExe
    DOSBox.exe 路径。
.PARAMETER Conf
    DOSBox 配置文件路径（需 mount 好 MASM 目录）。
.PARAMETER Only
    只采集指定 name 的 session，便于重拍单张。
.PARAMETER NoRestart
    不重启 DOSBox，复用当前已打开的窗口。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File capture_dosbox.ps1 -SessionsJson sessions.json -OutDir shots
#>
param(
    [Parameter(Mandatory = $true)][string]$SessionsJson,
    [string]$OutDir = '.\shots',
    [string]$DosBoxExe = 'C:\DOSBox\DOSBox.exe',
    [string]$Conf = '',
    [string]$Only = '',
    [switch]$NoRestart
)

$ErrorActionPreference = 'Stop'
Add-Type -Path (Join-Path $PSScriptRoot 'WinCap.cs') -ReferencedAssemblies System.Drawing
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
[WinCap]::SetProcessDPIAware() | Out-Null

$outDir = [System.IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$sessions = Get-Content -Raw -Encoding UTF8 $SessionsJson | ConvertFrom-Json

function Get-DosBox {
    $p = Get-Process DOSBox -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $p) { throw 'DOSBox 未运行' }
    $p.Refresh()
    return $p
}

# SendKeys 特殊字符转义
function Esc([string]$s) {
    if ($s -eq '') { return '' }
    return ($s -replace '([+^%~(){}\[\]])', '{$1}')
}

function Send-Line([string]$text, [IntPtr]$hwnd) {
    if ($hwnd -ne [IntPtr]::Zero -and [WinCap]::GetForegroundWindow() -ne $hwnd) {
        [WinCap]::ForceFocus($hwnd) | Out-Null
        Start-Sleep -Milliseconds 200
    }
    if ($text -ne '') { [System.Windows.Forms.SendKeys]::SendWait((Esc $text)) }
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    # g 命令会让程序全速运行，给 DEBUG 多一点刷新时间
    if ($text -match '^g') { Start-Sleep -Seconds 3 } else { Start-Sleep -Milliseconds 250 }
}

function Shot([string]$name, [IntPtr]$hwnd) {
    Start-Sleep -Milliseconds 800
    for ($i = 0; $i -lt 3; $i++) {
        if ([WinCap]::GetForegroundWindow() -ne $hwnd) {
            [WinCap]::ForceFocus($hwnd) | Out-Null
            Start-Sleep -Milliseconds 400
        } else { break }
    }
    # 置顶 DOSBox，避免被 WorkBuddy 等窗口盖住
    [WinCap]::SetWindowPos($hwnd, [IntPtr](-1), 0, 0, 0, 0, 0x0041) | Out-Null
    Start-Sleep -Milliseconds 500
    $cr = New-Object WinCap+RECT
    [WinCap]::GetClientRect($hwnd, [ref]$cr) | Out-Null
    $pt = New-Object WinCap+RECT
    $pt.Left = 0; $pt.Top = 0
    [WinCap]::ClientToScreen($hwnd, [ref]$pt) | Out-Null
    $bmp = New-Object System.Drawing.Bitmap([Math]::Max($cr.Right, 1), [Math]::Max($cr.Bottom, 1))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($pt.Left, $pt.Top, 0, 0, (New-Object System.Drawing.Size($cr.Right, $cr.Bottom)))
    $g.Dispose()
    # DEBUG 画面大部分是黑的；中心像素过亮说明截到了别的窗口，重试一次
    $center = $bmp.GetPixel([int]($cr.Right / 2), [int]($cr.Bottom / 2))
    if (($center.R + $center.G + $center.B) -gt 200) {
        Start-Sleep -Milliseconds 800
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($pt.Left, $pt.Top, 0, 0, (New-Object System.Drawing.Size($cr.Right, $cr.Bottom)))
        $g.Dispose()
    }
    $target = Join-Path $outDir ($name + '.png')
    $bmp.Save($target, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Output "saved $($name + '.png') ($($cr.Right)x$($cr.Bottom))"
}

function Reset-To([string]$startCmd, [IntPtr]$hwnd) {
    Send-Line 'q' $hwnd
    Start-Sleep -Milliseconds 600
    Send-Line 'cls' $hwnd
    Start-Sleep -Milliseconds 600
    if ($startCmd -ne 'none' -and $startCmd -ne '') {
        Send-Line $startCmd $hwnd
        Start-Sleep -Milliseconds 1200
    }
}

if (-not $NoRestart) {
    Stop-Process -Name DOSBox -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    if ($Conf -ne '') {
        Start-Process -FilePath $DosBoxExe -ArgumentList '-conf', $Conf, '-noconsole' | Out-Null
    } else {
        Start-Process -FilePath $DosBoxExe -ArgumentList '-noconsole' | Out-Null
    }
    Start-Sleep -Seconds 6
}

foreach ($s in $sessions) {
    if ($Only -and $s.name -ne $Only) { continue }
    $ok = $false
    for ($attempt = 1; $attempt -le 3 -and -not $ok; $attempt++) {
        try {
            $proc = Get-DosBox
            [WinCap]::ForceFocus($proc.MainWindowHandle) | Out-Null
            Start-Sleep -Milliseconds 400
            Reset-To $s.start $proc.MainWindowHandle
            foreach ($l in $s.lines) { Send-Line $l $proc.MainWindowHandle }
            Shot $s.name $proc.MainWindowHandle
            $ok = $true
        } catch {
            Write-Output "attempt $attempt failed for $($s.name): $($_.Exception.Message)"
            if (-not $NoRestart) {
                Stop-Process -Name DOSBox -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                if ($Conf -ne '') {
                    Start-Process -FilePath $DosBoxExe -ArgumentList '-conf', $Conf, '-noconsole' | Out-Null
                } else {
                    Start-Process -FilePath $DosBoxExe -ArgumentList '-noconsole' | Out-Null
                }
                Start-Sleep -Seconds 6
            }
        }
    }
}
Write-Output 'DONE'

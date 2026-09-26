<#
.SYNOPSIS
  在 Windows 上编译根目录的「OneNote导出工具.exe」启动器。

.DESCRIPTION
  启动器是个极小的原生程序（无运行时依赖），只负责用系统自带的
  Windows PowerShell 拉起 scripts\export-gui.ps1，并保证不闪控制台窗口。
  真正的界面逻辑仍在 scripts/export-gui.ps1 里，改界面不需要重新编译。

  只有改启动器逻辑或换图标时才需要跑本脚本。

  依赖（任选一套）：
    - MinGW-w64（gcc + windres）：MSYS2、w64devkit、Scoop 的 mingw 包都行
    - MSVC（cl + rc）：装了 Visual Studio 或 Build Tools 即可

.EXAMPLE
  # 在仓库根目录
  powershell -ExecutionPolicy Bypass -File scripts\launcher\build-exe.ps1
#>
# 注意：本文件必须以「UTF-8 带 BOM」保存。PowerShell 5.1 读无 BOM 的 .ps1 时
# 按系统 ANSI 代码页解码，文件里的中文会变成乱码，进而引发一连串语法错误
# （症状是报「意外的标记」并指出一行乱码）。改这个文件后请确认 BOM 还在。
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$here     = $PSScriptRoot
$repoRoot = (Resolve-Path (Join-Path $here "..\..")).Path
$outExe   = Join-Path $repoRoot "OneNote导出工具.exe"

$icon    = Join-Path $here "icon.ico"
$rcFile  = Join-Path $here "launcher.rc"
$cFile   = Join-Path $here "launcher.c"
$versionFile = Join-Path $repoRoot "VERSION"

foreach ($f in @($cFile, $rcFile, $icon)) {
  if (-not (Test-Path -LiteralPath $f)) {
    throw "缺少源文件：$f"
  }
}

function Find-Tool {
  param([string[]]$Names)
  foreach ($n in $Names) {
    $cmd = Get-Command $n -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
  }
  return $null
}

# 把 VERSION 里的版本号写进 exe 的版本资源，避免它和仓库版本各说各话。
# VERSION 缺失或格式不对时退回默认值，不中断编译。
$version = "1.0.0.0"
if (Test-Path -LiteralPath $versionFile) {
  $raw = (Get-Content -LiteralPath $versionFile -Raw).Trim()
  if ($raw -match '^(\d+)\.(\d+)\.(\d+)') {
    $version = "{0}.{1}.{2}.0" -f $Matches[1], $Matches[2], $Matches[3]
  }
}
# 资源脚本的 FILEVERSION / PRODUCTVERSION 必须是逗号分隔的四段数字，
# 写成 "1.0.0.0" 会编译失败；字符串版本才用点号。
$versionComma = $version -replace '\.', ','
Write-Host "版本资源：$version"

# launcher.rc 里 FILEVERSION / PRODUCTVERSION 写成占位符，编译前替换成实际版本
$rcText = (Get-Content -LiteralPath $rcFile -Raw) -replace '(FILEVERSION|PRODUCTVERSION)\s+\S+', ('$1     ' + $versionComma)
$rcText = $rcText -replace '("FileVersion",\s*)"[^"]*"', ('$1"' + $version + '"')
$rcText = $rcText -replace '("ProductVersion",\s*)"[^"]*"', ('$1"' + $version + '"')

# 编译产物不能与源文件混在一起，用临时目录
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("onenote-launcher-{0}" -f ([guid]::NewGuid().ToString("N")))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

# 临时 .rc 必须写成 UTF-8 无 BOM：资源编译器按 ANSI 代码页解析无 BOM 文件，
# 但写成带 BOM 又会被 windres 当成非法字符。这里只保留 ASCII 内容（版本号、
# 路径），本来就不含中文，所以无 BOM 的 UTF-8 是安全的。
$rcTmp = Join-Path $tmp "launcher.rc"
[System.IO.File]::WriteAllText($rcTmp, $rcText, (New-Object System.Text.UTF8Encoding($false)))

try {
  $gcc     = Find-Tool @("x86_64-w64-mingw32-gcc.exe", "gcc.exe")
  $windres = Find-Tool @("x86_64-w64-mingw32-windres.exe", "windres.exe")
  $cl      = Find-Tool @("cl.exe")
  $rc      = Find-Tool @("rc.exe")

  if ($gcc -and $windres) {
    Write-Host "使用 MinGW-w64 编译…"

    $resObj = Join-Path $tmp "launcher_res.o"
    # 工作目录设到 $here：launcher.rc 里写的是相对路径 icon.ico
    Push-Location $here
    try {
      & $windres -I $here -I $tmp -i $rcTmp -o $resObj
    } finally {
      Pop-Location
    }
    if ($LASTEXITCODE -ne 0) { throw "windres 失败（退出码 $LASTEXITCODE）" }

    # -municode：宽字符入口 wWinMain
    # -mwindows：GUI 子系统，不创建控制台
    # --stack：路径缓冲区按 32768 个宽字符开，默认栈偏紧，留足余量
    & $gcc -municode -mwindows -Os `
      -fno-asynchronous-unwind-tables -fno-ident `
      -ffunction-sections -fdata-sections `
      -Wall -Wextra `
      $cFile $resObj `
      -o $outExe `
      "-Wl,--gc-sections" "-Wl,--stack,8388608" -s -static-libgcc
    if ($LASTEXITCODE -ne 0) { throw "gcc 失败（退出码 $LASTEXITCODE）" }
  }
  elseif ($cl -and $rc) {
    Write-Host "使用 MSVC 编译…"

    $resFile = Join-Path $tmp "launcher.res"
    # rc.exe 同样按工作目录解析 launcher.rc 里的相对路径 icon.ico
    Push-Location $here
    try {
      & $rc /nologo /fo $resFile $rcTmp
    } finally {
      Pop-Location
    }
    if ($LASTEXITCODE -ne 0) { throw "rc 失败（退出码 $LASTEXITCODE）" }

    # /O1 体积优先；/MT 静态链接 CRT，避免依赖 VC 运行库
    Push-Location $tmp
    try {
      & $cl /nologo /O1 /MT /GS- /DUNICODE /D_UNICODE `
        $cFile $resFile `
        "/Fe:$outExe" `
        /link /SUBSYSTEM:WINDOWS /INCREMENTAL:NO
      if ($LASTEXITCODE -ne 0) { throw "cl 失败（退出码 $LASTEXITCODE）" }
    } finally {
      Pop-Location
    }
    # cl 会把中间文件写到当前目录，临时目录随用随删，不污染仓库
  }
  else {
    throw @"
找不到可用的 C 编译器。请任选一套装好再试：

  MinGW-w64（推荐，体积小）
    - Scoop:  scoop install mingw
    - MSYS2:  pacman -S mingw-w64-ucrt-x86_64-gcc
    - w64devkit: 解压后把 bin 目录加进 PATH

  MSVC
    - 安装 Visual Studio Build Tools，勾选「使用 C++ 的桌面开发」
    - 然后在「Developer PowerShell for VS」里运行本脚本
"@
  }

  if (-not (Test-Path -LiteralPath $outExe)) { throw "编译命令返回成功，但没有生成 $outExe" }

  $size = [Math]::Round((Get-Item -LiteralPath $outExe).Length / 1KB)
  Write-Host ""
  Write-Host ("已生成：{0}" -f $outExe)
  Write-Host ("大小：{0} KB" -f $size)
} finally {
  Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

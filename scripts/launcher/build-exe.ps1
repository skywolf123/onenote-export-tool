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
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$here     = $PSScriptRoot
$repoRoot = (Resolve-Path (Join-Path $here "..\..")).Path
$outExe   = Join-Path $repoRoot "OneNote导出工具.exe"

$icon    = Join-Path $here "icon.ico"
$rcFile  = Join-Path $here "launcher.rc"
$cFile   = Join-Path $here "launcher.c"

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

# 编译产物不能与源文件混在一起，用临时目录
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("onenote-launcher-{0}" -f ([guid]::NewGuid().ToString("N")))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

try {
  $gcc     = Find-Tool @("x86_64-w64-mingw32-gcc.exe", "gcc.exe")
  $windres = Find-Tool @("x86_64-w64-mingw32-windres.exe", "windres.exe")
  $cl      = Find-Tool @("cl.exe")
  $rc      = Find-Tool @("rc.exe")

  if ($gcc -and $windres) {
    Write-Host "使用 MinGW-w64 编译…"

    $resObj = Join-Path $tmp "launcher_res.o"
    & $windres -I $here -i $rcFile -o $resObj
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
    & $rc /nologo /fo $resFile $rcFile
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

<#
.SYNOPSIS
  OneNote 导出工具 —— 图形界面（WPF）。

.DESCRIPTION
  面向不习惯命令行的用户：选笔记本 / 分区 / 导出目录，点「开始导出」即可。

  本脚本只负责界面、参数拼装与进度显示，导出逻辑不在这里重复实现 ——
  它把参数交给同目录的 sync-onenote.ps1，在子进程中执行，通过一个临时的
  JSON Lines 进度文件回传状态。

  之所以用子进程而不是把 sync-onenote.ps1 当函数加载：后者是脚本（顶层
  直接执行 + exit），dot-source 会当场跑完并结束整个进程，无法复用。
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

try {
  Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
  Add-Type -AssemblyName System.Windows.Forms
} catch {
  [System.Windows.MessageBox]::Show(
    "无法加载 WPF 界面组件，本工具的图形界面需要 Windows 自带的 .NET Framework。`n`n原始错误：$($_.Exception.Message)",
    "OneNote 导出工具", [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Error) | Out-Null
  exit 1
}

$script:scriptPath = Join-Path $PSScriptRoot "sync-onenote.ps1"
if (-not (Test-Path -LiteralPath $script:scriptPath)) {
  [System.Windows.MessageBox]::Show(
    "找不到 sync-onenote.ps1，请确认 scripts 目录完整。", "OneNote 导出工具",
    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
  exit 1
}

# 用与当前会话相同的 PowerShell 宿主去跑子进程，避免 pwsh / powershell 混用
$script:psExe = Join-Path $PSHOME "powershell.exe"
if (-not (Test-Path -LiteralPath $script:psExe)) { $script:psExe = Join-Path $PSHOME "pwsh.exe" }
if (-not (Test-Path -LiteralPath $script:psExe)) { $script:psExe = "powershell.exe" }

# ── 界面 ────────────────────────────────────────────────────────────────────

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="OneNote 导出工具" Width="660" SizeToContent="Height"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen"
        FontFamily="Microsoft YaHei UI" FontSize="13">
  <Grid Margin="18">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="4"
               Text="把本机 OneNote 笔记本导出成 Markdown 文件夹树，供知识库导入。"
               TextWrapping="Wrap" Foreground="#444444" Margin="0,0,0,16"/>

    <TextBlock Grid.Row="1" Grid.Column="0" Text="笔记本" VerticalAlignment="Center" Margin="0,0,12,0"/>
    <ComboBox x:Name="nbCombo" Grid.Row="1" Grid.Column="1" Height="28" Margin="0,0,10,0"/>
    <Button x:Name="btnRefresh" Grid.Row="1" Grid.Column="2" Grid.ColumnSpan="2"
            Content="重新读取" Width="90" Height="28" HorizontalAlignment="Right"/>

    <TextBlock Grid.Row="2" Grid.Column="0" Text="分区" VerticalAlignment="Center" Margin="0,10,12,0"/>
    <ComboBox x:Name="secCombo" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="3"
              Height="28" Margin="0,10,0,0"/>

    <TextBlock Grid.Row="3" Grid.Column="0" Text="导出到" VerticalAlignment="Center" Margin="0,10,12,0"/>
    <TextBox x:Name="outBox" Grid.Row="3" Grid.Column="1" Height="28" Margin="0,10,10,0"
             VerticalContentAlignment="Center"/>
    <Button x:Name="btnBrowse" Grid.Row="3" Grid.Column="2" Content="浏览…" Width="70" Height="28" Margin="0,10,0,0"/>
    <Button x:Name="btnOpen" Grid.Row="3" Grid.Column="3" Content="打开" Width="60" Height="28" Margin="8,10,0,0"/>

    <StackPanel Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="3" Orientation="Horizontal" Margin="0,14,0,0">
      <CheckBox x:Name="forceChk" Content="强制全量重导（忽略增量记录）" VerticalAlignment="Center"/>
      <CheckBox x:Name="skipInkChk" Content="跳过手绘渲染" Margin="24,0,0,0" VerticalAlignment="Center"/>
    </StackPanel>

    <ProgressBar x:Name="bar" Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="4"
                 Height="8" Margin="0,18,0,0" Minimum="0" Maximum="100" Value="0"/>

    <TextBlock x:Name="statusText" Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="4"
               Margin="0,8,0,0" Text="正在读取笔记本列表…" TextWrapping="Wrap"/>

    <ListBox x:Name="logList" Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="4"
             Height="180" Margin="0,10,0,0" FontFamily="Consolas, Microsoft YaHei UI" FontSize="12"
             ScrollViewer.HorizontalScrollBarVisibility="Auto"/>

    <StackPanel Grid.Row="8" Grid.Column="0" Grid.ColumnSpan="4" Orientation="Horizontal"
                HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="btnStart" Content="开始导出" Width="110" Height="32" IsDefault="True"/>
      <Button x:Name="btnClose" Content="退出" Width="80" Height="32" Margin="10,0,0,0" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
'@

$win = [System.Windows.Markup.XamlReader]::Parse($xaml)

foreach ($n in @("nbCombo","secCombo","outBox","btnBrowse","btnOpen","btnRefresh",
                 "forceChk","skipInkChk","bar","statusText","logList","btnStart","btnClose")) {
  Set-Variable -Name $n -Value $win.FindName($n) -Scope Script
}

# ── 状态 ────────────────────────────────────────────────────────────────────

$script:running       = $false
$script:proc          = $null
$script:progressPath  = $null
$script:seenLineCount = 0
$script:hierarchy     = @()
$script:terminalSeen  = $false
$script:outTask       = $null
$script:errTask       = $null

$script:tempDir = [System.IO.Path]::GetTempPath()

# 导出目录默认值：我的文档\OneNote导出
$outBox.Text = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "OneNote导出"

# ── 小工具 ──────────────────────────────────────────────────────────────────

function Add-Log {
  param([string]$Text, [bool]$IsWarning = $false)
  if ([string]::IsNullOrWhiteSpace($Text)) { return }
  $item = New-Object System.Windows.Controls.ListBoxItem
  $item.Content = $Text
  $item.Padding = [System.Windows.Thickness]::new(4, 1, 4, 1)
  if ($IsWarning) { $item.Foreground = [System.Windows.Media.Brushes]::Firebrick }
  $logList.Items.Add($item) | Out-Null
  # 大导出会有几千行，只保留最近的一批，避免界面越来越卡
  while ($logList.Items.Count -gt 1500) { $logList.Items.RemoveAt(0) }
  $logList.ScrollIntoView($logList.Items[$logList.Items.Count - 1])
}

function Set-Status {
  param([string]$Text, [bool]$IsError = $false)
  $statusText.Text = $Text
  $statusText.Foreground = if ($IsError) {
    [System.Windows.Media.Brushes]::Firebrick
  } else {
    [System.Windows.Media.Brushes]::Black
  }
}

function Set-UiRunning {
  param([bool]$Running)
  $script:running = $Running
  foreach ($c in @($nbCombo, $secCombo, $outBox, $btnBrowse, $btnRefresh, $forceChk, $skipInkChk)) {
    $c.IsEnabled = -not $Running
  }
  $btnStart.IsEnabled = -not $Running
  $btnStart.Content = if ($Running) { "导出中…" } else { "开始导出" }
}

# 强制把已排队的界面更新立刻画出来。UI 线程一旦阻塞在 WaitForExit 之类的
# 调用上，Dispatcher 就不再处理渲染，刚设的状态文字会停在上一帧不显示。
function Update-Ui {
  $win.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})
}

# 把参数数组拼成命令行字符串。带空格或引号的参数必须加引号，并按 Windows
# 的规则转义反斜杠（反斜杠只在其后紧跟引号时才需要翻倍，结尾的连续反斜杠
# 也要翻倍，否则会吃掉闭合引号）。
function ConvertTo-ArgString {
  param([string[]]$ArgumentList)
  $parts = foreach ($a in $ArgumentList) {
    if ($a -notmatch '[\s"]') { $a; continue }
    $e = [regex]::Replace($a, '(\\*)"', '$1$1\"')
    $e = [regex]::Replace($e, '(\\+)$', '$1$1')
    '"' + $e + '"'
  }
  ($parts -join ' ')
}

function New-ProgressFile {
  return (Join-Path $script:tempDir ("onenote-export-{0}.jsonl" -f ([guid]::NewGuid().ToString("N"))))
}

function Remove-ProgressFile {
  if ($script:progressPath -and (Test-Path -LiteralPath $script:progressPath)) {
    Remove-Item -LiteralPath $script:progressPath -Force -ErrorAction SilentlyContinue
  }
  $script:progressPath = $null
}

# 读取进度文件中尚未处理过的完整行。
# 每次都重读整个文件：文件很小（几千行封顶），换来的是对「写入写了一半」
# 的天然容忍 —— 末行不完整时直接丢弃，下一轮再取。
function Read-NewProgressEvents {
  if (-not $script:progressPath -or -not (Test-Path -LiteralPath $script:progressPath)) { return @() }
  $text = ""
  try {
    # 必须显式声明 FileShare.ReadWrite：子进程正以追加方式持有该文件
    $fs = [System.IO.File]::Open($script:progressPath, [System.IO.FileMode]::Open,
                                 [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
      $sr = New-Object System.IO.StreamReader($fs, [System.Text.UTF8Encoding]::new($false))
      try { $text = $sr.ReadToEnd() } finally { $sr.Dispose() }
    } finally { $fs.Dispose() }
  } catch { return @() }

  if ([string]::IsNullOrEmpty($text)) { return @() }

  # split 后末元素总是"不完整"的：结尾有换行时它是空串，没有换行时它是
  # 写了一半的行。两种情况下都丢弃，下一轮再取。
  $all = $text -split "`n"
  $complete = $all.Count - 1
  if ($complete -le $script:seenLineCount) { return @() }

  $events = @()
  for ($i = $script:seenLineCount; $i -lt $complete; $i++) {
    $line = $all[$i].Trim()
    if (-not $line) { continue }
    try { $events += ($line | ConvertFrom-Json) } catch { }
  }
  $script:seenLineCount = $complete
  return $events
}

function Invoke-ProgressEvent {
  param($Evt)
  switch ($Evt.phase) {
    "log"      { Add-Log $Evt.message }
    "log-warn" { Add-Log $Evt.message $true }
    "hierarchy" {
      $items = @($Evt.data)
      $script:hierarchy = @($items)
      $nbCombo.Items.Clear()
      $nbCombo.Items.Add("（全部笔记本）") | Out-Null
      foreach ($it in $items) { $nbCombo.Items.Add($it.name) | Out-Null }
      $nbCombo.SelectedIndex = 0
      Update-SectionCombo
    }
    "start" {
      $bar.Value = 0
      Set-Status $Evt.message
    }
    "page" {
      if ($Evt.total -gt 0) {
        $bar.Value = [Math]::Min(100, [Math]::Round(100 * $Evt.done / $Evt.total))
      }
      Set-Status ("正在导出 {0}/{1}：{2}" -f $Evt.done, $Evt.total, $Evt.message)
    }
    "up-to-date" {
      $bar.Value = 100
      $script:terminalSeen = $true
      Set-Status $Evt.message
    }
    "done" {
      $bar.Value = 100
      $script:terminalSeen = $true
      Set-Status ("导出完成 —— {0}" -f $Evt.message)
    }
    "error" {
      $script:terminalSeen = $true
      Set-Status $Evt.message $true
      Add-Log $Evt.message $true
    }
  }
}

function Update-SectionCombo {
  $secCombo.Items.Clear()
  $secCombo.Items.Add("（全部分区）") | Out-Null
  $idx = $nbCombo.SelectedIndex
  if ($idx -gt 0 -and $idx -le $script:hierarchy.Count) {
    $nb = $script:hierarchy[$idx - 1]
    foreach ($s in @($nb.sections)) {
      if (-not [string]::IsNullOrWhiteSpace($s)) { $secCombo.Items.Add($s) | Out-Null }
    }
  }
  $secCombo.SelectedIndex = 0
}

# ── 读取笔记本列表（同步执行，量小） ────────────────────────────────────────

function Invoke-RefreshNotebooks {
  if ($script:running) { return }

  $tmp = New-ProgressFile
  $win.Cursor = [System.Windows.Input.Cursors]::Wait
  $btnRefresh.IsEnabled = $false
  $script:terminalSeen = $false
  Set-Status "正在读取 OneNote 笔记本列表…"
  $logList.Items.Clear()
  Update-Ui

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:psExe
    $psi.Arguments = ConvertTo-ArgString @(
      "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $script:scriptPath,
      "-ListNotebooks", "-ProgressFile", $tmp)
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $psi.WorkingDirectory = $PSScriptRoot

    $p = [System.Diagnostics.Process]::Start($psi)
    # 异步读走两个流，再等退出 —— 直接 ReadToEnd 会有管道缓冲区写满的死锁风险
    $outTask = $p.StandardOutput.ReadToEndAsync()
    $errTask = $p.StandardError.ReadToEndAsync()
    $p.WaitForExit()
    $stdErr = $errTask.Result
    $outTask.Result | Out-Null

    $script:progressPath = $tmp
    $script:seenLineCount = 0
    foreach ($evt in (Read-NewProgressEvents)) { Invoke-ProgressEvent $evt }

    if ($p.ExitCode -ne 0 -or $script:hierarchy.Count -eq 0) {
      if ($stdErr) { Add-Log $stdErr.Trim() $true }
      if (-not $script:terminalSeen -and $script:hierarchy.Count -eq 0) {
        Set-Status "没能读到笔记本列表，详见下方日志。" $true
      }
    } else {
      Set-Status ("已读取 {0} 个笔记本。" -f $script:hierarchy.Count)
    }
  } catch {
    Set-Status ("读取笔记本列表失败：{0}" -f $_.Exception.Message) $true
  } finally {
    $script:progressPath = $null
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    $win.Cursor = [System.Windows.Input.Cursors]::Arrow
    $btnRefresh.IsEnabled = $true
  }
}

# ── 开始导出 ────────────────────────────────────────────────────────────────

function Start-Export {
  if ($script:running) { return }

  $outPath = $outBox.Text.Trim()
  if (-not $outPath) {
    Set-Status "请先选择导出目录。" $true
    return
  }

  $nb = ""
  if ($nbCombo.SelectedIndex -gt 0) { $nb = [string]$nbCombo.SelectedItem }
  $sec = ""
  if ($secCombo.SelectedIndex -gt 0) { $sec = [string]$secCombo.SelectedItem }

  $argList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $script:scriptPath,
               "-OutputPath", $outPath)
  if ($nb)  { $argList += @("-Notebook", $nb) }
  if ($sec) { $argList += @("-Section", $sec) }
  if ($forceChk.IsChecked)   { $argList += "-Force" }
  if ($skipInkChk.IsChecked) { $argList += "-SkipInkImages" }

  $script:progressPath = New-ProgressFile
  $script:seenLineCount = 0
  $script:terminalSeen = $false
  $script:outTask = $null
  $script:errTask = $null
  $argList += @("-ProgressFile", $script:progressPath)

  $logList.Items.Clear()
  $bar.Value = 0
  Set-UiRunning $true
  Set-Status "正在启动导出…"
  Update-Ui

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:psExe
    $psi.Arguments = ConvertTo-ArgString $argList
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $psi.WorkingDirectory = $PSScriptRoot

    $script:proc = [System.Diagnostics.Process]::Start($psi)

    # 必须异步读走两个流。子进程每导出一个页面就往 stdout 写一行，若没人读，
    # 管道缓冲区（约 4KB）写满后子进程会阻塞在写操作上，整个导出卡死。
    # 这里只用 ReadToEndAsync 把数据抽干、不显示它的内容 —— 界面日志统一由
    # 进度文件驱动，避免同一行在列表里出现两次。
    $script:outTask = $script:proc.StandardOutput.ReadToEndAsync()
    $script:errTask = $script:proc.StandardError.ReadToEndAsync()
  } catch {
    Set-Status ("无法启动导出进程：{0}" -f $_.Exception.Message) $true
    Set-UiRunning $false
    Remove-ProgressFile
  }
}

function Complete-Run {
  param([int]$ExitCode)

  $script:proc = $null

  if ($ExitCode -eq 0) {
    if (-not $script:terminalSeen) {
      # 例如「没有匹配的页面」这类提前 exit 0 的路径
      Set-Status "导出结束（无错误），详见下方日志。"
    }
  } else {
    # 兜底错误（trap 捕获的那种）只出现在 stderr，没有对应的 progress 事件
    $err = ""
    if ($script:errTask) { try { $err = $script:errTask.Result } catch { } }
    if ($err) { Add-Log $err.Trim() $true }
    if (-not $script:terminalSeen) {
      Set-Status ("导出未完成（退出码 {0}），详见下方日志。" -f $ExitCode) $true
    }
  }

  Set-UiRunning $false
  Remove-ProgressFile
  $script:outTask = $null
  $script:errTask = $null
}

# ── UI 计时器：唯一的界面更新入口 ───────────────────────────────────────────

function Update-FromProgress {
  if ($script:progressPath) {
    foreach ($evt in (Read-NewProgressEvents)) { Invoke-ProgressEvent $evt }
  }

  if ($script:running -and $script:proc -and $script:proc.HasExited) {
    # 子进程的写入在退出前已全部落盘，这里再收一次尾即可
    if ($script:progressPath) {
      foreach ($evt in (Read-NewProgressEvents)) { Invoke-ProgressEvent $evt }
    }
    Complete-Run $script:proc.ExitCode
  }
}

# ── 事件绑定 ────────────────────────────────────────────────────────────────

$nbCombo.Add_SelectionChanged({ Update-SectionCombo })

$btnRefresh.Add_Click({ Invoke-RefreshNotebooks })
$btnStart.Add_Click({ Start-Export })

$btnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = "选择导出目录"
  $dlg.ShowNewFolderButton = $true
  if ($outBox.Text -and (Test-Path -LiteralPath $outBox.Text)) { $dlg.SelectedPath = $outBox.Text }
  # 不传 owner：WPF 窗口与 WinForms 对话框的线程模型不同，挂错反而会崩，
  # 而本窗口是模态的，对话框不会跑到它后面去。
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $outBox.Text = $dlg.SelectedPath }
})

$btnOpen.Add_Click({
  $p = $outBox.Text.Trim()
  if (-not $p) { return }
  if (-not (Test-Path -LiteralPath $p)) {
    try { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    catch { Set-Status ("无法创建目录：{0}" -f $_.Exception.Message) $true; return }
  }
  Start-Process explorer.exe -ArgumentList "`"$p`""
})

$btnClose.Add_Click({ $win.Close() })

$win.Add_Closing({
  # PowerShell 把处理脚本块接到 .NET 委托上时，事件参数可能出现在位置参数
  # （$args[1]）里，也可能只暴露成自动变量 $EventArgs —— 两种绑定都兜住，
  # 否则「取消关闭」会静默失效，用户点「否」窗口还是关了。
  $e = if ($args.Count -ge 2) { $args[1] } else { $EventArgs }

  if ($script:running -and $script:proc -and -not $script:proc.HasExited) {
    $r = [System.Windows.MessageBox]::Show(
      "导出还在进行中。现在关闭会中断导出（已写出的页面会保留，但这次不会记入增量记录，下次需要重导）。`n`n确定要关闭吗？",
      "OneNote 导出工具", [System.Windows.MessageBoxButton]::YesNo,
      [System.Windows.MessageBoxImage]::Warning)
    if ($r -ne [System.Windows.MessageBoxResult]::Yes) {
      if ($e) { $e.Cancel = $true }
      return
    }
    try { $script:proc.Kill() } catch { }
  }
})

$win.Add_Closed({
  if ($script:proc -and -not $script:proc.HasExited) {
    try { $script:proc.Kill() } catch { }
  }
  Remove-ProgressFile
})

# 单实例：两个界面同时导出会争抢同一个增量记录文件。
# 放在 ShowDialog 之前：抢不到就直接退出，不白开一个窗口。
$script:mutex = New-Object System.Threading.Mutex($false, "Local\OneNoteExportToolGui")
$gotMutex = $false
try { $gotMutex = $script:mutex.WaitOne(0, $false) } catch { $gotMutex = $true }
if (-not $gotMutex) {
  [System.Windows.MessageBox]::Show(
    "导出工具已经在运行了。", "OneNote 导出工具",
    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
  exit 0
}

$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds(300)
# 计时器回调里抛异常会一路冒到 Dispatcher，直接崩掉界面，所以整块兜住
$script:timer.Add_Tick({
  try { Update-FromProgress }
  catch { try { Add-Log ("界面刷新出错：" + $_.Exception.Message) $true } catch { } }
})
$script:timer.Start()

# 窗口画出来之后再拉笔记本列表，避免启动时白屏
$win.Add_Loaded({
  $script:loadTimer = New-Object System.Windows.Threading.DispatcherTimer
  $script:loadTimer.Interval = [TimeSpan]::FromMilliseconds(150)
  $script:loadTimer.Add_Tick({
    $script:loadTimer.Stop()
    Invoke-RefreshNotebooks
  })
  $script:loadTimer.Start()
})

$win.ShowDialog() | Out-Null

$script:timer.Stop()
if ($script:mutex) { try { $script:mutex.ReleaseMutex() } catch { } }

<#
.SYNOPSIS
  从本机 OneNote 桌面版导出笔记为 Markdown 文件夹树（便于导入各类知识库）。

.DESCRIPTION
  通过 OneNote COM 接口直接读取本机已同步的笔记本，按
  <输出目录>/<笔记本>/<分区组>/<分区>/<页面>.md 的结构落盘。
  子页面（pageLevel > 1）会成为父页面同名的子文件夹。

  图片内联为 data URI（base64），不产生旁挂图片目录 —— 因为下游知识库
  按「每个 .md 一个条目」入库，相对路径引用的兄弟图片文件不会被关联。

.PARAMETER OutputPath
  导出根目录。不存在时会创建。必须是空目录或专用目录。

.PARAMETER Notebook
  只导出指定笔记本（按名称精确匹配）。不指定则导出全部。

.PARAMETER Section
  只导出指定分区（按名称精确匹配）。可与 -Notebook 组合。

.PARAMETER Force
  忽略增量记录，强制重新导出全部页面。

.PARAMETER List
  只列出会导出哪些页面，不写任何文件。

.PARAMETER ListNotebooks
  只列出本机所有笔记本（含分区数、页面数），不写任何文件。
  用于在不知道笔记本确切名称时，先查名称再传给 -Notebook。

.PARAMETER Help
  打印参数说明与常用示例。

.PARAMETER SkipInkImages
  跳过手绘（InkDrawing）的图片渲染。手写文字的识别文本仍会提取。

.PARAMETER InkToleranceRatio
  ink 归属判定的横向容差，以容器宽度的比例表示（默认 0.5 = 半宽）。
  标注常画到容器边界外，容差用于把这类溢出笔迹拉回所属容器。

.PARAMETER InkGapThreshold
  纯 ink 聚类的纵向间距阈值（页面坐标，默认 50）。
  相邻笔画间距超过它就被切成两块，各成一张图。笔记写得越疏，值要越大。

.PARAMETER DumpXml
  诊断用：把每个页面的原始 OneNote XML 写到 <输出目录>/_xml/<页面>.xml。
  仅在排查坐标/ink/图片解析问题时使用 —— 这些文件不是给知识库导入的。

.PARAMETER ProgressFile
  内部参数，供 GUI 使用：把导出进度以 JSON Lines 追加写入该文件。
  命令行使用时不必传。

.EXAMPLE
  .\sync-onenote.ps1 -List
  .\sync-onenote.ps1 -OutputPath D:\onenote-export -Notebook "工作"
  .\sync-onenote.ps1 -OutputPath D:\onenote-export -Force
  .\sync-onenote.ps1 -OutputPath D:\onenote-xml -DumpXml -Notebook "学习笔记"
#>
[CmdletBinding()]
param(
  [string]$OutputPath,
  [string]$Notebook,
  [string]$Section,
  [switch]$Force,
  [switch]$List,
  [switch]$ListNotebooks,
  [switch]$Help,
  [switch]$SkipInkImages,
  [switch]$DumpXml,
  [string]$ProgressFile,
  [double]$InkToleranceRatio = 0.5,
  [double]$InkGapThreshold = 50
)

$ErrorActionPreference = "Stop"

# PowerShell 5.1 默认按系统 ANSI 代码页读取无 BOM 的 .ps1，中文会乱码。
# 本脚本以 UTF-8 BOM 保存，同时显式设定控制台输出编码。
try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
} catch { }

if ($Help) {
  Write-Host @"
OneNote 导出工具 —— 把本机 OneNote 笔记本导出为 Markdown 文件夹树

用法:
  sync-onenote.cmd [参数]

常用:
  -ListNotebooks           列出本机所有笔记本（名称、分区数、页面数），不写文件
  -List                    列出会被导出的页面与相对路径，不写文件
  -OutputPath <目录>       导出根目录，不存在则创建（必填，除非用上面两个 -List*）
  -Notebook <名称>         只导出指定笔记本（名称精确匹配，先用 -ListNotebooks 查）
  -Section <名称>          只导出指定分区，可与 -Notebook 组合
  -Force                   忽略增量记录，强制全部重导

可选:
  -SkipInkImages           跳过手绘渲染（手写识别文本仍会提取）
  -InkToleranceRatio <n>   ink 归属容差，相对容器宽度的比例，默认 0.5（半宽）
  -InkGapThreshold <n>     纯 ink 聚类的纵向间距阈值（页面坐标），默认 50
  -DumpXml                 诊断用：把原始页面 XML 写到 <输出目录>/_xml/
  -Help                    显示本帮助

示例:
  .\scripts\sync-onenote.cmd -ListNotebooks
  .\scripts\sync-onenote.cmd -List -Notebook "工作"
  .\scripts\sync-onenote.cmd -OutputPath D:\onenote-export -Notebook "工作"
  .\scripts\sync-onenote.cmd -OutputPath D:\onenote-export -Section "会议记录" -Force

环境要求:
  Windows 10/11 + OneNote 桌面版（Microsoft Store 版无自动化接口，不可用）
  + Windows PowerShell 5.1（系统自带）

输出结构:
  <输出目录>/<笔记本>/<分区组>/<分区>/<页面>.md
  子页面（pageLevel > 1）成为父页面同名的子文件夹。
  图片内联为 data URI，手写识别文本以正文形式插入，手绘图渲染成 PNG。

导入知识库:
  推荐用知识库的「上传文件夹」入口选导出目录，这样能保留目录层级。
  用命令行工具逐个文件上传时，部分实现不构造相对路径，层级会丢失。
"@
  exit 0
}


# 下游知识库对文件夹路径的常见限制：
#   单个路径段 <= 128 字节、整条路径 <= 1024 字节、层级 <= 16。
# 中文一段 3 字节，128 字节约 42 个汉字，因此按字符数保守截断。
$MaxSegmentChars = 80

$manifestPath = Join-Path $PSScriptRoot ".onenote-export-manifest.json"

function Write-Step([string]$Message) {
  Write-Host $Message
  Write-ProgressEvent -Phase "log" -Done 0 -Total 0 -Message $Message
}
function Write-Warn2([string]$Message) {
  Write-Warning $Message
  Write-ProgressEvent -Phase "log-warn" -Done 0 -Total 0 -Message $Message
}

function Get-SafeSegment([string]$Name) {
  if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
  $safe = ($Name -replace '[<>:"/\\|?*\x00-\x1f]', "_").Trim()
  $safe = $safe.TrimEnd('.', ' ')
  if ($safe.Length -gt $MaxSegmentChars) {
    $safe = $safe.Substring(0, $MaxSegmentChars).Trim()
  }
  return $safe
}

function Get-SafeFileName([string]$Title) {
  $safe = Get-SafeSegment $Title
  if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "无标题页面" }
  return $safe
}

function Clean-HtmlText([string]$html) {
  if ([string]::IsNullOrWhiteSpace($html)) { return "" }
  $s = [System.Net.WebUtility]::HtmlDecode($html)
  $s = $s -replace "(?i)<br\s*/?>", "`n"
  $s = $s -replace "(?i)<span\s+style=['""][^'""]*font-weight\s*:\s*bold[^'""]*['""]>(.*?)</span>", '**$1**'
  $s = $s -replace "(?i)<span\s+style=['""][^'""]*font-style\s*:\s*italic[^'""]*['""]>(.*?)</span>", '*$1*'
  $s = $s -replace "(?i)<b>(.*?)</b>", '**$1**'
  $s = $s -replace "(?i)<i>(.*?)</i>", '*$1*'
  $s = $s -replace "(?s)<[^>]+>", ""
  return $s.Trim()
}

# ── 图片处理 ────────────────────────────────────────────────────────────────

# EMF/WMF 是矢量格式，下游知识库的 MIME 映射表通常没有 emf 分支，
# 直接内联会静默丢图，因此必须在导出侧转成位图。
# 返回 byte[] 时必须用 unary comma（",$bytes"）：PowerShell 的输出管道会枚举
# IEnumerable，裸 `return $bytes` 会把字节数组拆成一个一个 byte，调用方拿到
# 的就不再是数组。这里不能像集合那样改填充式 —— 它只有一个返回值。
function Convert-MetafileToPngBytes {
  param([byte[]]$Bytes)
  Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  $inStream = New-Object System.IO.MemoryStream -ArgumentList @(,$Bytes)
  try {
    $meta = [System.Drawing.Image]::FromStream($inStream)
    try {
      $bmp = New-Object System.Drawing.Bitmap -ArgumentList @([int]$meta.Width, [int]$meta.Height)
      try {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
          $g.Clear([System.Drawing.Color]::White)
          $g.DrawImage($meta, 0, 0, [int]$meta.Width, [int]$meta.Height)
        } finally { $g.Dispose() }
        $outStream = New-Object System.IO.MemoryStream
        try {
          $bmp.Save($outStream, [System.Drawing.Imaging.ImageFormat]::Png)
          return , $outStream.ToArray()
        } finally { $outStream.Dispose() }
      } finally { $bmp.Dispose() }
    } finally { $meta.Dispose() }
  } finally { $inStream.Dispose() }
}

# OneNote 的 Image/@format 有时是 "auto"，此时按魔术字节判断。
function Get-ImageMimeFromBytes {
  param([byte[]]$Bytes, [string]$Format)
  $fmt = ""
  if ($Format) { $fmt = $Format.ToLowerInvariant() }
  if ($fmt -in @("png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp")) {
    if ($fmt -eq "jpg") { return "image/jpeg" }
    if ($fmt -in @("tif", "tiff")) { return "image/tiff" }
    return "image/$fmt"
  }
  if ($Bytes.Length -ge 8 -and
      $Bytes[0] -eq 0x89 -and $Bytes[1] -eq 0x50 -and $Bytes[2] -eq 0x4E -and $Bytes[3] -eq 0x47) {
    return "image/png"
  }
  if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xD8 -and $Bytes[2] -eq 0xFF) {
    return "image/jpeg"
  }
  if ($Bytes.Length -ge 6 -and $Bytes[0] -eq 0x47 -and $Bytes[1] -eq 0x49 -and $Bytes[2] -eq 0x46) {
    return "image/gif"
  }
  if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0x42 -and $Bytes[1] -eq 0x4D) {
    return "image/bmp"
  }
  return "image/png"
}

function Test-IsMetafileFormat {
  param([byte[]]$Bytes, [string]$Format)
  if ($Format -and ($Format.ToLowerInvariant() -in @("emf", "wmf"))) { return $true }
  # EMF 记录头：类型 1(EMR_HEADER) 或 0x00010000 (EMF+)
  if ($Bytes.Length -ge 44 -and $Bytes[0] -eq 0x01 -and $Bytes[1] -eq 0x00 -and $Bytes[2] -eq 0x00 -and $Bytes[3] -eq 0x00) {
    return $true
  }
  # 旧式 WMF（placeable）：D7 CD C6 9A
  if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0xD7 -and $Bytes[1] -eq 0xCD -and $Bytes[2] -eq 0xC6 -and $Bytes[3] -eq 0x9A) {
    return $true
  }
  return $false
}

# 把 <one:Image> / <one:InkDrawing> 的字节转成可直接内联的 data URI。
# 返回 $null 表示这张图应当被跳过（需调用方打印警告）。
function Convert-ImageBytesToDataUri {
  param([byte[]]$Bytes, [string]$Format, [string]$Context)
  if (-not $Bytes -or $Bytes.Length -eq 0) { return $null }
  $bytesToUse = $Bytes
  if (Test-IsMetafileFormat -Bytes $Bytes -Format $Format) {
    try {
      $bytesToUse = Convert-MetafileToPngBytes -Bytes $Bytes
      $Format = "png"
    } catch {
      Write-Warn2 "  [$Context] EMF/WMF 转 PNG 失败，已跳过该图：$($_.Exception.Message)"
      return $null
    }
  }
  $mime = Get-ImageMimeFromBytes -Bytes $bytesToUse -Format $Format
  return "data:$mime;base64," + [Convert]::ToBase64String($bytesToUse)
}

# ── 手绘（ink）处理 ────────────────────────────────────────────────────────

# OneNote 页面 XML 里的 ink 有三类元素：
#   InkWord      —— 一个个手写词，@recognizedText 是 OneNote 自己的识别结果
#   InkParagraph —— 装 InkWord 的容器（倾斜手写）
#   InkDrawing   —— 手绘图/涂鸦，只有 ISF 笔画数据，没有任何文本
# 前者的文本对检索极有价值，而微软的文档完全没记录这个属性（只能从 XSD
# 或实测 XML 里发现）。后者要变成可看的图，必须自己渲染。

function Get-InkRecognizedText {
  param($Scope, [System.Xml.XmlNamespaceManager]$NsManager)
  $sb = [System.Text.StringBuilder]::new()

  # 同 Add-InkBlobsFrom：".//one:InkWord" 不匹配上下文节点自身。
  # 传入单个 InkWord 时要单独处理，否则取不到它的 recognizedText。
  $candidates = [System.Collections.Generic.List[object]]::new()
  if ($Scope.LocalName -eq "InkWord") { $candidates.Add($Scope) }
  foreach ($w in $Scope.SelectNodes(".//one:InkWord", $NsManager)) { $candidates.Add($w) }

  foreach ($word in $candidates) {
    $recognized = $word.GetAttribute("recognizedText")
    if (-not [string]::IsNullOrWhiteSpace($recognized)) {
      [void]$sb.Append($recognized)
      continue
    }
    # 没有识别文本的 InkWord 只可能是分隔符（空格 / 换行）
    if ($word.SelectSingleNode("one:Space", $NsManager)) {
      [void]$sb.Append(" ")
    } elseif ($word.SelectSingleNode("one:EndOfLine", $NsManager)) {
      [void]$sb.Append("`n")
    }
  }
  return $sb.ToString().Trim()
}


# ── 页面几何：块提取、ink 聚类、归属判定 ────────────────────────────────────
#
# 渲染 ink 用的是 WPF（System.Windows.Ink.StrokeCollection），不是 Tablet PC
# 的 Microsoft.Ink COM 类 —— 后者在 Windows 11 上默认不再注册，New-Object
# 会抛 80040154 REGDB_E_CLASSNOTREG（CLSID 全零 = 连 ProgID 都没解析出来）。
# PresentationCore 属于 .NET Framework，Windows 上必然存在；代价是必须跑在
# STA 线程上（powershell.exe 控制台宿主默认就是 STA，该约束自动满足）。

# 页面上一个带坐标的"块"：Outline（文本流，可能内含图片）或页面级 InkDrawing。
#
# 填充式（写入调用方给的 $Target）而不是返回值。PowerShell 的输出管道会递归
# 展平 IEnumerable，`return $list` 会把 List[T] 拆成一个个 T 甚至单个字段；
# 集合一律用填充式，避免依赖调用方记得写 ,$list 或 @()。
function Add-PageBlocks {
  param(
    [xml]$PageDoc,
    [System.Xml.XmlNamespaceManager]$NsManager,
    [System.Collections.Generic.List[PSCustomObject]]$Target
  )

  foreach ($child in $PageDoc.DocumentElement.ChildNodes) {
    if ($child.LocalName -notin @("Outline", "Image", "InsertedFile", "InkDrawing", "InkParagraph")) {
      continue
    }
    $pos = $child.SelectSingleNode("one:Position", $NsManager)
    $size = $child.SelectSingleNode("one:Size", $NsManager)
    $x = if ($pos -and $pos.x) { [double]$pos.x } else { 0.0 }
    $y = if ($pos -and $pos.y) { [double]$pos.y } else { 0.0 }
    $w = if ($size -and $size.width) { [double]$size.width } else { 0.0 }
    $h = if ($size -and $size.height) { [double]$size.height } else { 0.0 }
    $Target.Add([PSCustomObject]@{
        Node = $child
        Tag  = $child.LocalName
        X0   = $x
        Y0   = $y
        X1   = $x + $w
        Y1   = $y + $h
      })
  }
}

# 把页面级 ink 按纵向邻近度聚类。
#
# 比较基准是"该簇迄今的最大 bottom"，不是"最后一笔的 bottom" —— 笔画在
# OneNote 里可能不按 y 顺序排列，用最后一笔会把本该同簇的笔画切断
# （实测：红标注间隔 38.7px 却因为中间夹了一笔更高的笔画而被误分）。
function Add-InkGroups {
  param(
    [object[]]$InkBlocks,
    [double]$GapThreshold,
    [System.Collections.Generic.List[object]]$Target
  )
  if (-not $InkBlocks -or $InkBlocks.Count -eq 0) { return }

  $sorted = @($InkBlocks | Sort-Object Y0)
  $current = [System.Collections.Generic.List[object]]::new()
  $currentBottom = [double]::NegativeInfinity

  foreach ($b in $sorted) {
    if ($current.Count -gt 0 -and ($b.Y0 - $currentBottom) -gt $GapThreshold) {
      $Target.Add($current.ToArray())   # 存数组：List 本身也是 IEnumerable，嵌套时会一并被展平
      $current = [System.Collections.Generic.List[object]]::new()
      $currentBottom = [double]::NegativeInfinity
    }
    $current.Add($b)
    if ($b.Y1 -gt $currentBottom) { $currentBottom = $b.Y1 }
  }
  if ($current.Count -gt 0) { $Target.Add($current.ToArray()) }
}

# 一组块的并集包围盒。
function Get-BoundsUnion {
  param([object[]]$Blocks)
  return [PSCustomObject]@{
    X0 = ($Blocks | Measure-Object -Property X0 -Minimum).Minimum
    Y0 = ($Blocks | Measure-Object -Property Y0 -Minimum).Minimum
    X1 = ($Blocks | Measure-Object -Property X1 -Maximum).Maximum
    Y1 = ($Blocks | Measure-Object -Property Y1 -Maximum).Maximum
  }
}

# ink 簇是否属于某个 Outline：y 与 x 都要重叠。
# x 方向额外给 tolX 容差，因为标注常常画到容器边界外（实测溢出约 10px，
# 而默认容差是容器半宽，宁可多捕）。容差只放宽"重叠"判定，不影响渲染范围。
function Test-InkGroupBelongsToBlock {
  param(
    [object]$Bounds,
    [object]$Block,
    [double]$TolX
  )
  $yOverlap = $Bounds.Y0 -lt $Block.Y1 -and $Bounds.Y1 -gt $Block.Y0
  $xOverlap = $Bounds.X0 -lt ($Block.X1 + $TolX) -and $Bounds.X1 -gt ($Block.X0 - $TolX)
  return ($yOverlap -and $xOverlap)
}


# 从一个元素或块里取出所有 ISF blob（InkDrawing 与 InkWord 的 Data），
# 追加到 $Target。填充式的原因见 Add-PageBlocks：返回值会被管道展平，
# 而 byte[] 是 IEnumerable —— `return $blobs` 会把每个 blob 拆成单个 byte，
# 表现为 StrokeCollection 报几百条"1 字节 ISF 解析失败"。
function Add-InkBlobsFrom {
  param(
    $Scope,
    [System.Xml.XmlNamespaceManager]$NsManager,
    [System.Collections.Generic.List[byte[]]]$Target
  )

  # 上下文节点（$Scope）自身可能就是 InkDrawing / InkWord。XPath 的 ".//x"
  # 展开为 descendant-or-self::node()/child::x —— 只匹配后代、不匹配自身，
  # 而 InkDrawing 内部不会嵌套 InkDrawing，所以不单独处理就永远取不到数据
  # （这正是"页面全黑无 ink"的成因）。
  if ($Scope.LocalName -in @("InkDrawing", "InkWord")) {
    $selfData = $Scope.SelectSingleNode("one:Data", $NsManager)
    if ($selfData -and -not [string]::IsNullOrWhiteSpace($selfData.InnerText)) {
      try { $Target.Add([Convert]::FromBase64String($selfData.InnerText.Trim())) } catch { }
    }
  }

  foreach ($node in $Scope.SelectNodes(".//one:InkDrawing | .//one:InkWord", $NsManager)) {
    $dataNode = $node.SelectSingleNode("one:Data", $NsManager)
    if (-not $dataNode -or [string]::IsNullOrWhiteSpace($dataNode.InnerText)) { continue }
    try { $Target.Add([Convert]::FromBase64String($dataNode.InnerText.Trim())) } catch { continue }
  }
}

# 把一个块（image 字节列表 + ink blob 列表）合成渲染成一张 PNG。
# 图与 ink 在同一个 DrawingVisual 里叠加，位置关系得以保留。
# 图片按显示尺寸（不是原始像素尺寸）绘制，与 OneNote 的布局一致。
# 返回 byte[]，必须用 unary comma（原因见 Convert-MetafileToPngBytes）。
function New-CompositeBlockPng {
  param(
    [System.Collections.Generic.List[object]]$Images,
    [System.Collections.Generic.List[byte[]]]$InkBlobs,
    # ink 在页面坐标系里的目标矩形（来自 XML 的 Position/Size）。提供时会把
    # ISF 自身的包围盒仿射映射到它 —— ISF 内部坐标与页面坐标不是 1:1
    # （实测比值约 1.35），不映射就会与图片错位。缺省时按 ISF 原始坐标绘制。
    [object]$InkTargetRect
  )

  if (($Images.Count + $InkBlobs.Count) -eq 0) { return $null }
  Add-Type -AssemblyName WindowsBase, PresentationCore, System.Drawing -ErrorAction Stop

  # 先算出所有内容的页面坐标包围盒，据此决定画布大小与平移。
  $ink = New-Object System.Windows.Ink.StrokeCollection
  foreach ($bytes in $InkBlobs) {
    $ms = New-Object System.IO.MemoryStream -ArgumentList @(,$bytes)
    try {
      $part = New-Object System.Windows.Ink.StrokeCollection -ArgumentList @($ms)
      if ($part.Count -gt 0) { $ink.Add($part) }
    } catch {
      # ISF 解不出笔画：记下但不中断（其它 blob 仍要渲染）
      Write-Warn2 ("    ISF 笔画解析失败（{0} 字节）：{1}" -f $bytes.Length, $_.Exception.Message)
    } finally { $ms.Dispose() }
  }

  $minX = [double]::PositiveInfinity; $minY = [double]::PositiveInfinity
  $maxX = [double]::NegativeInfinity; $maxY = [double]::NegativeInfinity

  foreach ($img in $Images) {
    if ($img.X0 -lt $minX) { $minX = $img.X0 }
    if ($img.Y0 -lt $minY) { $minY = $img.Y0 }
    if ($img.X1 -gt $maxX) { $maxX = $img.X1 }
    if ($img.Y1 -gt $maxY) { $maxY = $img.Y1 }
  }
  $inkBounds = if ($ink.Count -gt 0) { $ink.GetBounds() } else { $null }
  # 参与画布计算的必须是"映射后的页面坐标"，否则 ISF 原始尺寸会撑大画布。
  if ($InkTargetRect) {
    if ($InkTargetRect.X0 -lt $minX) { $minX = $InkTargetRect.X0 }
    if ($InkTargetRect.Y0 -lt $minY) { $minY = $InkTargetRect.Y0 }
    if ($InkTargetRect.X1 -gt $maxX) { $maxX = $InkTargetRect.X1 }
    if ($InkTargetRect.Y1 -gt $maxY) { $maxY = $InkTargetRect.Y1 }
  } elseif ($inkBounds) {
    if ($inkBounds.Left -lt $minX) { $minX = $inkBounds.Left }
    if ($inkBounds.Top -lt $minY) { $minY = $inkBounds.Top }
    if ($inkBounds.Right -gt $maxX) { $maxX = $inkBounds.Right }
    if ($inkBounds.Bottom -gt $maxY) { $maxY = $inkBounds.Bottom }
  }
  if ([double]::IsInfinity($minX)) { return $null }

  $margin = 8.0
  $scale = 2.0   # 放大一倍，细笔画与文字对 VLM 更友好
  $w = [int][Math]::Max(1, [Math]::Ceiling(($maxX - $minX + $margin * 2) * $scale))
  $h = [int][Math]::Max(1, [Math]::Ceiling(($maxY - $minY + $margin * 2) * $scale))
  # 画布上限保护：坐标异常（如 Inf/NaN）会算出天文数字，RenderTargetBitmap
  # 在这种尺寸上会抛 OutOfMemory 而不是给出可诊断的错误。
  $maxDim = 20000
  if ($w -gt $maxDim -or $h -gt $maxDim) {
    throw "合成画布过大（${w}x${h}，上限 ${maxDim}）：坐标可能异常（minX=$minX minY=$minY maxX=$maxX maxY=$maxY）"
  }

  $visual = New-Object System.Windows.Media.DrawingVisual
  $dc = $visual.RenderOpen()
  try {
    $dc.DrawRectangle([System.Windows.Media.Brushes]::White, $null,
      (New-Object System.Windows.Rect(0, 0, $w, $h)))
    $dc.PushTransform((New-Object System.Windows.Media.ScaleTransform($scale, $scale)))
    $dc.PushTransform((New-Object System.Windows.Media.TranslateTransform(
      (-$minX + $margin), (-$minY + $margin))))

    foreach ($img in $Images) {
      if (-not $img.Bitmap) { continue }
      $dc.DrawImage($img.Bitmap,
        (New-Object System.Windows.Rect($img.X0, $img.Y0, ($img.X1 - $img.X0), ($img.Y1 - $img.Y0))))
    }
    if ($ink.Count -gt 0) {
      if ($InkTargetRect -and $inkBounds -and $inkBounds.Width -gt 0 -and $inkBounds.Height -gt 0) {
        # ISF 包围盒 → XML 目标矩形 的仿射映射（先平移到原点，再缩放到目标尺寸，
        # 最后平移到目标位置）。OneNote 自己就是这么把笔画放进页面坐标的。
        $sx = ($InkTargetRect.X1 - $InkTargetRect.X0) / $inkBounds.Width
        $sy = ($InkTargetRect.Y1 - $InkTargetRect.Y0) / $inkBounds.Height
        $m = [System.Windows.Media.Matrix]::Identity
        $m.Scale($sx, $sy)
        $m.Translate(($InkTargetRect.X0 - $inkBounds.Left * $sx), ($InkTargetRect.Y0 - $inkBounds.Top * $sy))
        $dc.PushTransform((New-Object System.Windows.Media.MatrixTransform($m)))
        $ink.Draw($dc)
        $dc.Pop()
      } else {
        $ink.Draw($dc)
      }
    }

    $dc.Pop(); $dc.Pop()
  } finally { $dc.Close() }

  $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
    $w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
  $rtb.Render($visual)

  $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
  $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
  $outStream = New-Object System.IO.MemoryStream
  try {
    $encoder.Save($outStream)
    return , $outStream.ToArray()
  } finally { $outStream.Dispose() }
}

# 从 <one:Image> 节点取显示尺寸与原始字节。
function Get-ImageNodeInfo {
  param($Node, [System.Xml.XmlNamespaceManager]$NsManager)
  $dataNode = $Node.SelectSingleNode("one:Data", $NsManager)
  if (-not $dataNode -or [string]::IsNullOrWhiteSpace($dataNode.InnerText)) { return $null }
  try {
    $bytes = [Convert]::FromBase64String($dataNode.InnerText.Trim())
  } catch { return $null }
  $size = $Node.SelectSingleNode("one:Size", $NsManager)
  $w = if ($size -and $size.width) { [double]$size.width } else { 0.0 }
  $h = if ($size -and $size.height) { [double]$size.height } else { 0.0 }
  $fmt = if ($Node.format) { $Node.format } else { "auto" }
  return [PSCustomObject]@{ Bytes = $bytes; Width = $w; Height = $h; Format = $fmt }
}

# ── 页面 XML → Markdown ────────────────────────────────────────────────────

# WPF 位图（用于合成渲染时把图片画进 DrawingVisual）。
# EMF/WMF 无法被 BitmapImage 直接加载，先栅格化成 PNG（复用同一转换器）。
function New-WpfBitmapFromBytes {
  param([byte[]]$Bytes, [string]$Format)
  if (-not $Bytes -or $Bytes.Length -eq 0) { return $null }
  # 必须在本函数内加载：BitmapImage 属于 PresentationCore，而本函数会在
  # New-CompositeBlockPng 之前被调用，后者的 Add-Type 那时还没执行。
  Add-Type -AssemblyName WindowsBase, PresentationCore -ErrorAction Stop
  $work = $Bytes
  if (Test-IsMetafileFormat -Bytes $Bytes -Format $Format) {
    try { $work = Convert-MetafileToPngBytes -Bytes $Bytes } catch {
      Write-Warn2 ("    EMF/WMF 转 PNG 失败：{0}" -f $_.Exception.Message)
      return $null
    }
  }
  $ms = New-Object System.IO.MemoryStream -ArgumentList @(,$work)
  try {
    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bmp.StreamSource = $ms
    $bmp.EndInit()
    $bmp.Freeze()
    return $bmp
  } catch {
    # 不静默：图片加载失败会让"底图消失"这种症状难以定位
    Write-Warn2 ("    图片解码为 WPF 位图失败：{0} @行{1} —— {2}" -f `
      $_.Exception.GetType().Name, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    return $null
  } finally { $ms.Dispose() }
}

function Convert-PageXmlToMarkdown {
  param(
    [xml]$PageDoc,
    [System.Xml.XmlNamespaceManager]$NsManager,
    [string]$PageTitle,
    [string]$BaseFileName,
    [switch]$SkipInkImages,
    [double]$InkToleranceRatio = 0.5,
    [double]$InkGapThreshold = 50
  )

  $title = $PageDoc.DocumentElement.name
  if (-not $title) { $title = $PageTitle }

  $lines = [System.Collections.Generic.List[string]]::new()
  $lines.Add("# $title")
  $lines.Add("")

  $blocks = [System.Collections.Generic.List[PSCustomObject]]::new()
  Add-PageBlocks -PageDoc $PageDoc -NsManager $NsManager -Target $blocks

  # ── 1. ink 聚类 ──
  # 先聚类再判归属（顺序不能反）：聚类把"一次手写"还原成一个整体，归属判定
  # 只需整体与容器求交，容差从"必须够大才捕获得到"降级为兜底。
  $inkGroups = [System.Collections.Generic.List[object]]::new()
  if (-not $SkipInkImages) {
    $inkBlocks = @($blocks | Where-Object { $_.Tag -eq "InkDrawing" -or $_.Tag -eq "InkParagraph" })
    Add-InkGroups -InkBlocks $inkBlocks -GapThreshold $InkGapThreshold -Target $inkGroups
  }
  $inkAssigned = New-Object 'bool[]' ([Math]::Max(1, $inkGroups.Count))

  # ── 2. ink 簇归属 Outline ──
  # 判据 y 与 x 双向重叠，x 给容器半宽容差（标注常画到容器边界外）。
  $outlineGroups = @{}
  foreach ($ob in @($blocks | Where-Object { $_.Tag -eq "Outline" })) {
    $tolX = [Math]::Abs($ob.X1 - $ob.X0) * $InkToleranceRatio
    for ($gi = 0; $gi -lt $inkGroups.Count; $gi++) {
      if ($inkAssigned[$gi]) { continue }
      $bounds = Get-BoundsUnion -Blocks $inkGroups[$gi]
      if (Test-InkGroupBelongsToBlock -Bounds $bounds -Block $ob -TolX $tolX) {
        $key = $ob.Node.objectID
        if (-not $outlineGroups.ContainsKey($key)) { $outlineGroups[$key] = @() }
        $outlineGroups[$key] += , $inkGroups[$gi]
        $inkAssigned[$gi] = $true
      }
    }
  }

  # ── 3. 组装渲染单元并按 y 排序 ──
  $units = [System.Collections.Generic.List[PSCustomObject]]::new()
  foreach ($b in $blocks) {
    if ($b.Tag -in @("Outline", "Image")) {
      $units.Add([PSCustomObject]@{ Y0 = $b.Y0; Kind = "Block"; Block = $b })
    }
  }
  for ($gi = 0; $gi -lt $inkGroups.Count; $gi++) {
    if ($inkAssigned[$gi]) { continue }
    $bounds = Get-BoundsUnion -Blocks $inkGroups[$gi]
    $units.Add([PSCustomObject]@{ Y0 = $bounds.Y0; Kind = "InkGroup"; Group = $inkGroups[$gi] })
  }

  $imgIndex = 0

  foreach ($unit in ($units | Sort-Object Y0)) {

    # ── 独立 ink 簇：各合成一张图 ──
    if ($unit.Kind -eq "InkGroup") {
      try {
        $blobs = [System.Collections.Generic.List[byte[]]]::new()
        foreach ($b in $unit.Group) {
          Add-InkBlobsFrom -Scope $b.Node -NsManager $NsManager -Target $blobs
        }
        # 纯 ink 不与图片对齐，故不做坐标映射 —— 保持 ISF 原生分辨率，细节更足
        $png = New-CompositeBlockPng -Images ([System.Collections.Generic.List[object]]::new()) -InkBlobs $blobs
        if ($png) {
          $lines.Add("![ink](data:image/png;base64,$([Convert]::ToBase64String($png)))")
          $lines.Add("")
        }
      } catch { $inkFailures++ }
      continue
    }

    $elem = $unit.Block.Node

    # ── 页面级独立图片（贴在画板上，不在任何 Outline 内） ──
    if ($unit.Block.Tag -eq "Image") {
      $imgIndex++
      $info = Get-ImageNodeInfo -Node $elem -NsManager $NsManager
      if ($info) {
        $uri = Convert-ImageBytesToDataUri -Bytes $info.Bytes -Format $info.Format `
          -Context "$BaseFileName #$imgIndex"
        if ($uri) { $lines.Add("![image]($uri)"); $lines.Add("") }
      }
      continue
    }

    if ($unit.Block.Tag -ne "Outline") { continue }

    # ── Outline：按 OE 顺序收集「文本 / 图片」条目，并记录图片的页面坐标 ──
    #
    # Outline 内的图片没有 Position —— 它们参与流式布局，纵向位置由前面条目的
    # 高度累加决定。图片高度精确（Image/Size），文本行高未知，但可以用
    # 「Outline 总高 - 图片高合计」除以文本行数反推：实测该值与空 Outline
    # 的高度一致（26.86 = 2 × 13.4277），说明这是 OneNote 自己的算法。
    $items = [System.Collections.Generic.List[PSCustomObject]]::new()

    $tables = $elem.SelectNodes(".//one:Table", $NsManager)
    if ($tables -and $tables.Count -gt 0) {
      foreach ($table in $tables) {
        $isFirst = $true
        foreach ($row in $table.SelectNodes("one:Row", $NsManager)) {
          $cells = foreach ($cell in $row.SelectNodes("one:Cell", $NsManager)) {
            $texts = foreach ($t in $cell.SelectNodes(".//one:T", $NsManager)) { Clean-HtmlText $t.InnerText }
            (($texts | Where-Object { $_ }) -join " ") -replace "\|", "&#124;"
          }
          $items.Add([PSCustomObject]@{ Type = "Text"; Text = ("| " + ($cells -join " | ") + " |") })
          if ($isFirst) {
            $seps = foreach ($c in $cells) { "---" }
            $items.Add([PSCustomObject]@{ Type = "Text"; Text = ("| " + ($seps -join " | ") + " |") })
            $isFirst = $false
          }
        }
        $items.Add([PSCustomObject]@{ Type = "Text"; Text = "" })
      }
    }

    foreach ($oe in $elem.SelectNodes(".//one:OE", $NsManager)) {
      if ($oe.SelectSingleNode("ancestor::one:Cell", $NsManager)) { continue }

      $depth = 0
      $curr = $oe.ParentNode
      while ($curr -and $curr.LocalName -ne "Outline") {
        if ($curr.LocalName -eq "OEChildren") { $depth++ }
        $curr = $curr.ParentNode
      }
      $indent = "  " * [Math]::Max(0, $depth - 1)

      $listNode = $oe.SelectSingleNode("one:List", $NsManager)
      $prefix = ""
      if ($listNode) {
        $prefix = if ($listNode.SelectSingleNode("one:Number", $NsManager)) { "1. " } else { "- " }
      }

      $tNode = $oe.SelectSingleNode("one:T", $NsManager)
      $hasText = $false
      if ($tNode) {
        $text = Clean-HtmlText $tNode.InnerText
        if ($text -and $text -ne $title) {
          $hasText = $true
          $body = if ($prefix) { "$indent$prefix$text" } else { "$indent$text" }
          # 临近的文本行合并成一个段落，避免逐行成段
          $last = if ($items.Count -gt 0) { $items[$items.Count - 1] } else { $null }
          if ($last -and $last.Type -eq "Text" -and -not $last.IsImageRow) {
            $items[$items.Count - 1] = [PSCustomObject]@{ Type = "Text"; Text = ($last.Text + "`n" + $body); IsImageRow = $false }
          } else {
            $items.Add([PSCustomObject]@{ Type = "Text"; Text = $body; IsImageRow = $false })
          }
        }
      }
      if (-not $hasText) {
        # 空文本行仍占纵向空间，显式记为一行（供行高反推计数）
        $last = if ($items.Count -gt 0) { $items[$items.Count - 1] } else { $null }
        if (-not $last -or $last.Type -ne "Text" -or $last.IsImageRow) {
          $items.Add([PSCustomObject]@{ Type = "Text"; Text = ""; IsImageRow = $false })
        }
      }

      $imgNode = $oe.SelectSingleNode("one:Image", $NsManager)
      if ($imgNode) {
        $info = Get-ImageNodeInfo -Node $imgNode -NsManager $NsManager
        $items.Add([PSCustomObject]@{ Type = "Image"; Info = $info; IsImageRow = $true })
      }

      $inkText = Get-InkRecognizedText -Scope $oe -NsManager $NsManager
      if ($inkText) {
        $items.Add([PSCustomObject]@{ Type = "Text"; Text = "$indent$inkText"; IsImageRow = $false })
      }
    }

    $imageItems = @($items | Where-Object { $_.Type -eq "Image" -and $_.Info })
    $textRowCount = @($items | Where-Object { $_.Type -eq "Text" }).Count
    $imageHeightSum = 0.0
    foreach ($it in $imageItems) { $imageHeightSum += $it.Info.Height }

    $rowHeight = 0.0
    if ($textRowCount -gt 0) {
      $residual = [Math]::Abs($unit.Block.Y1 - $unit.Block.Y0) - $imageHeightSum
      if ($residual -gt 0) { $rowHeight = $residual / $textRowCount }
    }

    # 给每张图算出页面坐标（容器原点 + 纵向累加）
    $offset = 0.0
    foreach ($it in $items) {
      if ($it.Type -eq "Image") {
        if ($it.Info) {
          $it | Add-Member -NotePropertyName X0 -NotePropertyValue $unit.Block.X0 -Force
          $it | Add-Member -NotePropertyName Y0 -NotePropertyValue ($unit.Block.Y0 + $offset) -Force
          $it | Add-Member -NotePropertyName X1 -NotePropertyValue ($unit.Block.X0 + $it.Info.Width) -Force
          $it | Add-Member -NotePropertyName Y1 -NotePropertyValue ($unit.Block.Y0 + $offset + $it.Info.Height) -Force
        }
        $offset += if ($it.Info) { $it.Info.Height } else { 0.0 }
      } else {
        $offset += $rowHeight
      }
    }

    $assigned = if ($outlineGroups.ContainsKey($elem.objectID)) { $outlineGroups[$elem.objectID] } else { @() }
    $inkBlobs = [System.Collections.Generic.List[byte[]]]::new()
    # 与 blob 并行记录每处 ink 的页面坐标矩形，用于把 ISF 坐标映射到页面坐标。
    $inkRects = [System.Collections.Generic.List[object]]::new()
    foreach ($grp in $assigned) {
      foreach ($b in $grp) {
        $inkRects.Add([PSCustomObject]@{ X0 = $b.X0; Y0 = $b.Y0; X1 = $b.X1; Y1 = $b.Y1 })
        Add-InkBlobsFrom -Scope $b.Node -NsManager $NsManager -Target $inkBlobs
      }
    }
    # Outline 内部自带的 ink（若有）也算进同一张合成图
    foreach ($n in $elem.SelectNodes(".//one:InkDrawing", $NsManager)) {
      $npos = $n.SelectSingleNode("one:Position", $NsManager)
      $nsz = $n.SelectSingleNode("one:Size", $NsManager)
      if ($npos -and $nsz) {
        $nx = [double]$npos.x; $ny = [double]$npos.y
        $inkRects.Add([PSCustomObject]@{
            X0 = $nx; Y0 = $ny
            X1 = ($nx + [double]$nsz.width); Y1 = ($ny + [double]$nsz.height)
          })
      }
      Add-InkBlobsFrom -Scope $n -NsManager $NsManager -Target $inkBlobs
    }
    $inkTargetRect = if ($inkRects.Count -gt 0) { Get-BoundsUnion -Blocks $inkRects } else { $null }

    if ($inkBlobs.Count -gt 0 -and $imageItems.Count -gt 0) {
      # ── 有 ink 也有图：文本先出，图与标注合成一张 ──
      $textOnly = @($items | Where-Object { $_.Type -eq "Text" } | ForEach-Object { $_.Text })
      $t = ($textOnly -join "`n").Trim()
      if ($t) { $lines.Add($t); $lines.Add("") }

      $compositeOk = $false
      try {
        $imgs = [System.Collections.Generic.List[object]]::new()
        foreach ($it in $imageItems) {
          $bmp = New-WpfBitmapFromBytes -Bytes $it.Info.Bytes -Format $it.Info.Format
          if (-not $bmp) { continue }
          $imgs.Add([PSCustomObject]@{ Bitmap = $bmp; X0 = $it.X0; Y0 = $it.Y0; X1 = $it.X1; Y1 = $it.Y1 })
        }
        $png = New-CompositeBlockPng -Images $imgs -InkBlobs $inkBlobs -InkTargetRect $inkTargetRect
        if ($png) {
          $lines.Add("![image](data:image/png;base64,$([Convert]::ToBase64String($png)))")
          $lines.Add("")
          $compositeOk = $true
        }
      } catch {
        Write-Warn2 ("  [$BaseFileName] 图文合成失败：{0} @行{1} —— {2}" -f `
          $_.Exception.GetType().Name, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
      }

      if (-not $compositeOk) {
        # 降级：合成是"增强"，不该让它把图也一起丢掉。
        # 图片恢复为独立输出，ink 退化成单独一张，内容都不丢。
        $inkFailures++
        foreach ($it in $imageItems) {
          $imgIndex++
          $uri = Convert-ImageBytesToDataUri -Bytes $it.Info.Bytes -Format $it.Info.Format `
            -Context "$BaseFileName #$imgIndex"
          if ($uri) { $lines.Add("![image]($uri)"); $lines.Add("") }
        }
        try {
          $png2 = New-CompositeBlockPng -Images ([System.Collections.Generic.List[object]]::new()) -InkBlobs $inkBlobs -InkTargetRect $inkTargetRect
          if ($png2) {
            $lines.Add("![ink](data:image/png;base64,$([Convert]::ToBase64String($png2)))")
            $lines.Add("")
          }
        } catch {
          Write-Warn2 ("  [$BaseFileName] ink 单独渲染也失败：{0} @行{1} —— {2}" -f `
            $_.Exception.GetType().Name, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
        }
      }

    } elseif ($inkBlobs.Count -gt 0 -and $imageItems.Count -eq 0) {
      # ── 只有 ink（含 Outline 内部 ink）：合成一张 ──
      $textOnly = @($items | Where-Object { $_.Type -eq "Text" } | ForEach-Object { $_.Text })
      $t = ($textOnly -join "`n").Trim()
      if ($t) { $lines.Add($t); $lines.Add("") }
      try {
        $png = New-CompositeBlockPng -Images ([System.Collections.Generic.List[object]]::new()) -InkBlobs $inkBlobs -InkTargetRect $inkTargetRect
        if ($png) {
          $lines.Add("![ink](data:image/png;base64,$([Convert]::ToBase64String($png)))")
          $lines.Add("")
        } else {
          $inkFailures++
        }
      } catch {
        $inkFailures++
        Write-Warn2 ("  [$BaseFileName] ink 渲染失败：{0} @行{1} —— {2}" -f `
          $_.Exception.GetType().Name, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
      }

    } else {
      # ── 无 ink：按 OE 顺序输出文本与图片的交错 ──
      $pending = [System.Collections.Generic.List[string]]::new()
      foreach ($it in $items) {
        if ($it.Type -eq "Text") {
          if ($it.Text) { $pending.Add($it.Text) }
          continue
        }
        if ($pending.Count -gt 0) {
          $lines.Add((($pending -join "`n").Trim()))
          $lines.Add("")
          $pending.Clear()
        }
        $imgIndex++
        if (-not $it.Info) { continue }
        $uri = Convert-ImageBytesToDataUri -Bytes $it.Info.Bytes -Format $it.Info.Format `
          -Context "$BaseFileName #$imgIndex"
        if ($uri) { $lines.Add("![image]($uri)"); $lines.Add("") }
      }
      if ($pending.Count -gt 0) {
        $lines.Add((($pending -join "`n").Trim()))
        $lines.Add("")
      }
    }
  }

  # 兜底：页面确有 ink，但最终一张 ink 图都没产出 —— 这通常意味着渲染失败
  # 或 blob 提取断了。静默输出空文档是最难排查的失败模式（"会话.md" 曾经
  # 只有标题），所以显式留下标记。
  if (-not $SkipInkImages) {
    $inkNodeCount = $PageDoc.SelectNodes("//one:InkDrawing | //one:InkWord", $NsManager).Count
    $inkUriCount = @($lines | Where-Object { $_ -match '^!\[ink\]\(' }).Count
    $compositeCount = @($lines | Where-Object { $_ -match '^!\[image\]\(data:image/png' }).Count
    if ($inkNodeCount -gt 0 -and $inkUriCount -eq 0 -and $compositeCount -eq 0) {
      $lines.Add("<!-- 本页有 $inkNodeCount 处手写内容，但未能渲染（导出器已跳过）。 -->")
      $lines.Add("")
      $inkFailures++
    }
  }

  return (($lines -join "`n").Trim()) + "`n"
}

# ── 参数校验 ────────────────────────────────────────────────────────────────

# GUI 以子进程方式调用本脚本（powershell.exe -File ... -ProgressFile <path>），
# 由它读取进度；命令行调用时不传该参数，行为完全不变。
function Write-ProgressEvent {
  param([string]$Phase, [int]$Done, [int]$Total, [string]$Message = "", $Data)
  if (-not $ProgressFile) { return }
  try {
    $evt = [ordered]@{
      phase   = $Phase
      done    = $Done
      total   = $Total
      message = $Message
      time    = (Get-Date).ToString("HH:mm:ss")
    }
    if ($null -ne $Data) { $evt["data"] = @($Data) }
    $json = $evt | ConvertTo-Json -Compress -Depth 6
    # 追加写入：GUI 侧从头读到尾即可，无需知道文件是否被重写
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json + "`n")
    $fs = [System.IO.File]::Open($ProgressFile, [System.IO.FileMode]::Append,
                                 [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    try { $fs.Write($bytes, 0, $bytes.Length) } finally { $fs.Dispose() }
  } catch {
    # 进度上报失败绝不能影响导出本身
  }
}

# 顶层兜底：终止性错误（连不上 OneNote、层级读取失败等）也转成一条 progress
# 事件，否则 GUI 只能看到退出码非 0、拿不到原因。CLI 侧同时把完整错误写到
# stderr，与原先未捕获 throw 的表现一致。
trap {
  Write-ProgressEvent -Phase "error" -Done 0 -Total 0 -Message $_.Exception.Message
  ($_ | Out-String).Trim() | ForEach-Object { [Console]::Error.WriteLine($_) }
  exit 1
}

if (-not $List -and -not $ListNotebooks) {
  if (-not $OutputPath) {
    throw "缺少 -OutputPath 参数（导出目录）。若只想查看有哪些笔记本，加 -ListNotebooks；查看页面清单加 -List。"
  }
  if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
  }
  $OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

  if ($DumpXml) {
    $xmlDumpDir = Join-Path $OutputPath "_xml"
    New-Item -ItemType Directory -Path $xmlDumpDir -Force | Out-Null
    Write-Step "诊断模式：原始页面 XML 将写入 $xmlDumpDir"
  }
}

# ── 连接 OneNote ────────────────────────────────────────────────────────────

try {
  $oneNote = New-Object -ComObject OneNote.Application
} catch {
  throw "无法连接 OneNote。请确认已安装 OneNote 桌面版（Microsoft Store 版本没有自动化接口），并先手动打开一次。原始错误：$($_.Exception.Message)"
}

$hierarchyXml = ""
try {
  $oneNote.GetHierarchy("", 4, [ref]$hierarchyXml)
} catch {
  throw "读取 OneNote 层级失败：$($_.Exception.Message)"
}

if ([string]::IsNullOrWhiteSpace($hierarchyXml)) {
  throw "OneNote 返回了空的层级。请确认本机已同步至少一个笔记本。"
}

[xml]$hierarchyDoc = $hierarchyXml
$ns = New-Object System.Xml.XmlNamespaceManager($hierarchyDoc.NameTable)
$ns.AddNamespace("one", $hierarchyDoc.DocumentElement.NamespaceURI)

# ── 计算每个页面的相对路径（含子页面嵌套） ──────────────────────────────────

# OneNote 的子页面不是一个"关系"，而是一个"位置"：pageLevel 记录缩进层级
# （1..3），某个页面是它前面最近的、层级更低的那个页面的子页面。
# 因此要用一组游标，按文档顺序扫描页面时维护各层当前父页面名。
function Get-SubPageFolderPrefix {
  param([string]$SectionKey, [int]$Depth)
  if (-not $script:sectionCursors.ContainsKey($SectionKey)) {
    $script:sectionCursors[$SectionKey] = @("", "", "")
  }
  $cursors = $script:sectionCursors[$SectionKey]
  $parts = @()
  for ($i = 0; $i -lt $Depth; $i++) {
    if ($i -lt $cursors.Count -and $cursors[$i]) { $parts += $cursors[$i] }
  }
  return $parts
}

function Set-SubPageCursor {
  param([string]$SectionKey, [int]$Depth, [string]$Name)
  if (-not $script:sectionCursors.ContainsKey($SectionKey)) {
    $script:sectionCursors[$SectionKey] = @("", "", "")
  }
  $cursors = $script:sectionCursors[$SectionKey]
  if ($Depth -lt $cursors.Count) {
    $cursors[$Depth] = $Name
    # 更深的层级失效：新页面的子页面关系由它自己建立
    for ($i = $Depth + 1; $i -lt $cursors.Count; $i++) { $cursors[$i] = "" }
  }
}

$script:sectionCursors = @{}

function Add-AncestorPathSegments {
  param(
    $PageNode,
    [System.Collections.Generic.List[string]]$Target
  )
  $curr = $PageNode.ParentNode
  while ($curr -and $curr.LocalName -ne "Notebooks" -and $curr.LocalName -ne "#document") {
    if ($curr.name) {
      $safe = Get-SafeSegment $curr.name
      if ($safe) { $Target.Add($safe) }
    }
    $curr = $curr.ParentNode
  }
  $Target.Reverse()
}

# ── 列出笔记本 ──────────────────────────────────────────────────────────────

if ($ListNotebooks) {
  # 只统计非回收站的页面，否则计数会与 -List 对不上
  $nbNodes = @($hierarchyDoc.SelectNodes("//one:Notebook", $ns))
  $rows = foreach ($nb in $nbNodes) {
    [PSCustomObject]@{
      笔记本 = $nb.name
      分区数 = $nb.SelectNodes(".//one:Section", $ns).Count
      页面数 = @($nb.SelectNodes(".//one:Page", $ns) |
                 Where-Object { $_.GetAttribute("isInRecycleBin") -ne "true" }).Count
    }
  }

  if (-not $rows) {
    Write-Step "未找到任何笔记本。请确认 OneNote 已登录并同步至少一个笔记本。"
    Write-ProgressEvent -Phase "error" -Done 0 -Total 0 `
      -Message "未找到任何笔记本。请确认 OneNote 已登录并同步至少一个笔记本。"
    exit 1
  }

  $rows | Format-Table -AutoSize
  Write-Step ""
  Write-Step ("共 {0} 个笔记本。" -f @($rows).Count)

  # 结构化数据供 GUI 填充下拉框：笔记本 → 其下所有分区名（去重）
  $hierarchy = foreach ($nb in $nbNodes) {
    [PSCustomObject]@{
      name     = $nb.name
      sections = @($nb.SelectNodes(".//one:Section", $ns) |
                   ForEach-Object { $_.name } | Sort-Object -Unique)
    }
  }
  Write-ProgressEvent -Phase "hierarchy" -Done @($rows).Count -Total @($rows).Count `
    -Message ("共 {0} 个笔记本" -f @($rows).Count) -Data @($hierarchy)
  exit 0
}

# ── 收集页面 ────────────────────────────────────────────────────────────────

$allPageNodes = $hierarchyDoc.SelectNodes("//one:Page", $ns)
$filteredPages = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()

foreach ($page in $allPageNodes) {
  if ($Notebook) {
    $nb = $page.SelectSingleNode("ancestor::one:Notebook", $ns)
    if (-not $nb -or $nb.name -ne $Notebook) { continue }
  }
  if ($Section) {
    $sec = $page.SelectSingleNode("ancestor::one:Section", $ns)
    if (-not $sec -or $sec.name -ne $Section) { continue }
  }
  $filteredPages.Add($page)
}

if ($filteredPages.Count -eq 0) {
  Write-Step "没有匹配的页面（笔记本：$(if ($Notebook) { $Notebook } else { '全部' })；分区：$(if ($Section) { $Section } else { '全部' })）。"
  exit 0
}

# 为每个页面算出 (相对目录, 文件名)，并处理同秒级别的重名冲突。
$usedPaths = @{}
$planned = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($page in $filteredPages) {
  $secNode = $page.SelectSingleNode("ancestor::one:Section", $ns)
  $sectionKey = if ($secNode) { $secNode.ID } else { "" }

  $depth = 0
  if ($page.pageLevel) {
    $parsed = 0
    if ([int]::TryParse($page.pageLevel, [ref]$parsed) -and $parsed -gt 1) {
      $depth = [Math]::Min($parsed - 1, 2)
    }
  }

  $folderSegments = [System.Collections.Generic.List[string]]::new()
  Add-AncestorPathSegments -PageNode $page -Target $folderSegments
  foreach ($prefixPart in (Get-SubPageFolderPrefix -SectionKey $sectionKey -Depth $depth)) {
    $folderSegments += $prefixPart
  }

  $safeTitle = Get-SafeFileName $page.name
  Set-SubPageCursor -SectionKey $sectionKey -Depth $depth -Name $safeTitle

  $relFolder = ($folderSegments -join "\")
  $relFile = if ($relFolder) { Join-Path $relFolder "$safeTitle.md" } else { "$safeTitle.md" }

  # 重名去重：周会.md / 周会 (2).md / 周会 (3).md …
  $key = $relFile.ToLowerInvariant()
  if ($usedPaths.ContainsKey($key)) {
    $n = $usedPaths[$key] + 1
    while ($true) {
      $candidate = if ($relFolder) { Join-Path $relFolder "$safeTitle ($n).md" } else { "$safeTitle ($n).md" }
      if (-not $usedPaths.ContainsKey($candidate.ToLowerInvariant())) {
        $relFile = $candidate
        $usedPaths[$candidate.ToLowerInvariant()] = 1
        break
      }
      $n++
    }
  } else {
    $usedPaths[$key] = 1
  }

  $planned.Add([PSCustomObject]@{
      Page     = $page
      RelFile  = $relFile
      Title    = $page.name
    })
}

# ── 仅列出 ──────────────────────────────────────────────────────────────────

if ($List) {
  $rows = foreach ($p in $planned) {
    [PSCustomObject]@{
      笔记本 = ($p.Page.SelectSingleNode("ancestor::one:Notebook", $ns)).name
      相对路径 = $p.RelFile
      修改时间 = $p.Page.lastModifiedTime
    }
  }
  $rows | Format-Table -AutoSize
  Write-Step ""
  Write-Step ("共 {0} 个页面。" -f $planned.Count)
  exit 0
}

# ── 增量判断 ────────────────────────────────────────────────────────────────

$manifest = @{}
if (Test-Path -LiteralPath $manifestPath) {
  try {
    # 必须显式按 UTF-8 读取：manifest 是以无 BOM 的 UTF-8 写出的，
    # 而 PowerShell 5.1 的 Get-Content 默认按系统 ANSI 代码页解码，
    # 会把中文路径读成乱码，其中某些字节对还会被 JSON 当成非法转义序列
    # （表现为 "无法识别的转义序列"），导致每次都误判为记录损坏。
    $rawText = [System.IO.File]::ReadAllText($manifestPath, [System.Text.UTF8Encoding]::new($false))
    $rawJson = $rawText | ConvertFrom-Json
    if ($rawJson -and $rawJson.pages) {
      foreach ($prop in $rawJson.pages.PSObject.Properties) {
        $manifest[$prop.Name] = [ordered]@{
          lastModifiedTime = $prop.Value.lastModifiedTime
          path             = $prop.Value.path
        }
      }
    }
  } catch {
    Write-Warn2 "增量记录损坏，本次改为全量导出：$($_.Exception.Message)"
    $manifest = @{}
  }
}

$toExport = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($p in $planned) {
  $absolute = Join-Path $OutputPath $p.RelFile
  $needs = $false
  if ($Force) {
    $needs = $true
  } elseif (-not (Test-Path -LiteralPath $absolute)) {
    $needs = $true
  } elseif ($manifest.ContainsKey($p.Page.ID)) {
    $recorded = [datetime]$manifest[$p.Page.ID].lastModifiedTime
    if ([datetime]$p.Page.lastModifiedTime -gt $recorded.AddSeconds(2)) { $needs = $true }
  } else {
    $fileInfo = Get-Item -LiteralPath $absolute
    if ([datetime]$p.Page.lastModifiedTime -gt $fileInfo.LastWriteTime.ToUniversalTime().AddSeconds(5)) {
      $needs = $true
    }
  }
  if ($needs) { $toExport.Add($p) }
}

if ($toExport.Count -eq 0) {
  Write-Step ("OneNote 内容无变化，跳过导出（共 {0} 个页面）。" -f $planned.Count)
  Write-ProgressEvent -Phase "up-to-date" -Done 0 -Total 0 `
    -Message ("内容无变化，无需导出（共 {0} 个页面）。" -f $planned.Count)
  exit 0
}

Write-Step ("开始导出：{0} / {1} 个页面有更新。" -f $toExport.Count, $planned.Count)
Write-ProgressEvent -Phase "start" -Done 0 -Total $toExport.Count `
  -Message ("共 {0} 个页面待导出。" -f $toExport.Count)
Write-Step ""

# ── 导出 ────────────────────────────────────────────────────────────────────

$successCount = 0
$failedPages = [System.Collections.Generic.List[string]]::new()
# 手绘渲染失败是"已知降级"（ISF 解不出来 / 渲染器不可用），单张跳过即可，
# 但必须在末尾汇总告知，否则用户会以为手绘图都导出来了。
$totalInkFailures = 0

foreach ($p in $toExport) {
  try {
    $inkFailures = 0
    $pxml = ""
    $oneNote.GetPageContent($p.Page.ID, [ref]$pxml, 7)

    # 诊断：在解析之前落盘原始 XML，保留 OneNote 返回的原始形态
    # （含所有 Position/Size 坐标与 ink 的 Data），用于排查坐标问题。
    if ($DumpXml) {
      $xmlPath = Join-Path $xmlDumpDir (($p.RelFile -replace '[\\/]', '__') + '.xml')
      [System.IO.File]::WriteAllText($xmlPath, $pxml, [System.Text.UTF8Encoding]::new($false))
    }

    [xml]$pdoc = $pxml
    $pns = New-Object System.Xml.XmlNamespaceManager($pdoc.NameTable)
    $pns.AddNamespace("one", $pdoc.DocumentElement.NamespaceURI)

    $absolute = Join-Path $OutputPath $p.RelFile
    $targetDir = Split-Path $absolute -Parent
    if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
      New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $md = Convert-PageXmlToMarkdown -PageDoc $pdoc -NsManager $pns `
      -PageTitle $p.Title -BaseFileName (Get-SafeFileName $p.Title) `
      -SkipInkImages:$SkipInkImages `
      -InkToleranceRatio $InkToleranceRatio -InkGapThreshold $InkGapThreshold

    [System.IO.File]::WriteAllText($absolute, $md, [System.Text.UTF8Encoding]::new($false))

    $manifest[$p.Page.ID] = [ordered]@{
      lastModifiedTime = $p.Page.lastModifiedTime
      path             = $p.RelFile
    }
    $successCount++
    $totalInkFailures += $inkFailures
    $suffix = if ($inkFailures -gt 0) { "（$inkFailures 张手绘图渲染失败，已跳过）" } else { "" }
    Write-Step ("  [成功] {0}{1}" -f $p.RelFile, $suffix)
    Write-ProgressEvent -Phase "page" -Done ($successCount + $failedPages.Count) `
      -Total $toExport.Count -Message ("{0}{1}" -f $p.RelFile, $suffix)
  } catch {
    $failedPages.Add($p.RelFile)
    Write-Warn2 ("  [失败] {0} —— {1}" -f $p.RelFile, $_.Exception.Message)
    Write-ProgressEvent -Phase "page" -Done ($successCount + $failedPages.Count) `
      -Total $toExport.Count -Message ("{0} —— {1}" -f $p.RelFile, $_.Exception.Message)
  }
}

$manifestData = [ordered]@{
  version  = 1
  lastSync = (Get-Date).ToUniversalTime().ToString("o")
  pages    = $manifest
}
$manifestJson = $manifestData | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.UTF8Encoding]::new($false))

Write-Step ""
Write-Step ("完成：成功 {0} 个，失败 {1} 个。" -f $successCount, $failedPages.Count)
Write-ProgressEvent -Phase "done" -Done $toExport.Count -Total $toExport.Count `
  -Message ("成功 {0} 个，失败 {1} 个。" -f $successCount, $failedPages.Count)
if ($totalInkFailures -gt 0) {
  Write-Step ("注意：另有 {0} 张手绘图渲染失败被跳过（页面本身已导出，只是缺这几张图）。" -f $totalInkFailures)
}
if ($failedPages.Count -gt 0) {
  Write-Step "失败清单："
  foreach ($f in $failedPages) { Write-Step ("  - {0}" -f $f) }
}
Write-Step ""
Write-Step ("导出目录：{0}" -f $OutputPath)
exit 0

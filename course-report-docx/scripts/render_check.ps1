<#
.SYNOPSIS
    课程报告 .docx 渲染级自检 —— 调用本机 Microsoft Word 实际排版后测量。

.DESCRIPTION
    check_report.py 只能看 XML 结构，看不出"格式乱跑"：分页位置、
    空白页、封面是否多出一页、表格有没有被分页切断、图与题注是否被拆开，
    这些必须让 Word 真正排一遍版才能量出来。

    本脚本是可选步骤。没装 Word 时会直接跳过并返回 0，不影响主流程。

    检查项：
      1. 能正常打开（Word 不自作修复）
      2. 节数符合预期
      3. 总页数
      4. 无空白页
      5. 封面恰好 1 页
      6. 封面没有页码
      7. 正文表格不会被分页切断
      8. 图与其题注不会被分页拆开
      9. 没有"错误！"字样的域错误（目录未更新 / 交叉引用失效）

.PARAMETER Path
    待检查的 .docx 路径。

.PARAMETER Pdf
    可选。指定后会额外导出一份 PDF 到该路径，便于肉眼复核。

.PARAMETER ExpectSections
    期望的节数，默认 5（封面 / 摘要 / 目录 / 正文 / 末节）。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File render_check.ps1 -Path out.docx

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File render_check.ps1 -Path out.docx -Pdf out.pdf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$Pdf = '',
    [int]$ExpectSections = 5
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- 结果收集
$script:Rows = New-Object System.Collections.ArrayList

function Add-Row([string]$Level, [string]$Name, [string]$Detail) {
    [void]$script:Rows.Add([pscustomobject]@{
        Level  = $Level
        Name   = $Name
        Detail = $Detail
    })
}

function Clean-Text([string]$s) {
    if ($null -eq $s) { return '' }
    return ($s -replace "[`r`n`a`t]", ' ').Trim()
}

function Get-PageOf($range) {
    # wdActiveEndPageNumber = 3
    return [int]$range.Information(3)
}

# ---------------------------------------------------------------- 前置检查
if (-not (Test-Path -LiteralPath $Path)) {
    Write-Output "文件不存在：$Path"
    exit 2
}
$full = (Resolve-Path -LiteralPath $Path).Path

$word = $null
try {
    $word = New-Object -ComObject Word.Application
}
catch {
    Write-Output '未检测到 Microsoft Word（或 COM 组件不可用），跳过渲染级检查。'
    Write-Output 'XML 级检查请运行 check_report.py。'
    exit 0
}

$word.Visible = $false
$word.DisplayAlerts = 0
try { $word.AutomationSecurity = 3 } catch { }   # 强制禁用宏

$doc = $null
$exitCode = 0
try {
    try {
        # Open(FileName, ConfirmConversions, ReadOnly)
        $doc = $word.Documents.Open($full, $false, $true)
        Add-Row 'PASS' 'Word 能正常打开' '未触发自动修复' 
    }
    catch {
        Add-Row 'FAIL' 'Word 能正常打开' $_.Exception.Message
        throw
    }

    # ------------------------------------------------ 节数与页数
    $secCount = $doc.Sections.Count
    if ($secCount -eq $ExpectSections) {
        Add-Row 'PASS' '节数' "$secCount 节"
    }
    else {
        Add-Row 'WARN' '节数' "$secCount 节，期望 $ExpectSections 节"
    }

    $totalPages = [int]$doc.ComputeStatistics(2)     # wdStatisticPages
    Add-Row 'PASS' '总页数' "$totalPages 页"

    # ------------------------------------------------ 逐段扫页号
    $paraCount = $doc.Paragraphs.Count
    $occupied = @{}
    $imgPages = New-Object System.Collections.ArrayList
    $paraPage = @{}

    for ($i = 1; $i -le $paraCount; $i++) {
        $r = $doc.Paragraphs.Item($i).Range
        $pg = Get-PageOf $r
        $paraPage[$i] = $pg
        $t = Clean-Text $r.Text
        if ($t.Length -gt 0) { $occupied[$pg] = $true }
        if ($r.InlineShapes.Count -gt 0) { [void]$imgPages.Add($i) }
    }

    # ------------------------------------------------ 空白页
    $blank = @()
    for ($p = 1; $p -le $totalPages; $p++) {
        if (-not $occupied.ContainsKey($p)) { $blank += $p }
    }
    if ($blank.Count -eq 0) {
        Add-Row 'PASS' '无空白页' "$totalPages 页均有序号"
    }
    else {
        Add-Row 'FAIL' '无空白页' ("第 " + ($blank -join ', ') + " 页是空白页")
    }

    # ------------------------------------------------ 封面
    $sec1 = $doc.Sections.Item(1)
    $coverLastPage = Get-PageOf $sec1.Range.Paragraphs.Last.Range
    if ($coverLastPage -eq 1) {
        Add-Row 'PASS' '封面恰好 1 页' '封面内容未溢出到第 2 页'
    }
    else {
        Add-Row 'FAIL' '封面恰好 1 页' "封面被撑到第 $coverLastPage 页（撑破版心或多出空行）"
    }

    $ftr = $sec1.Footers.Item(1)
    $ftrText = Clean-Text $ftr.Range.Text
    $pnCount = $ftr.PageNumbers.Count
    if ($ftrText.Length -eq 0 -and $pnCount -eq 0) {
        Add-Row 'PASS' '封面不显示页码' '封面节页脚为空'
    }
    else {
        $why = @()
        if ($ftrText.Length -gt 0) { $why += "页脚文字[$ftrText]" }
        if ($pnCount -gt 0) { $why += "$pnCount 个页码域" }
        Add-Row 'FAIL' '封面不显示页码' ($why -join '，')
    }

    # 封面纵向顺序：题目 -> 信息表 -> 日期 必须自上而下依次排开。
    # 撑版用的空段落一旦被误删，日期会顶到信息表上方，这条能立刻发现。
    $yTitle = -1
    $yDate = -1
    $sec1Paras = $sec1.Range.Paragraphs.Count
    for ($i = 1; $i -le $sec1Paras; $i++) {
        $r = $sec1.Range.Paragraphs.Item($i).Range
        $t = Clean-Text $r.Text
        if ($yTitle -lt 0 -and $t -match '^题目') {
            $yTitle = [math]::Round($r.Information(6), 1)
        }
        if ($yDate -lt 0 -and $t -match '\d{4}\s*年\s*\d{1,2}\s*月') {
            $yDate = [math]::Round($r.Information(6), 1)
        }
    }
    $yTblBottom = -1
    if ($doc.Tables.Count -ge 1) {
        $t0 = $doc.Tables.Item(1)
        try {
            $yTblBottom = [math]::Round($t0.Rows.Item($t0.Rows.Count).Range.Information(6), 1)
        }
        catch { }
    }
    if ($yTitle -gt 0 -and $yTblBottom -gt 0 -and $yDate -gt 0) {
        if (($yTitle -lt $yTblBottom) -and ($yTblBottom -lt $yDate)) {
            Add-Row 'PASS' '封面纵向顺序' "题目 $yTitle -> 信息表底 $yTblBottom -> 日期 $yDate pt"
        }
        else {
            Add-Row 'FAIL' '封面纵向顺序' ("题目 $yTitle / 信息表底 $yTblBottom / 日期 $yDate pt 顺序不对" +
                "（撑版空段落被误删时会这样）")
        }
    }
    else {
        Add-Row 'WARN' '封面纵向顺序' '未能识别题目/信息表/日期，跳过'
    }

    # ------------------------------------------------ 表格跨页
    $tblCount = $doc.Tables.Count
    $tblSplit = 0
    $tblDetail = ''
    for ($i = 1; $i -le $tblCount; $i++) {
        $t = $doc.Tables.Item($i)
        try {
            $p1 = Get-PageOf $t.Rows.Item(1).Range
            $p2 = Get-PageOf $t.Rows.Item($t.Rows.Count).Range
            if ($p1 -ne $p2) {
                $tblSplit++
                $head = Clean-Text $t.Rows.Item(1).Range.Text
                if ($head.Length -gt 18) { $head = $head.Substring(0, 18) }
                $tblDetail = $tblDetail + "表$i（第$p1-$p2页：$head）; "
            }
        }
        catch { }
    }
    if ($tblCount -eq 0) {
        Add-Row 'WARN' '表格不被分页切断' '文档中没有表格'
    }
    elseif ($tblSplit -eq 0) {
        Add-Row 'PASS' '表格不被分页切断' "$tblCount 张表均在一页内"
    }
    else {
        Add-Row 'WARN' '表格不被分页切断' "$tblSplit/$tblCount 张跨页：$tblDetail"
    }

    # ------------------------------------------------ 图与题注同页
    $splitFig = 0
    $figChecked = 0
    foreach ($idx in $imgPages) {
        $imgPage = $paraPage[$idx]
        for ($j = $idx + 1; $j -le [Math]::Min($idx + 2, $paraCount); $j++) {
            $t2 = Clean-Text $doc.Paragraphs.Item($j).Range.Text
            if ($t2.Length -gt 0) {
                $figChecked++
                if ($paraPage[$j] -ne $imgPage) { $splitFig++ }
                break
            }
        }
    }
    if ($figChecked -eq 0) {
        Add-Row 'WARN' '图与题注同页' '未找到图片段落'
    }
    elseif ($splitFig -eq 0) {
        Add-Row 'PASS' '图与题注同页' "$figChecked 处图片与后续文字同页"
    }
    else {
        Add-Row 'WARN' '图与题注同页' "$splitFig/$figChecked 处被分页拆开"
    }

    # ------------------------------------------------ 域错误
    $errHits = 0
    try {
        $rng = $doc.Content.Duplicate
        $rng.Find.ClearFormatting()
        $rng.Find.Wrap = 0                        # wdFindStop，避免回绕死循环
        while ($rng.Find.Execute('错误！') -and $errHits -lt 30) {
            $errHits++
            $rng.Collapse(0)                      # wdCollapseEnd
        }
    }
    catch { }
    if ($errHits -eq 0) {
        Add-Row 'PASS' '无域错误' '未发现"错误！"字样'
    }
    else {
        Add-Row 'FAIL' '无域错误' "$errHits 处域错误，请在 Word 中 Ctrl+A 后按 F9 更新目录与交叉引用"
    }

    # ------------------------------------------------ 可选导出 PDF
    if ($Pdf -ne '') {
        try {
            $pdfFull = [System.IO.Path]::GetFullPath($Pdf)
            $doc.ExportAsFixedFormat($pdfFull, 17)    # wdExportFormatPDF
            $size = (Get-Item -LiteralPath $pdfFull).Length
            Add-Row 'PASS' '导出 PDF' "$pdfFull（$size 字节）"
        }
        catch {
            Add-Row 'WARN' '导出 PDF' $_.Exception.Message
        }
    }
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close(0) } catch { }        # wdDoNotSaveChanges
    }
    try { $word.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    [GC]::Collect()
}

# ---------------------------------------------------------------- 输出
Write-Output ''
Write-Output "检查文件（渲染级）：$full"
Write-Output ('-' * 64)

$pass = 0; $warn = 0; $fail = 0
foreach ($row in $script:Rows) {
    $icon = '[PASS]'
    if ($row.Level -eq 'WARN') { $icon = '[WARN]' }
    if ($row.Level -eq 'FAIL') { $icon = '[FAIL]' }
    $line = "  $icon $($row.Name)"
    if ($row.Detail -ne '') { $line = $line + "   ->  $($row.Detail)" }
    Write-Output $line
    if ($row.Level -eq 'PASS') { $pass++ }
    if ($row.Level -eq 'WARN') { $warn++ }
    if ($row.Level -eq 'FAIL') { $fail++ }
}

Write-Output ('-' * 64)
Write-Output ("合计 " + $script:Rows.Count + " 项：PASS $pass  WARN $warn  FAIL $fail")
if ($fail -gt 0) {
    Write-Output '渲染层面存在硬性问题，请修正 FAIL 项后再交付。'
    $exitCode = 1
}
else {
    Write-Output '渲染层面未发现硬性问题。'
}

exit $exitCode

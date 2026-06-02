param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [ValidateSet('Audit', 'Fix', 'AuditAndFix')]
    [string]$Mode = 'AuditAndFix',

    [ValidateSet('Auto', 'First', 'Second', 'Third')]
    [string]$HeadingSystem = 'Auto',

    [string]$OutputDocx,
    [string]$BackupDir,
    [string]$ReportJson,
    [string]$ReportMd,
    [string]$ReportPdf,
    [switch]$NoReportPdf
)

$ErrorActionPreference = 'Stop'

function Resolve-FullPath([string]$Path) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function CmToPt([double]$cm) { $cm * 28.3464567 }

function Near([double]$a, [double]$b, [double]$tol = 1.2) {
    [Math]::Abs($a - $b) -le $tol
}

function Clean-Text($range) {
    if ($null -eq $range) { return '' }
    (($range.Text -replace "[`r`a`f]", '') -replace '\s+', ' ').Trim()
}

function Style-Name($range) {
    try { return $range.Style.NameLocal } catch { return '' }
}

function Is-InRange($range, $ranges) {
    foreach ($r in $ranges) {
        if ($range.Start -ge $r.Start -and $range.End -le $r.End) { return $true }
    }
    return $false
}

function Add-Check($list, [string]$category, [string]$requirement, [string]$status, [string]$result, [string]$note = '') {
    $list.Add([pscustomobject]@{
        category = $category
        requirement = $requirement
        status = $status
        result = $result
        note = $note
    }) | Out-Null
}

function Main-TextFingerprint($doc) {
    $tocRanges = @()
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $tocRanges += $doc.TablesOfContents.Item($i).Range
    }
    $parts = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if (Is-InRange $p.Range $tocRanges) { continue }
        $parts.Add((Clean-Text $p.Range)) | Out-Null
    }
    return ($parts -join "`n")
}

function Configure-Style($doc, [string]$styleName, [string]$eastAsia, [string]$ascii, [double]$size, [int]$bold, [int]$align, [double]$before, [double]$after, [bool]$exact22, [double]$firstIndentPt) {
    try {
        $style = $doc.Styles.Item($styleName)
        $style.Font.NameFarEast = $eastAsia
        $style.Font.NameAscii = $ascii
        $style.Font.NameOther = $ascii
        $style.Font.Size = $size
        $style.Font.Bold = $bold
        $style.ParagraphFormat.Alignment = $align
        $style.ParagraphFormat.SpaceBefore = $before
        $style.ParagraphFormat.SpaceAfter = $after
        if ($exact22) {
            $style.ParagraphFormat.LineSpacingRule = 4
            $style.ParagraphFormat.LineSpacing = 22
        }
        $style.ParagraphFormat.FirstLineIndent = $firstIndentPt
    } catch {}
}

function Set-ParaFormat($p, [string]$eastAsia, [string]$ascii, [double]$size, [int]$bold, [int]$align, [double]$before, [double]$after, [bool]$exact22, [double]$firstIndentPt) {
    $p.Range.Font.NameFarEast = $eastAsia
    $p.Range.Font.NameAscii = $ascii
    $p.Range.Font.NameOther = $ascii
    $p.Range.Font.Size = $size
    $p.Range.Font.Bold = $bold
    $p.Format.Alignment = $align
    $p.Format.SpaceBefore = $before
    $p.Format.SpaceAfter = $after
    if ($exact22) {
        $p.Format.LineSpacingRule = 4
        $p.Format.LineSpacing = 22
    }
    $p.Format.FirstLineIndent = $firstIndentPt
}

function Apply-YnuFormatting($doc) {
    $script:YnuHeadingSystem = Resolve-YnuHeadingSystem $doc $HeadingSystem

    foreach ($section in @($doc.Sections)) {
        $section.PageSetup.PageWidth = 595.3
        $section.PageSetup.PageHeight = 841.9
        $section.PageSetup.TopMargin = CmToPt 2.5
        $section.PageSetup.BottomMargin = CmToPt 2.0
        $section.PageSetup.LeftMargin = CmToPt 3.0
        $section.PageSetup.RightMargin = CmToPt 2.0

        foreach ($hf in @($section.Headers)) {
            try {
                $hf.Range.Font.NameFarEast = '宋体'
                $hf.Range.Font.NameAscii = 'Times New Roman'
                $hf.Range.Font.Size = 10.5
                $hf.Range.ParagraphFormat.Alignment = 1
            } catch {}
        }
        foreach ($ff in @($section.Footers)) {
            try {
                $ff.Range.Font.NameFarEast = '宋体'
                $ff.Range.Font.NameAscii = 'Times New Roman'
                $ff.Range.Font.Size = 10.5
                $ff.Range.ParagraphFormat.Alignment = 1
            } catch {}
        }
    }

    Configure-Style $doc '正文' '宋体' 'Times New Roman' 12 0 3 0 0 $true 24
    Configure-Style $doc 'Normal' '宋体' 'Times New Roman' 12 0 3 0 0 $true 24
    Configure-Style $doc '标题 1' '黑体' 'Times New Roman' 16 -1 1 9.6 6 $true 0
    Configure-Style $doc '标题 2' '黑体' 'Times New Roman' 14 -1 0 6 0 $true 0
    Configure-Style $doc '标题 3' '黑体' 'Times New Roman' 12 -1 0 0 0 $true 24
    Configure-Style $doc 'TOC 1' '黑体' 'Times New Roman' 12 0 0 0 0 $true 0
    Configure-Style $doc 'TOC 2' '宋体' 'Times New Roman' 12 0 0 0 0 $true 24
    Configure-Style $doc 'TOC 3' '宋体' 'Times New Roman' 12 0 0 0 0 $true 48
    Configure-Style $doc '题注' '黑体' 'Times New Roman' 12 0 1 0 0 $true 0

    $tocRanges = @()
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $tocRanges += $doc.TablesOfContents.Item($i).Range
    }

    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $text = Clean-Text $p.Range
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        $style = Style-Name $p.Range
        $inToc = Is-InRange $p.Range $tocRanges

        if ($inToc) {
            $level = 0
            if ($style -match 'TOC\s*(\d+)') { $level = [int]$Matches[1] }
            if ($level -eq 1) { Set-ParaFormat $p '黑体' 'Times New Roman' 12 0 0 0 0 $true 0 }
            elseif ($level -eq 2) { Set-ParaFormat $p '宋体' 'Times New Roman' 12 0 0 0 0 $true 24 }
            elseif ($level -eq 3) { Set-ParaFormat $p '宋体' 'Times New Roman' 12 0 0 0 0 $true 48 }
            continue
        }

        $level = Get-HeadingLevelGuess $style $text
        if ($level -eq 1) {
            Set-ParaFormat $p '黑体' 'Times New Roman' 16 -1 1 9.6 6 $true 0
            continue
        }
        if ($level -eq 2) {
            Set-ParaFormat $p '黑体' 'Times New Roman' 14 -1 0 6 0 $true 0
            continue
        }
        if ($level -ge 3) {
            Set-ParaFormat $p '黑体' 'Times New Roman' 12 -1 0 0 0 $true 24
            continue
        }
        if ($style -match '题注' -or $text -match '^(图|表)\s*\d+[\.\d]*') {
            Set-ParaFormat $p '黑体' 'Times New Roman' 12 0 1 0 0 $true 0
            continue
        }
        if ($p.Range.OMaths.Count -gt 0) {
            Set-ParaFormat $p '宋体' 'Times New Roman' 12 0 1 0 0 $true 0
            continue
        }
        if ($p.Range.InlineShapes.Count -gt 0) {
            $p.Format.Alignment = 1
            continue
        }

        Set-ParaFormat $p '宋体' 'Times New Roman' 12 0 3 0 0 $true 24
    }

    foreach ($table in @($doc.Tables)) {
        try {
            $table.Range.Font.NameFarEast = '宋体'
            $table.Range.Font.NameAscii = 'Times New Roman'
            $table.Range.Font.Size = 10.5
            $table.Range.ParagraphFormat.Alignment = 1
        } catch {}
    }

    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        try { $doc.TablesOfContents.Item($i).Update() | Out-Null } catch {}
    }
}

function Find-ParagraphIndex($doc, [string]$pattern) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $text = Clean-Text $doc.Paragraphs.Item($i).Range
        if ($text -match $pattern) { return $i }
    }
    return $null
}

function Starts-NewPage($doc, [int]$idx) {
    if ($idx -le 1) { return $true }
    try {
        $p = $doc.Paragraphs.Item($idx)
        if ($p.Format.PageBreakBefore -eq -1) { return $true }
        $page = $p.Range.Information(3)
        for ($j = $idx - 1; $j -ge 1; $j--) {
            $prevText = Clean-Text $doc.Paragraphs.Item($j).Range
            if ($prevText) {
                $prevPage = $doc.Paragraphs.Item($j).Range.Information(3)
                return ($page -gt $prevPage)
            }
        }
    } catch {}
    return $false
}

function Para-FormatSummary($p) {
    [pscustomobject]@{
        text = Clean-Text $p.Range
        style = Style-Name $p.Range
        font = $p.Range.Font.NameFarEast
        ascii = $p.Range.Font.NameAscii
        size = [Math]::Round([double]$p.Range.Font.Size, 1)
        bold = $p.Range.Font.Bold
        align = $p.Format.Alignment
        before = [Math]::Round([double]$p.Format.SpaceBefore, 1)
        after = [Math]::Round([double]$p.Format.SpaceAfter, 1)
        lineRule = $p.Format.LineSpacingRule
        line = [Math]::Round([double]$p.Format.LineSpacing, 1)
        first = [Math]::Round([double]$p.Format.FirstLineIndent, 1)
        charFirst = [Math]::Round([double]$p.Format.CharacterUnitFirstLineIndent, 1)
    }
}

function Is-BodyHeadingText([string]$text) {
    if (-not $text) { return $false }
    return ($text -match '^([一二三四五六七八九十]+、.*|（[一二三四五六七八九十]+）.*|第[一二三四五六七八九十]+章.*|第[一二三四五六七八九十]+节.*|\d+(\.\d+){0,3}\s+\S.*|前言|绪论|引言|结论|参考文献|致\s*谢|附录\b.*|摘\s*要|Abstract)$')
}

function Get-RefEntryNumber($p) {
    $text = Clean-Text $p.Range
    if ($text -match '^\[(\d+)\]') { return [int]$Matches[1] }
    if ($text -match '^(\d+)[\.、\)]') { return [int]$Matches[1] }
    try {
        $ls = $p.Range.ListFormat.ListString
        if ($ls -match '(\d+)') { return [int]$Matches[1] }
    } catch {}
    return $null
}

function Get-FieldCodes($range) {
    $codes = @()
    try {
        for ($i = 1; $i -le $range.Fields.Count; $i++) {
            $codes += (($range.Fields.Item($i).Code.Text -replace '\s+', ' ').Trim())
        }
    } catch {}
    return ($codes -join ' | ')
}

function Get-HeaderFooterVisibleText($hf) {
    $texts = @()
    try { $texts += (Clean-Text $hf.Range) } catch {}
    try {
        for ($i = 1; $i -le $hf.Shapes.Count; $i++) {
            $shape = $hf.Shapes.Item($i)
            if ($shape.TextFrame.HasText) {
                $texts += (Clean-Text $shape.TextFrame.TextRange)
            }
        }
    } catch {}
    return (($texts | Where-Object { $_ }) -join ' ')
}

function Get-HeaderFooterFieldCodes($hf) {
    $codes = @()
    try {
        $direct = Get-FieldCodes $hf.Range
        if ($direct) { $codes += $direct }
    } catch {}
    try {
        for ($i = 1; $i -le $hf.Shapes.Count; $i++) {
            $shape = $hf.Shapes.Item($i)
            if ($shape.TextFrame.HasText) {
                $shapeCodes = Get-FieldCodes $shape.TextFrame.TextRange
                if ($shapeCodes) { $codes += $shapeCodes }
            }
        }
    } catch {}
    return ($codes -join ' | ')
}

function Expand-CitationNumbers([string]$text) {
    $nums = New-Object System.Collections.Generic.List[int]
    foreach ($m in [regex]::Matches($text, '\[(\d+)(?:\s*[-–]\s*(\d+))?\]')) {
        $a = [int]$m.Groups[1].Value
        if ($m.Groups[2].Success) {
            $b = [int]$m.Groups[2].Value
            if ($b -ge $a -and $b -le 300) {
                for ($n = $a; $n -le $b; $n++) { $nums.Add($n) | Out-Null }
            }
        } else {
            $nums.Add($a) | Out-Null
        }
    }
    return $nums
}

function Is-Fixed22($p) {
    return ($p.Format.LineSpacingRule -eq 4 -and (Near ([double]$p.Format.LineSpacing) 22 0.8))
}

function Is-FirstLineIndent2Chars($p) {
    $pointIndent = $false
    $charIndent = $false
    try { $pointIndent = Near ([double]$p.Format.FirstLineIndent) 24 2.5 } catch {}
    try { $charIndent = Near ([double]$p.Format.CharacterUnitFirstLineIndent) 2 0.35 } catch {}
    return ($pointIndent -or $charIndent)
}

function Font-Contains([string]$actual, [string]$expected) {
    if ([string]::IsNullOrWhiteSpace($actual)) { return $false }
    return ($actual -like "*$expected*")
}

function Font-IsSongTimes12($p) {
    $eastAsiaOk = Font-Contains $p.Range.Font.NameFarEast '宋体'
    $asciiOk = Font-Contains $p.Range.Font.NameAscii 'Times New Roman'
    $sizeOk = Near ([double]$p.Range.Font.Size) 12 0.6
    return ($eastAsiaOk -and $asciiOk -and $sizeOk)
}

function Font-IsHeiTimes($p, [double]$size) {
    $eastAsiaOk = Font-Contains $p.Range.Font.NameFarEast '黑体'
    $asciiOk = Font-Contains $p.Range.Font.NameAscii 'Times New Roman'
    $sizeOk = Near ([double]$p.Range.Font.Size) $size 0.6
    return ($eastAsiaOk -and $asciiOk -and $sizeOk)
}

function Is-StaticTocLine([string]$text) {
    if (-not $text) { return $false }
    return ($text -match '\.{3,}|…{2,}|[·.]{3,}\s*[0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXLC]+$')
}

function Resolve-YnuHeadingSystem($doc, [string]$preference = 'Auto') {
    if ($preference -ne 'Auto') { return $preference }
    $secondTop = 0
    $thirdTop = 0
    $firstTop = 0
    $tocRanges = @()
    try {
        for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
            $tocRanges += $doc.TablesOfContents.Item($i).Range
        }
    } catch {}

    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if (Is-InRange $p.Range $tocRanges) { continue }
        $text = Clean-Text $p.Range
        if (-not $text -or (Is-StaticTocLine $text)) { continue }
        if ($text -match '^第[一二三四五六七八九十]+章') { $secondTop++ }
        elseif ($text -match '^\d+(\s|　|□)+\S') { $thirdTop++ }
        elseif ($text -match '^[一二三四五六七八九十]+、') { $firstTop++ }
    }

    $topKinds = @()
    if ($secondTop -gt 0) { $topKinds += 'Second' }
    if ($thirdTop -gt 0) { $topKinds += 'Third' }
    if ($firstTop -gt 0 -and $secondTop -eq 0) { $topKinds += 'First' }
    if ($topKinds.Count -gt 1) { return 'Mixed' }
    if ($topKinds.Count -eq 1) { return $topKinds[0] }
    return 'Auto'
}

function Get-HeadingLevelGuess([string]$style, [string]$text) {
    if (-not $text -or (Is-StaticTocLine $text)) {
        if ($style -match '标题\s*([1-9])') { return [int]$Matches[1] }
        return 0
    }
    if ($text -match '^(摘\s*要|Abstract|参考文献|致\s*谢|附录\b|前言|绪论|引言|结论$)') { return 1 }

    $system = $script:YnuHeadingSystem
    if (-not $system) { $system = 'Auto' }

    if ($system -eq 'Second') {
        if ($text -match '^第[一二三四五六七八九十]+章') { return 1 }
        if ($text -match '^第[一二三四五六七八九十]+节') { return 2 }
        if ($text -match '^[一二三四五六七八九十]+、') { return 3 }
        if ($text -match '^（[一二三四五六七八九十]+）') { return 3 }
        return 0
    }
    if ($system -eq 'Third') {
        if ($text -match '^\d+(\s|　|□)+\S') { return 1 }
        if ($text -match '^\d+\.\d+(\s|　|□)+\S') { return 2 }
        if ($text -match '^\d+\.\d+\.\d+') { return 3 }
        if ($text -match '^\d+\.\d+\.\d+\.\d+') { return 3 }
        return 0
    }
    if ($system -eq 'First') {
        if ($text -match '^[一二三四五六七八九十]+、') { return 1 }
        if ($text -match '^（[一二三四五六七八九十]+）') { return 2 }
        if ($text -match '^\d+\.') { return 3 }
        if ($text -match '^（\d+）') { return 3 }
        return 0
    }

    if ($text -match '^第[一二三四五六七八九十]+章') { return 1 }
    if ($text -match '^第[一二三四五六七八九十]+节') { return 2 }
    if ($text -match '^\d+\.\d+\.\d+') { return 3 }
    if ($text -match '^\d+\.\d+(\s|　|□)+\S') { return 2 }
    if ($text -match '^\d+(\s|　|□)+\S') { return 1 }
    if ($text -match '^[一二三四五六七八九十]+、') { return 1 }
    if ($style -match '标题\s*([1-9])') { return [int]$Matches[1] }
    return 0
}

function Is-ReferenceLikeParagraph([string]$text) {
    if (-not $text) { return $false }
    return ($text -match '^\s*(\[\d+\]|\d+[\.、\)])\s*')
}

function Get-NearestParagraphBefore($doc, [int]$position) {
    $best = $null
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if ($p.Range.End -le $position) {
            if (Clean-Text $p.Range) { $best = $p }
        } elseif ($p.Range.Start -gt $position) {
            break
        }
    }
    return $best
}

function Get-NearestParagraphAfter($doc, [int]$position) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if ($p.Range.Start -ge $position -and (Clean-Text $p.Range)) { return $p }
    }
    return $null
}

function Audit-YnuDocument($doc, [string]$headingPreference, [string]$path, [string]$backupPath, [string]$outputPath, [bool]$contentPreserved) {
    $checks = New-Object System.Collections.Generic.List[object]
    try {
        $doc.Fields.Update() | Out-Null
        $doc.Repaginate()
    } catch {}
    $script:YnuHeadingSystem = Resolve-YnuHeadingSystem $doc $headingPreference

    $pageSetupBad = @()
    foreach ($section in @($doc.Sections)) {
        $ps = $section.PageSetup
        if (-not ((Near $ps.TopMargin (CmToPt 2.5)) -and (Near $ps.BottomMargin (CmToPt 2.0)) -and (Near $ps.LeftMargin (CmToPt 3.0)) -and (Near $ps.RightMargin (CmToPt 2.0)) -and (Near $ps.PageWidth 595.3 2) -and (Near $ps.PageHeight 841.9 2))) {
            $pageSetupBad += $section.Index
        }
    }
    Add-Check $checks '页面' 'A4；页边距上2.5cm、下2.0cm、左3.0cm、右2.0cm' ($(if ($pageSetupBad.Count -eq 0) { 'PASS' } else { 'FAIL' })) "badSections=$($pageSetupBad -join ',')"

    $allText = $doc.Content.Text
    $hasAbstractCn = $allText -match '摘\s*要'
    $hasAbstractEn = $allText -match 'Abstract'
    $hasToc = $doc.TablesOfContents.Count -gt 0
    $hasRefs = $allText -match '参考文献'
    $hasAck = $allText -match '致\s*谢'
    Add-Check $checks '结构' '摘要、Abstract、目录、正文、参考文献、致谢等结构齐全' ($(if ($hasAbstractCn -and $hasAbstractEn -and $hasToc -and $hasRefs -and $hasAck) { 'PASS' } else { 'WARN' })) "摘要=$hasAbstractCn; Abstract=$hasAbstractEn; TOC=$hasToc; 参考文献=$hasRefs; 致谢=$hasAck"

    $absCnIdx = Find-ParagraphIndex $doc '^摘\s*要$'
    $absEnIdx = Find-ParagraphIndex $doc '^Abstract$'
    $kwCnIdx = Find-ParagraphIndex $doc '^关键词[:：]'
    $kwEnIdx = Find-ParagraphIndex $doc '^Key\s*Words[:：]'
    $absIssues = @()
    if ($null -eq $absCnIdx) { $absIssues += '缺少中文摘要标题' }
    else {
        $p = $doc.Paragraphs.Item($absCnIdx)
        if (-not ((Near ([double]$p.Range.Font.Size) 16 0.6) -and $p.Format.Alignment -eq 1)) { $absIssues += "中文摘要标题格式:$absCnIdx" }
        if (-not (Starts-NewPage $doc $absCnIdx)) { $absIssues += "中文摘要未另起页:$absCnIdx" }
    }
    if ($null -eq $absEnIdx) { $absIssues += '缺少Abstract标题' }
    else {
        $p = $doc.Paragraphs.Item($absEnIdx)
        if (-not ((Near ([double]$p.Range.Font.Size) 16 0.6) -and $p.Format.Alignment -eq 1)) { $absIssues += "Abstract标题格式:$absEnIdx" }
        if (-not (Starts-NewPage $doc $absEnIdx)) { $absIssues += "Abstract未另起页:$absEnIdx" }
    }
    if ($null -eq $kwCnIdx) { $absIssues += '缺少关键词' }
    else {
        $t = Clean-Text $doc.Paragraphs.Item($kwCnIdx).Range
        $terms = (($t -replace '^关键词[:：]', '') -split '；' | Where-Object { $_.Trim() })
        if ($terms.Count -lt 3 -or $terms.Count -gt 5) { $absIssues += "关键词数量=$($terms.Count)" }
        if ($t -match '[。.;；]$') { $absIssues += '关键词末尾有标点' }
    }
    if ($null -eq $kwEnIdx) { $absIssues += '缺少Key Words' }
    else {
        $t = Clean-Text $doc.Paragraphs.Item($kwEnIdx).Range
        $terms = (($t -replace '^Key\s*Words[:：]', '') -split ';' | Where-Object { $_.Trim() })
        if ($terms.Count -lt 3 -or $terms.Count -gt 5) { $absIssues += "KeyWords数量=$($terms.Count)" }
        if ($t -match '[。.;；]$') { $absIssues += 'KeyWords末尾有标点' }
    }
    Add-Check $checks '摘要关键词' '摘要/Abstract另起页；标题三号居中；关键词3-5个、分隔符和末尾标点正确' ($(if ($absIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "issues=$($absIssues -join '; ')"

    $prelimSections = New-Object System.Collections.Generic.HashSet[int]
    foreach ($idx in @($absCnIdx, $absEnIdx, (Find-ParagraphIndex $doc '^目\s*录$'))) {
        if ($null -ne $idx) {
            try { $prelimSections.Add([int]$doc.Paragraphs.Item($idx).Range.Information(2)) | Out-Null } catch {}
        }
    }
    $bodySectionNumber = $null
    $tocRangesForBody = @()
    $tocEndForBody = $null
    try {
        for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
            $tocRange = $doc.TablesOfContents.Item($i).Range
            $tocRangesForBody += $tocRange
            if ($null -eq $tocEndForBody -or $tocRange.End -gt $tocEndForBody) { $tocEndForBody = $tocRange.End }
        }
    } catch {}
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if ($null -ne $tocEndForBody -and $p.Range.Start -lt $tocEndForBody) { continue }
        if (Is-InRange $p.Range $tocRangesForBody) { continue }
        $text = Clean-Text $p.Range
        if (-not $text -or $text -match '^(摘\s*要|Abstract|目\s*录)$') { continue }
        if ((Get-HeadingLevelGuess (Style-Name $p.Range) $text) -eq 1) {
            $bodySectionNumber = [int]$p.Range.Information(2)
            break
        }
    }

    $hfIssues = @()
    for ($s = 1; $s -le $doc.Sections.Count; $s++) {
        $section = $doc.Sections.Item($s)
        try {
            $headerObj = $section.Headers.Item(1)
            $footerObj = $section.Footers.Item(1)
            $header = $headerObj.Range
            $footer = $footerObj.Range
            $hText = Get-HeaderFooterVisibleText $headerObj
            $fText = Get-HeaderFooterVisibleText $footerObj
            $hCodes = Get-HeaderFooterFieldCodes $headerObj
            $fCodes = Get-HeaderFooterFieldCodes $footerObj
            $isPrelim = $prelimSections.Contains($s)
            $isBody = ($null -ne $bodySectionNumber -and $s -ge $bodySectionNumber)

            if ($isBody) {
                if (-not ($hText -or $hCodes -match 'STYLEREF')) { $hfIssues += "section${s}:页眉空" }
                if ($hText -or $hCodes) {
                    if (-not ((Near ([double]$header.Font.Size) 10.5 0.6) -and $header.ParagraphFormat.Alignment -eq 1)) { $hfIssues += "section${s}:页眉格式" }
                }
                if (-not ($fText -match '第\s*\d+\s*页\s*（共\s*\d+\s*页）' -or ($fCodes -match 'PAGE' -and $fCodes -match 'SECTIONPAGES'))) { $hfIssues += "section${s}:正文页码缺失或格式错误" }
                elseif (-not ((Near ([double]$footer.Font.Size) 10.5 0.6) -and $footer.ParagraphFormat.Alignment -eq 1)) { $hfIssues += "section${s}:页码格式" }
            } elseif ($isPrelim) {
                if (-not ($fText -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXLC]+$' -or ($fCodes -match 'PAGE' -and $fCodes -match 'ROMAN'))) { $hfIssues += "section${s}:摘要/目录罗马页码缺失" }
                if ($fText -match '^第\s*\d+\s*页' -or $fCodes -match 'SECTIONPAGES|NUMPAGES') { $hfIssues += "section${s}:摘要/目录误用正文页码" }
                if ($fText -or $fCodes) {
                    if (-not ((Near ([double]$footer.Font.Size) 10.5 0.6) -and $footer.ParagraphFormat.Alignment -eq 1)) { $hfIssues += "section${s}:页码格式" }
                }
            } else {
                if ($fText -or $fCodes -match 'PAGE') { $hfIssues += "section${s}:封面/前置页不应编页码" }
            }
        } catch {
            $hfIssues += "section${s}:页眉页脚读取失败"
        }
    }
    Add-Check $checks '页眉页码' '通用检查不强制封面/评价表/声明页；摘要/目录罗马数字；正文第×页（共×页）；正文页眉为当前一级标题' ($(if ($hfIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "issues=$($hfIssues -join '; ')" '封面、评价表、声明页等仅在用户提供模板或明确要求时核查；页码分节语义复杂时需结合模板人工复核。'

    $headingTocRanges = @()
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) { $headingTocRanges += $doc.TablesOfContents.Item($i).Range }

    $headingRows = @()
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if (Is-InRange $p.Range $headingTocRanges) { continue }
        $style = Style-Name $p.Range
        if ($style -match 'TOC\s*\d|目录\s*\d') { continue }
        $text = Clean-Text $p.Range
        $level = Get-HeadingLevelGuess $style $text
        if ($level -gt 0 -and $text) {
            $headingRows += [pscustomobject]@{
                index = $i
                style = $style
                level = $level
                text = $text
                font = $p.Range.Font.NameFarEast
                ascii = $p.Range.Font.NameAscii
                size = [double]$p.Range.Font.Size
                bold = $p.Range.Font.Bold
                align = $p.Format.Alignment
                before = [double]$p.Format.SpaceBefore
                after = [double]$p.Format.SpaceAfter
                lineRule = $p.Format.LineSpacingRule
                line = [double]$p.Format.LineSpacing
            }
        }
    }
    $detectedSystem = $script:YnuHeadingSystem
    $systemOk = ($detectedSystem -ne 'Mixed')
    if ($systemOk -and $headingPreference -ne 'Auto') {
        $systemOk = ($detectedSystem -eq $headingPreference)
    }
    Add-Check $checks '题序体系' '选定一种题序体系且前后一致；不得混用；第二种题序允许“第一章/第一节/一、/（一）”同用' ($(if ($systemOk) { 'PASS' } else { 'FAIL' })) "detected=$detectedSystem; preference=$headingPreference" '脚本不会自动改写题序文字；混用时需人工确认后再改。'

    $badHeadingFormat = @()
    foreach ($h in $headingRows) {
        if ($h.level -eq 1) {
            if (-not ((Near $h.size 16 0.6) -and $h.align -eq 1 -and (Font-Contains $h.font '黑体') -and $h.bold -ne 0 -and (Near $h.before 9.6 2.2) -and (Near $h.after 6 2.2) -and $h.lineRule -eq 4 -and (Near $h.line 22 0.8))) { $badHeadingFormat += $h.index }
        } elseif ($h.level -eq 2) {
            if (-not ((Near $h.size 14 0.6) -and (Font-Contains $h.font '黑体') -and $h.bold -ne 0 -and (Near $h.before 6 2.2) -and (Near $h.after 0 1.2) -and $h.lineRule -eq 4 -and (Near $h.line 22 0.8))) { $badHeadingFormat += $h.index }
        } elseif ($h.level -ge 3) {
            if (-not ((Near $h.size 12 0.6) -and (Font-Contains $h.font '黑体') -and $h.bold -ne 0 -and (Near $h.before 0 1.2) -and (Near $h.after 0 1.2) -and $h.lineRule -eq 4 -and (Near $h.line 22 0.8))) { $badHeadingFormat += $h.index }
        }
    }
    Add-Check $checks '标题格式' '一级三号黑体居中且段前0.8行段后0.5行；二级四号黑体；三级及以下小四黑体；固定22磅' ($(if ($badHeadingFormat.Count -eq 0) { 'PASS' } else { 'WARN' })) "badHeadingParagraphs=$($badHeadingFormat -join ',')"

    $chapterBreakBad = @()
    foreach ($h in ($headingRows | Where-Object { $_.level -eq 1 -and $_.text -notmatch '^(摘\s*要|Abstract|目\s*录)$' })) {
        if (-not (Starts-NewPage $doc $h.index)) { $chapterBreakBad += "$($h.index):$($h.text)" }
        if ($chapterBreakBad.Count -ge 12) { break }
    }
    Add-Check $checks '章节分页' '摘要、目录、正文各章、参考文献、致谢、附录等独立部分应另起页' ($(if ($chapterBreakBad.Count -eq 0) { 'PASS' } else { 'WARN' })) "notNewPage=$($chapterBreakBad -join '; ')"

    $tocIssues = @()
    $tocTitleIdx = Find-ParagraphIndex $doc '^目\s*录$'
    if ($null -eq $tocTitleIdx) {
        $tocIssues += '缺少目录标题'
    } else {
        $p = $doc.Paragraphs.Item($tocTitleIdx)
        if (-not ((Near ([double]$p.Range.Font.Size) 16 0.6) -and $p.Format.Alignment -eq 1 -and (Font-Contains $p.Range.Font.NameFarEast '黑体'))) { $tocIssues += "目录标题格式:$tocTitleIdx" }
        if (-not (Starts-NewPage $doc $tocTitleIdx)) { $tocIssues += "目录未另起页:$tocTitleIdx" }
    }
    if ($doc.TablesOfContents.Count -eq 0) {
        $tocIssues += '缺少Word自动目录'
    } else {
        $toc = $doc.TablesOfContents.Item(1)
        for ($ti = 1; $ti -le $toc.Range.Paragraphs.Count; $ti++) {
            $tp = $toc.Range.Paragraphs.Item($ti)
            $line = Clean-Text $tp.Range
            if (-not $line) { continue }
            $style = Style-Name $tp.Range
            $entry = ($line -replace '\s*[0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXLC]+$', '').Trim()
            if ($entry.Length -gt 120 -or (-not (Is-BodyHeadingText $entry))) {
                $tocIssues += "疑似正文混入目录:$line"
            }
            if ($style -match 'TOC\s*1') {
                if (-not ((Near ([double]$tp.Range.Font.Size) 12 0.6) -and (Font-Contains $tp.Range.Font.NameFarEast '黑体') -and (Is-Fixed22 $tp))) { $tocIssues += "TOC1格式:$line" }
            } elseif ($style -match 'TOC\s*2') {
                if (-not ((Near ([double]$tp.Range.Font.Size) 12 0.6) -and (Font-Contains $tp.Range.Font.NameFarEast '宋体') -and (Is-Fixed22 $tp) -and ((Near ([double]$tp.Format.FirstLineIndent) 24 3) -or (Near ([double]$tp.Format.LeftIndent) 24 3)))) { $tocIssues += "TOC2格式:$line" }
            } elseif ($style -match 'TOC\s*3') {
                if (-not ((Near ([double]$tp.Range.Font.Size) 12 0.6) -and (Font-Contains $tp.Range.Font.NameFarEast '宋体') -and (Is-Fixed22 $tp) -and ((Near ([double]$tp.Format.FirstLineIndent) 48 4) -or (Near ([double]$tp.Format.LeftIndent) 48 4)))) { $tocIssues += "TOC3格式:$line" }
            }
            if ($line -notmatch '[0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXLC]+$') { $tocIssues += "目录页码缺失:$line" }
        }
    }
    Add-Check $checks '目录' '目录自动生成；标题三号黑体居中；包含至三级标题；点引导符和右页码；不得混入正文段落' ($(if ($tocIssues.Count -eq 0) { 'PASS' } else { 'FAIL' })) "issues=$($tocIssues -join '; ')"

    $tocRanges = @()
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) { $tocRanges += $doc.TablesOfContents.Item($i).Range }
    $refsIdxForBody = Find-ParagraphIndex $doc '^参考文献$'
    $ackIdxForBody = Find-ParagraphIndex $doc '^致\s*谢$'
    $bodyBad = @()
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $text = Clean-Text $p.Range
        if (-not $text) { continue }
        if (Is-InRange $p.Range $tocRanges) { continue }
        if ($text -match '^目\s*录$') { continue }
        if ($null -ne $refsIdxForBody -and $i -gt $refsIdxForBody -and ($null -eq $ackIdxForBody -or $i -lt $ackIdxForBody)) { continue }
        if ((Style-Name $p.Range) -match '标题|题注|TOC|目录') { continue }
        if ((Get-HeadingLevelGuess (Style-Name $p.Range) $text) -gt 0) { continue }
        if ($p.Range.OMaths.Count -gt 0 -or $p.Range.InlineShapes.Count -gt 0) { continue }
        if ($text -match '^(参考文献|致\s*谢|摘\s*要|Abstract|图\s*\d|表\s*\d)' -or (Is-ReferenceLikeParagraph $text)) { continue }
        if (-not ((Font-IsSongTimes12 $p) -and (Is-Fixed22 $p) -and $p.Format.Alignment -eq 3 -and (Is-FirstLineIndent2Chars $p))) {
            $bodyBad += $i
            if ($bodyBad.Count -ge 20) { break }
        }
    }
    Add-Check $checks '正文' '小四宋体/Times New Roman；固定22磅；两端对齐；首行缩进2字符' ($(if ($bodyBad.Count -eq 0) { 'PASS' } else { 'WARN' })) "sampleBadParagraphs=$($bodyBad -join ',')"

    $refsIdx = Find-ParagraphIndex $doc '^参考文献$'
    $ackIdx = Find-ParagraphIndex $doc '^致\s*谢$'

    $citationIssues = @()
    $citationOrder = New-Object System.Collections.Generic.List[int]
    $seenCitations = @{}
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        if ($null -ne $refsIdx -and $i -ge $refsIdx) { break }
        $p = $doc.Paragraphs.Item($i)
        if (Is-InRange $p.Range $tocRanges) { continue }
        $text = Clean-Text $p.Range
        if (-not $text) { continue }
        if ((Get-HeadingLevelGuess (Style-Name $p.Range) $text) -gt 0 -and $text -match '\[\d+') {
            $citationIssues += "标题含引文标注:$i"
        }
        foreach ($n in (Expand-CitationNumbers $text)) {
            if (-not $seenCitations.ContainsKey($n)) {
                $seenCitations[$n] = $true
                $citationOrder.Add($n) | Out-Null
            }
        }
    }
    $expectedCitation = 1
    foreach ($n in $citationOrder) {
        if ($n -ne $expectedCitation) {
            $citationIssues += "首次出现顺序异常:expect$expectedCitation got$n"
            if ($n -gt $expectedCitation) { $expectedCitation = $n + 1 }
        } else {
            $expectedCitation++
        }
        if ($citationIssues.Count -ge 20) { break }
    }
    Add-Check $checks '引文标注' '顺序编码制；首次出现按[1]起连续编号；标题中不得含引文标注' ($(if ($citationIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "issues=$($citationIssues -join '; ')"

    $errorFields = @()
    for ($i = 1; $i -le $doc.Fields.Count; $i++) {
        $res = Clean-Text $doc.Fields.Item($i).Result
        if ($res -match '错误|Error|引用源|Reference source|不能识别') {
            $errorFields += $i
        }
    }
    Add-Check $checks '交叉引用' '不得出现损坏的 REF/PAGEREF/SEQ 等字段' ($(if ($errorFields.Count -eq 0) { 'PASS' } else { 'FAIL' })) "badFields=$($errorFields -join ',')"

    $figureIssues = @()
    $tableIssues = @()
    $figNums = @{}
    $tblNums = @{}
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $text = Clean-Text $p.Range
        if ($text -match '^图\s*([0-9]+(\.[0-9]+)+)\s+\S') {
            $num = $Matches[1]
            if ($figNums.ContainsKey($num)) { $figureIssues += "图号重复:$num" } else { $figNums[$num] = $true }
            if (-not ((Near ([double]$p.Range.Font.Size) 12 0.6) -and (Font-Contains $p.Range.Font.NameFarEast '黑体') -and $p.Format.Alignment -eq 1)) { $figureIssues += "图题格式:$i" }
        } elseif ($text -match '^图\s*\d+') {
            $figureIssues += "图题编号非按章编号:$i"
        }
        if ($text -match '^表\s*([0-9]+(\.[0-9]+)+)\s+\S') {
            $num = $Matches[1]
            if ($tblNums.ContainsKey($num)) { $tableIssues += "表号重复:$num" } else { $tblNums[$num] = $true }
            if (-not ((Near ([double]$p.Range.Font.Size) 12 0.6) -and (Font-Contains $p.Range.Font.NameFarEast '黑体') -and $p.Format.Alignment -eq 1)) { $tableIssues += "表题格式:$i" }
        } elseif ($text -match '^表\s*\d+') {
            $tableIssues += "表题编号非按章编号:$i"
        }
    }
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        $cap = Get-NearestParagraphAfter $doc $shape.Range.End
        $capText = if ($cap) { Clean-Text $cap.Range } else { '' }
        if ($capText -notmatch '^图\s*[0-9]+(\.[0-9]+)+\s+\S') { $figureIssues += "图片${i}:缺少下方图题或编号格式错误" }
        if ($figureIssues.Count -ge 30) { break }
    }
    for ($i = 1; $i -le $doc.Tables.Count; $i++) {
        $table = $doc.Tables.Item($i)
        $cap = Get-NearestParagraphBefore $doc $table.Range.Start
        $capText = if ($cap) { Clean-Text $cap.Range } else { '' }
        if ($capText -notmatch '^表\s*[0-9]+(\.[0-9]+)+\s+\S') { $tableIssues += "表格${i}:缺少上方表题或编号格式错误" }
        if (-not ((Near ([double]$table.Range.Font.Size) 10.5 0.8) -and (Font-Contains $table.Range.Font.NameFarEast '宋体'))) { $tableIssues += "表格${i}:表内字号字体异常" }
        if ($tableIssues.Count -ge 30) { break }
    }
    Add-Check $checks '图题表题' '图题置于图下、表题置于表上；按章编号；居中；黑体小四；表内五号宋体/Times New Roman' ($(if ($figureIssues.Count -eq 0 -and $tableIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "figureIssues=$($figureIssues -join '; '); tableIssues=$($tableIssues -join '; ')"

    $formulaBad = @()
    $formulaCandidateCount = 0
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $text = Clean-Text $p.Range
        if (-not $text -and $p.Range.OMaths.Count -eq 0) { continue }
        $hasMath = $p.Range.OMaths.Count -gt 0
        $looksFormula = $hasMath -or (($text.Length -le 160) -and ($text -match '[=∑Σ√∫]|[A-Za-z]\s*[_=]') -and ($text -match '[A-Za-z0-9α-ωΑ-Ω]') -and ($text -notmatch '^(图|表|关键词|Key\s*Words|参考文献|致\s*谢|摘\s*要|Abstract)'))
        if ($looksFormula) {
            $formulaCandidateCount++
            $hasNumber = ($text -match '[（(]\d+(\.\d+)?[)）]\s*$')
            $hasRightLayout = ($p.Format.Alignment -eq 1 -or $p.Range.ParagraphFormat.TabStops.Count -gt 0)
            if (-not $hasMath) { $formulaBad += "p${i}:非Word公式对象" }
            if (-not $hasRightLayout) { $formulaBad += "p${i}:未居中/缺右侧编号制表位" }
            if (-not $hasNumber) { $formulaBad += "p${i}:缺少按章公式编号" }
            if ($text -match '\.{3,}|…{2,}') { $formulaBad += "p${i}:公式与编号之间有虚线/点引导" }
            if ($text -match '/') { $formulaBad += "p${i}:疑似斜杠分式，宜改为上下结构" }
            if ($text -match '[。；，,]\s*\S') { $formulaBad += "p${i}:疑似未独立成行" }
        }
    }
    Add-Check $checks '公式' 'Word公式对象；独立一行；主体居中；编号右对齐；按章编号；无虚线；分式优先上下结构' ($(if ($formulaBad.Count -eq 0) { 'PASS' } else { 'WARN' })) "officeMathObjects=$($doc.OMaths.Count); candidateParagraphs=$formulaCandidateCount; issues=$($formulaBad -join '; ')" '公式语义和编号文字不自动改写。'

    $refIssues = @()
    $refNums = New-Object System.Collections.Generic.List[int]
    $refEntryCount = 0
    if ($null -eq $refsIdx) {
        $refIssues += '缺少参考文献标题'
    } else {
        $p = $doc.Paragraphs.Item($refsIdx)
        if (-not (Starts-NewPage $doc $refsIdx)) { $refIssues += "参考文献未另起页:$refsIdx" }
        if (-not ((Near ([double]$p.Range.Font.Size) 16 0.6) -and $p.Format.Alignment -eq 1 -and (Font-Contains $p.Range.Font.NameFarEast '黑体'))) { $refIssues += "参考文献标题格式:$refsIdx" }
        $endIdx = $doc.Paragraphs.Count
        if ($null -ne $ackIdx -and $ackIdx -gt $refsIdx) { $endIdx = $ackIdx - 1 }
        for ($i = $refsIdx + 1; $i -le $endIdx; $i++) {
            $rp = $doc.Paragraphs.Item($i)
            $text = Clean-Text $rp.Range
            if (-not $text) { continue }
            if ($text -match '^(附录\b|致\s*谢)$') { break }
            $refEntryCount++
            $num = Get-RefEntryNumber $rp
            if ($null -eq $num) { $refIssues += "参考文献缺编号:$i" } else { $refNums.Add($num) | Out-Null }
            if (-not ((Font-IsSongTimes12 $rp) -and (Is-Fixed22 $rp) -and $rp.Format.Alignment -eq 3)) { $refIssues += "参考文献条目格式:$i" }
            if ($refIssues.Count -ge 30) { break }
        }
        if ($refEntryCount -lt 10) { $refIssues += "参考文献少于10篇:$refEntryCount" }
        for ($j = 0; $j -lt $refNums.Count; $j++) {
            if ($refNums[$j] -ne ($j + 1)) { $refIssues += "参考文献编号不连续:pos$($j + 1)=$($refNums[$j])"; break }
        }
        foreach ($n in $citationOrder) {
            if ($refNums.Count -gt 0 -and -not ($refNums.Contains($n))) { $refIssues += "正文引文[$n]未匹配参考文献编号"; break }
        }
    }
    Add-Check $checks '参考文献' '另起页；标题三号黑体居中；不少于10篇；顺序编码与正文对应；条目小四固定22磅两端对齐；GB/T 7714-2015' ($(if ($refIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "entries=$refEntryCount; issues=$($refIssues -join '; ')" '脚本只能核查编号、格式和明显结构；著录项事实准确性需人工或专门工具复核。'

    $ackIssues = @()
    if ($null -eq $ackIdx) {
        $ackIssues += '缺少致谢'
    } else {
        $p = $doc.Paragraphs.Item($ackIdx)
        if (-not (Starts-NewPage $doc $ackIdx)) { $ackIssues += "致谢未另起页:$ackIdx" }
        if (-not ((Near ([double]$p.Range.Font.Size) 16 0.6) -and $p.Format.Alignment -eq 1 -and (Font-Contains $p.Range.Font.NameFarEast '黑体'))) { $ackIssues += "致谢标题格式:$ackIdx" }
        for ($i = $ackIdx + 1; $i -le $doc.Paragraphs.Count; $i++) {
            $ap = $doc.Paragraphs.Item($i)
            $text = Clean-Text $ap.Range
            if (-not $text) { continue }
            if ($text -match '^附录\b') { break }
            if (-not ((Font-IsSongTimes12 $ap) -and (Is-Fixed22 $ap) -and $ap.Format.Alignment -eq 3 -and (Is-FirstLineIndent2Chars $ap))) { $ackIssues += "致谢正文格式:$i" }
            if ($ackIssues.Count -ge 10) { break }
        }
    }
    Add-Check $checks '致谢' '致谢另起页；标题三号黑体居中；正文小四宋体/Times New Roman、固定22磅、首行缩进2字符' ($(if ($ackIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "issues=$($ackIssues -join '; ')"

    $appendixIssues = @()
    $appendixIdx = Find-ParagraphIndex $doc '^附录\b'
    if ($null -ne $appendixIdx) {
        $p = $doc.Paragraphs.Item($appendixIdx)
        if (-not (Starts-NewPage $doc $appendixIdx)) { $appendixIssues += "附录未另起页:$appendixIdx" }
        if (-not ((Near ([double]$p.Range.Font.Size) 16 0.6) -and $p.Format.Alignment -eq 1 -and (Font-Contains $p.Range.Font.NameFarEast '黑体'))) { $appendixIssues += "附录标题格式:$appendixIdx" }
    }
    Add-Check $checks '附录' '附录如存在应另起页，标题按一级标题格式；无附录时不强制添加' ($(if ($appendixIssues.Count -eq 0) { 'PASS' } else { 'WARN' })) "issues=$($appendixIssues -join '; ')"

    Add-Check $checks '内容保护' '格式调整版不得修改论文内容文字' ($(if ($contentPreserved) { 'PASS' } else { 'FAIL' })) "contentPreserved=$contentPreserved"

    $pass = @($checks | Where-Object { $_.status -eq 'PASS' }).Count
    $warn = @($checks | Where-Object { $_.status -eq 'WARN' }).Count
    $fail = @($checks | Where-Object { $_.status -eq 'FAIL' }).Count

    return [pscustomobject]@{
        path = $path
        backup = $backupPath
        output = $outputPath
        mode = $Mode
        generatedAt = (Get-Date).ToString('s')
        pages = $doc.ComputeStatistics(2)
        paragraphs = $doc.Paragraphs.Count
        sections = $doc.Sections.Count
        fields = $doc.Fields.Count
        formulas = $doc.OMaths.Count
        tables = $doc.Tables.Count
        inlineShapes = $doc.InlineShapes.Count
        summary = [pscustomobject]@{ PASS = $pass; WARN = $warn; FAIL = $fail }
        checks = $checks
    }
}

function Write-MarkdownReport($audit, [string]$path) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# 云南大学本科生学年/毕业论文格式严格核查报告') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("- 文件：$($audit.path)") | Out-Null
    if ($audit.backup) { $lines.Add("- 备份：$($audit.backup)") | Out-Null }
    if ($audit.output) { $lines.Add("- 格式调整版：$($audit.output)") | Out-Null }
    $lines.Add("- 生成时间：$($audit.generatedAt)") | Out-Null
    $lines.Add("- 页数：$($audit.pages)；段落：$($audit.paragraphs)；公式：$($audit.formulas)；表格：$($audit.tables)；图片：$($audit.inlineShapes)") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| 类别 | 要求 | 结果 | 状态 | 备注 |') | Out-Null
    $lines.Add('|---|---|---|---|---|') | Out-Null
    foreach ($c in $audit.checks) {
        $cat = ($c.category -replace '\|', '/')
        $req = ($c.requirement -replace '\|', '/')
        $res = ($c.result -replace '\|', '/' -replace "`r?`n", ' ')
        $note = ($c.note -replace '\|', '/' -replace "`r?`n", ' ')
        $lines.Add("| $cat | $req | $res | $($c.status) | $note |") | Out-Null
    }
    $lines.Add('') | Out-Null
    $lines.Add("## 汇总") | Out-Null
    $lines.Add("PASS=$($audit.summary.PASS); WARN=$($audit.summary.WARN); FAIL=$($audit.summary.FAIL)") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('说明：本报告只检查和处理格式问题；允许修复既有 Word 自动域/交叉引用结构并刷新字段结果。涉及文字增删、题序文字改写、手工引文编号、参考文献条目事实准确性、公式语义、图表数据等内容问题，仅提出风险，不自动修改。') | Out-Null
    $lines.Add('封面、评价表、原创性声明、授权页等前置表单不作为通用强制项；仅在用户提供模板或明确要求时核查。') | Out-Null
    $lines | Set-Content -LiteralPath $path -Encoding UTF8
}

function Add-ReportParagraph($selection, [string]$text, [double]$size, [int]$bold, [int]$align) {
    $selection.Font.NameFarEast = '宋体'
    $selection.Font.NameAscii = 'Times New Roman'
    $selection.Font.NameOther = 'Times New Roman'
    $selection.Font.Size = $size
    $selection.Font.Bold = $bold
    $selection.ParagraphFormat.Alignment = $align
    $selection.ParagraphFormat.LineSpacingRule = 4
    $selection.ParagraphFormat.LineSpacing = 16
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Split-MarkdownTableRow([string]$line) {
    $trimmed = $line.Trim()
    if ($trimmed.StartsWith('|')) { $trimmed = $trimmed.Substring(1) }
    if ($trimmed.EndsWith('|')) { $trimmed = $trimmed.Substring(0, $trimmed.Length - 1) }
    return @($trimmed -split '\|' | ForEach-Object { $_.Trim() })
}

function Add-ReportTable($doc, $selection, [string[]]$rows) {
    if ($rows.Count -lt 2) { return }
    $dataRows = @()
    foreach ($row in $rows) {
        if ($row -match '^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$') { continue }
        $cells = @(Split-MarkdownTableRow $row)
        if ($cells.Count -gt 0) { $dataRows += ,$cells }
    }
    if ($dataRows.Count -eq 0) { return }

    $cols = 0
    foreach ($r in $dataRows) { if ($r.Count -gt $cols) { $cols = $r.Count } }
    $range = $selection.Range
    $table = $doc.Tables.Add($range, $dataRows.Count, $cols)
    $table.Borders.Enable = 1
    $table.Range.Font.NameFarEast = '宋体'
    $table.Range.Font.NameAscii = 'Times New Roman'
    $table.Range.Font.Size = 8
    $table.Range.ParagraphFormat.LineSpacingRule = 0
    for ($r = 0; $r -lt $dataRows.Count; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            $value = ''
            if ($c -lt $dataRows[$r].Count) { $value = $dataRows[$r][$c] }
            $table.Cell($r + 1, $c + 1).Range.Text = $value
            if ($r -eq 0) { $table.Cell($r + 1, $c + 1).Range.Font.Bold = -1 }
        }
    }
    try { $table.AutoFitBehavior(2) | Out-Null } catch {}
    $selection.SetRange($table.Range.End, $table.Range.End)
    $selection.TypeParagraph()
}

function Export-MarkdownReportPdf($word, [string]$markdownPath, [string]$pdfPath) {
    if (-not $pdfPath) { return }
    $lines = Get-Content -LiteralPath $markdownPath -Encoding UTF8
    $reportDoc = $null
    try {
        $reportDoc = $word.Documents.Add()
        $reportDoc.PageSetup.Orientation = 1
        $reportDoc.PageSetup.TopMargin = CmToPt 1.2
        $reportDoc.PageSetup.BottomMargin = CmToPt 1.2
        $reportDoc.PageSetup.LeftMargin = CmToPt 1.2
        $reportDoc.PageSetup.RightMargin = CmToPt 1.2
        $selection = $word.Selection

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = [string]$lines[$i]
            if ($line.Trim().StartsWith('|')) {
                $tableRows = New-Object System.Collections.Generic.List[string]
                while ($i -lt $lines.Count -and ([string]$lines[$i]).Trim().StartsWith('|')) {
                    $tableRows.Add([string]$lines[$i]) | Out-Null
                    $i++
                }
                $i--
                Add-ReportTable $reportDoc $selection $tableRows.ToArray()
                continue
            }

            if ($line -match '^#\s+(.+)$') {
                Add-ReportParagraph $selection $Matches[1] 18 -1 1
            } elseif ($line -match '^##\s+(.+)$') {
                Add-ReportParagraph $selection $Matches[1] 14 -1 0
            } elseif ([string]::IsNullOrWhiteSpace($line)) {
                $selection.TypeParagraph()
            } else {
                Add-ReportParagraph $selection $line 10.5 0 0
            }
        }
        $reportDoc.ExportAsFixedFormat($pdfPath, 17)
    } finally {
        if ($reportDoc) {
            $reportDoc.Close($false) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($reportDoc) | Out-Null
        }
    }
}

$inputFull = Resolve-FullPath $InputDocx
if (-not (Test-Path -LiteralPath $inputFull)) { throw "Input DOCX not found: $inputFull" }
if ([IO.Path]::GetExtension($inputFull).ToLowerInvariant() -ne '.docx') { throw 'Input must be a .docx file.' }

$dir = [IO.Path]::GetDirectoryName($inputFull)
$base = [IO.Path]::GetFileNameWithoutExtension($inputFull)
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $BackupDir) { $BackupDir = $dir }
$backupFull = Join-Path (Resolve-FullPath $BackupDir) ($base + "_ynu_format_backup_$stamp.docx")
if (-not $OutputDocx) { $OutputDocx = Join-Path $dir ($base + '_ynu_format_adjusted.docx') }
$outputFull = Resolve-FullPath $OutputDocx
if (-not $ReportJson) { $ReportJson = Join-Path $dir ($base + '_ynu_format_audit.json') }
if (-not $ReportMd) { $ReportMd = Join-Path $dir ($base + '_ynu_format_audit.md') }
$jsonFull = Resolve-FullPath $ReportJson
$mdFull = Resolve-FullPath $ReportMd
if (-not $ReportPdf) { $ReportPdf = [IO.Path]::ChangeExtension($mdFull, '.pdf') }
$pdfFull = Resolve-FullPath $ReportPdf

$target = $inputFull
$backupForReport = $null
$outputForReport = $null
$contentPreserved = $true

if ($Mode -ne 'Audit') {
    Copy-Item -LiteralPath $inputFull -Destination $backupFull -Force
    Copy-Item -LiteralPath $inputFull -Destination $outputFull -Force
    $target = $outputFull
    $backupForReport = $backupFull
    $outputForReport = $outputFull
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $readOnly = ($Mode -eq 'Audit')
    $doc = $word.Documents.Open($target, $false, $readOnly)
    $beforeFingerprint = Main-TextFingerprint $doc

    if ($Mode -ne 'Audit') {
        Apply-YnuFormatting $doc
        $doc.Save()
        $afterFingerprint = Main-TextFingerprint $doc
        $contentPreserved = ($beforeFingerprint -eq $afterFingerprint)
    }

    $audit = Audit-YnuDocument $doc $HeadingSystem $target $backupForReport $outputForReport $contentPreserved
    $audit | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonFull -Encoding UTF8
    Write-MarkdownReport $audit $mdFull
    $doc.Close($false)
    $doc = $null
    if (-not $NoReportPdf) {
        Export-MarkdownReportPdf $word $mdFull $pdfFull
    }
}
finally {
    $word.Quit()
    if ($doc) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

"YNU_FORMAT_DONE mode=$Mode output=$outputForReport backup=$backupForReport json=$jsonFull md=$mdFull pdf=$(if ($NoReportPdf) { '' } else { $pdfFull })"

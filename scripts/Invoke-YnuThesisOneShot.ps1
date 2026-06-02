param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [ValidateSet('Audit', 'Fix', 'AuditAndFix')]
    [string]$Mode = 'AuditAndFix',

    [ValidateSet('Auto', 'First', 'Second', 'Third')]
    [string]$HeadingSystem = 'Auto',

    [string]$OutputDocx,
    [string]$ReportJson,
    [string]$ReportMd,
    [string]$ReportPdf,
    [string]$BeforeReportJson,
    [string]$BeforeReportMd,
    [string]$BeforeReportPdf,
    [switch]$NoReportPdf,
    [switch]$ExportPreviewImages,
    [string]$PreviewDir
)

$ErrorActionPreference = 'Stop'

function Resolve-FullPath([string]$Path) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function CmToPt([double]$cm) { $cm * 28.3464567 }

function Clean-Text($range) {
    if ($null -eq $range) { return '' }
    Clean-String $range.Text
}

function Clean-String([string]$text) {
    if ($null -eq $text) { return '' }
    (($text -replace "[`r`a`f]", '') -replace '\s+', ' ').Trim()
}

function Near([double]$a, [double]$b, [double]$tol = 1.2) {
    [Math]::Abs($a - $b) -le $tol
}

function Is-StaticTocLine([string]$text) {
    if (-not $text) { return $false }
    return ($text -match '\.{3,}|…{2,}|[·.]{3,}\s*[0-9IVXLC]+$|[·.]{3,}\s*[\u2160-\u216B]+$')
}

function Is-TocTitle([string]$text) {
    return ($text -match '^\u76ee\s*\u5f55$')
}

function Is-ReferencesTitle([string]$text) {
    return ($text -match '^\u53c2\u8003\u6587\u732e$')
}

function Is-AckTitle([string]$text) {
    return ($text -match '^\u81f4\s*\u8c22$')
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

function Get-HeadingSystem($doc, [string]$preference) {
    if ($preference -ne 'Auto') { return $preference }
    $cn = '[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+'
    $second = 0
    $third = 0
    $first = 0

    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $text = Clean-Text $doc.Paragraphs.Item($i).Range
        if (-not $text -or (Is-StaticTocLine $text) -or (Is-TocTitle $text)) { continue }
        if ($text -match "^第${cn}章") { $second++ }
        elseif ($text -match '^\d+(\s|　|□)+\S') { $third++ }
        elseif ($text -match "^$cn、") { $first++ }
    }

    $kinds = @()
    if ($second -gt 0) { $kinds += 'Second' }
    if ($third -gt 0) { $kinds += 'Third' }
    if ($first -gt 0 -and $second -eq 0) { $kinds += 'First' }
    if ($kinds.Count -eq 1) { return $kinds[0] }
    if ($kinds.Count -gt 1) { return 'Mixed' }
    return 'Auto'
}

function Get-HeadingLevel([string]$text, [string]$system) {
    if (-not $text -or (Is-StaticTocLine $text) -or (Is-TocTitle $text)) { return 0 }

    $cn = '[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+'
    if ($text -match '^(\u6458\s*\u8981|Abstract|\u53c2\u8003\u6587\u732e|\u81f4\s*\u8c22|\u9644\u5f55\b|\u524d\u8a00|\u7eea\u8bba|\u5f15\u8a00|\u7ed3\u8bba$)') { return 1 }

    if ($system -eq 'Second') {
        if ($text -match "^第${cn}章") { return 1 }
        if ($text -match "^第${cn}节") { return 2 }
        if ($text -match "^$cn、") { return 3 }
        if ($text -match "^（$cn）") { return 3 }
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
        if ($text -match "^$cn、") { return 1 }
        if ($text -match "^（$cn）") { return 2 }
        if ($text -match '^\d+\.') { return 3 }
        if ($text -match '^（\d+）') { return 3 }
        return 0
    }

    if ($text -match "^第${cn}章") { return 1 }
    if ($text -match "^第${cn}节") { return 2 }
    if ($text -match '^\d+\.\d+\.\d+') { return 3 }
    if ($text -match '^\d+\.\d+(\s|　|□)+\S') { return 2 }
    if ($text -match '^\d+(\s|　|□)+\S') { return 1 }
    if ($text -match "^$cn、") { return 1 }
    return 0
}

function Apply-HeadingStyle($doc, $p, [int]$level) {
    try {
        if ($level -eq 1) { $p.Range.Style = $doc.Styles.Item('标题 1') }
        elseif ($level -eq 2) { $p.Range.Style = $doc.Styles.Item('标题 2') }
        else { $p.Range.Style = $doc.Styles.Item('标题 3') }
    } catch {
        try {
            if ($level -eq 1) { $p.Range.Style = $doc.Styles.Item(-2) }
            elseif ($level -eq 2) { $p.Range.Style = $doc.Styles.Item(-3) }
            else { $p.Range.Style = $doc.Styles.Item(-4) }
        } catch {}
    }
    try { $p.OutlineLevel = $level } catch {}

    if ($level -eq 1) {
        Set-ParaFormat $p '黑体' 'Times New Roman' 16 -1 1 9.6 6 $true 0
        $p.Format.PageBreakBefore = -1
    } elseif ($level -eq 2) {
        Set-ParaFormat $p '黑体' 'Times New Roman' 14 -1 0 6 0 $true 0
    } else {
        Set-ParaFormat $p '黑体' 'Times New Roman' 12 -1 0 0 0 $true 24
    }
}

function Find-TocTitleIndex($doc) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        if (Is-TocTitle (Clean-Text $doc.Paragraphs.Item($i).Range)) { return $i }
    }
    return $null
}

function Find-BodyStartAfterToc($doc, [int]$tocIndex, [string]$system) {
    $tocRanges = @()
    $tocEnd = $null
    try {
        for ($t = 1; $t -le $doc.TablesOfContents.Count; $t++) {
            $tocRange = $doc.TablesOfContents.Item($t).Range
            $tocRanges += $tocRange
            if ($null -eq $tocEnd -or $tocRange.End -gt $tocEnd) { $tocEnd = $tocRange.End }
        }
    } catch {}

    for ($i = $tocIndex + 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if ($null -ne $tocEnd -and $p.Range.Start -lt $tocEnd) { continue }
        $inToc = $false
        foreach ($r in $tocRanges) {
            if ($p.Range.Start -ge $r.Start -and $p.Range.End -le $r.End) { $inToc = $true; break }
        }
        if ($inToc) { continue }
        $style = ''
        try { $style = $p.Range.Style.NameLocal } catch {}
        if ($style -match 'TOC\s*\d|目录\s*\d') { continue }
        $text = Clean-Text $p.Range
        if (-not $text -or (Is-StaticTocLine $text)) { continue }
        $level = Get-HeadingLevel $text $system
        if ($level -eq 1 -and $text -notmatch '^(\u6458\s*\u8981|Abstract)$') { return $i }
    }
    return $null
}

function Remove-ExistingToc($doc, [string]$system) {
    while ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Delete()
    }

    $tocIndex = Find-TocTitleIndex $doc
    if ($null -eq $tocIndex) { return }
    $bodyIndex = Find-BodyStartAfterToc $doc $tocIndex $system
    if ($null -eq $bodyIndex) { return }

    $start = $doc.Paragraphs.Item($tocIndex).Range.End
    $end = $doc.Paragraphs.Item($bodyIndex).Range.Start
    if ($end -gt $start) {
        $doc.Range($start, $end).Delete() | Out-Null
    }
}

function Add-AutomaticToc($doc) {
    $tocIndex = Find-TocTitleIndex $doc
    if ($null -eq $tocIndex) { return }

    while ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Delete()
    }

    $tocTitle = $doc.Paragraphs.Item($tocIndex)
    Set-ParaFormat $tocTitle '黑体' 'Times New Roman' 16 -1 1 0 0 $true 0
    $tocTitle.Format.PageBreakBefore = -1

    $insertStart = $tocTitle.Range.End
    $spacer = $doc.Range($insertStart, $insertStart)
    $spacer.InsertBefore(([string][char]13) + ([string][char]13)) | Out-Null
    $tocRange = $doc.Range($insertStart, $insertStart + 1)
    $toc = $doc.TablesOfContents.Add($tocRange, $true, 1, 3, $false, '', $true, $true)
    $toc.RightAlignPageNumbers = $true
    $toc.IncludePageNumbers = $true
    $toc.UseHyperlinks = $true
    $toc.Update()

    for ($j = 1; $j -le $toc.Range.Paragraphs.Count; $j++) {
        $tp = $toc.Range.Paragraphs.Item($j)
        $style = ''
        try { $style = $tp.Range.Style.NameLocal } catch {}
        if ($style -match 'TOC\s*1|目录\s*1') {
            Set-ParaFormat $tp '黑体' 'Times New Roman' 12 0 0 0 0 $true 0
        } elseif ($style -match 'TOC\s*2|目录\s*2') {
            Set-ParaFormat $tp '宋体' 'Times New Roman' 12 0 0 0 0 $true 24
            $tp.Format.LeftIndent = 24
        } elseif ($style -match 'TOC\s*3|目录\s*3') {
            Set-ParaFormat $tp '宋体' 'Times New Roman' 12 0 0 0 0 $true 48
            $tp.Format.LeftIndent = 48
        }
    }
}

function Find-ParagraphIndexByPattern($doc, [string]$pattern) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $text = Clean-Text $doc.Paragraphs.Item($i).Range
        if ($text -match $pattern) { return $i }
    }
    return $null
}

function Ensure-SectionBreakBeforeParagraph($doc, [int]$idx) {
    if ($idx -le 1) { return }
    $p = $doc.Paragraphs.Item($idx)
    $prev = $doc.Paragraphs.Item($idx - 1)
    $curSection = $p.Range.Information(2)
    $prevSection = $prev.Range.Information(2)
    if ($curSection -eq $prevSection) {
        $r = $doc.Range($p.Range.Start, $p.Range.Start)
        $r.InsertBreak(2) | Out-Null
    }
    try { $doc.Paragraphs.Item($idx).Format.PageBreakBefore = 0 } catch {}
}

function Clear-HeaderFooter($hf) {
    try { $hf.LinkToPrevious = $false } catch {}
    try {
        for ($i = $hf.Shapes.Count; $i -ge 1; $i--) {
            $hf.Shapes.Item($i).Delete()
        }
    } catch {}
    try { $hf.Range.Delete() | Out-Null } catch {}
    try {
        $hf.Range.Font.NameFarEast = '宋体'
        $hf.Range.Font.NameAscii = 'Times New Roman'
        $hf.Range.Font.NameOther = 'Times New Roman'
        $hf.Range.Font.Size = 10.5
        $hf.Range.Font.Bold = 0
        $hf.Range.ParagraphFormat.Alignment = 1
        $hf.Range.ParagraphFormat.FirstLineIndent = 0
        $hf.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        $hf.Range.ParagraphFormat.LeftIndent = 0
        $hf.Range.ParagraphFormat.RightIndent = 0
    } catch {}
}

function Format-HeaderFooterRange($hf) {
    try {
        $hf.Range.Font.NameFarEast = '宋体'
        $hf.Range.Font.NameAscii = 'Times New Roman'
        $hf.Range.Font.NameOther = 'Times New Roman'
        $hf.Range.Font.Size = 10.5
        $hf.Range.Font.Bold = 0
        $hf.Range.ParagraphFormat.Alignment = 1
        $hf.Range.ParagraphFormat.FirstLineIndent = 0
        $hf.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        $hf.Range.ParagraphFormat.LeftIndent = 0
        $hf.Range.ParagraphFormat.RightIndent = 0
        $hf.Range.ParagraphFormat.SpaceBefore = 0
        $hf.Range.ParagraphFormat.SpaceAfter = 0
    } catch {}
}

function Reset-SectionHeadersFooters($section) {
    $section.PageSetup.DifferentFirstPageHeaderFooter = 0
    foreach ($idx in 1, 2, 3) {
        try { Clear-HeaderFooter $section.Headers.Item($idx) } catch {}
        try { Clear-HeaderFooter $section.Footers.Item($idx) } catch {}
    }
}

function Set-SelectionFooterFont($selection) {
    $selection.Font.NameFarEast = '宋体'
    $selection.Font.NameAscii = 'Times New Roman'
    $selection.Font.NameOther = 'Times New Roman'
    $selection.Font.Size = 10.5
    $selection.Font.Bold = 0
    $selection.ParagraphFormat.Alignment = 1
    $selection.ParagraphFormat.FirstLineIndent = 0
    $selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.ParagraphFormat.RightIndent = 0
}

function Write-RomanFooter($doc, $section) {
    $footer = $section.Footers.Item(1)
    Clear-HeaderFooter $footer
    $footer.PageNumbers.RestartNumberingAtSection = $true
    $footer.PageNumbers.StartingNumber = 1
    $footer.PageNumbers.NumberStyle = 1
    $footer.Range.Select()
    $selection = $doc.Application.Selection
    Set-SelectionFooterFont $selection
    $selection.Fields.Add($selection.Range, -1, 'PAGE \* ROMAN', $true) | Out-Null
    $selection.Collapse(0) | Out-Null
    Format-HeaderFooterRange $footer
}

function Write-BodyFooter($doc, $section) {
    $footer = $section.Footers.Item(1)
    Clear-HeaderFooter $footer
    $footer.PageNumbers.RestartNumberingAtSection = $true
    $footer.PageNumbers.StartingNumber = 1
    $footer.PageNumbers.NumberStyle = 0
    $footer.Range.Select()
    $selection = $doc.Application.Selection
    Set-SelectionFooterFont $selection
    $selection.TypeText('第 ')
    $selection.Fields.Add($selection.Range, -1, 'PAGE', $true) | Out-Null
    $selection.Collapse(0) | Out-Null
    $selection.TypeText(' 页（共 ')
    $selection.Fields.Add($selection.Range, -1, 'SECTIONPAGES', $true) | Out-Null
    $selection.Collapse(0) | Out-Null
    $selection.TypeText(' 页）')
    Format-HeaderFooterRange $footer
}

function Write-BodyHeader($doc, $section) {
    $header = $section.Headers.Item(1)
    Clear-HeaderFooter $header
    $header.Range.Select()
    $selection = $doc.Application.Selection
    Set-SelectionFooterFont $selection
    $selection.Fields.Add($selection.Range, -1, 'STYLEREF "标题 1" \* MERGEFORMAT', $true) | Out-Null
    $selection.Collapse(0) | Out-Null
    Format-HeaderFooterRange $header
}

function Normalize-YnuSectionsAndPageNumbers($doc, [string]$system) {
    $absCnIdx = Find-ParagraphIndexByPattern $doc '^摘\s*要$'
    $absEnIdx = Find-ParagraphIndexByPattern $doc '^Abstract$'
    $tocIdx = Find-TocTitleIndex $doc
    $bodyIdx = $null
    if ($null -ne $tocIdx) { $bodyIdx = Find-BodyStartAfterToc $doc $tocIdx $system }

    $anchors = @($absCnIdx, $absEnIdx, $tocIdx, $bodyIdx) | Where-Object { $null -ne $_ } | Sort-Object -Descending
    foreach ($idx in $anchors) { Ensure-SectionBreakBeforeParagraph $doc $idx }
    $doc.Repaginate()

    $absCnIdx = Find-ParagraphIndexByPattern $doc '^摘\s*要$'
    $absEnIdx = Find-ParagraphIndexByPattern $doc '^Abstract$'
    $tocIdx = Find-TocTitleIndex $doc
    if ($null -ne $tocIdx) { $bodyIdx = Find-BodyStartAfterToc $doc $tocIdx $system }

    foreach ($idx in @($absCnIdx, $absEnIdx, $tocIdx, $bodyIdx)) {
        if ($null -ne $idx) {
            try { $doc.Paragraphs.Item($idx).Format.PageBreakBefore = 0 } catch {}
        }
    }
    if ($null -ne $bodyIdx) {
        for ($i = $bodyIdx - 1; $i -ge 1; $i--) {
            $p = $doc.Paragraphs.Item($i)
            if (Clean-Text $p.Range) { break }
            try {
                $p.Format.PageBreakBefore = 0
                $p.Format.FirstLineIndent = 0
                $p.Format.CharacterUnitFirstLineIndent = 0
            } catch {}
        }
    }

    $doc.PageSetup.OddAndEvenPagesHeaderFooter = 0
    foreach ($section in @($doc.Sections)) {
        $section.PageSetup.PageWidth = 595.3
        $section.PageSetup.PageHeight = 841.9
        $section.PageSetup.TopMargin = CmToPt 2.5
        $section.PageSetup.BottomMargin = CmToPt 2.0
        $section.PageSetup.LeftMargin = CmToPt 3.0
        $section.PageSetup.RightMargin = CmToPt 2.0
        Reset-SectionHeadersFooters $section
    }

    $frontSectionNumbers = New-Object System.Collections.Generic.HashSet[int]
    foreach ($idx in @($absCnIdx, $absEnIdx, $tocIdx)) {
        if ($null -ne $idx) { $frontSectionNumbers.Add([int]$doc.Paragraphs.Item($idx).Range.Information(2)) | Out-Null }
    }
    foreach ($sectionNumber in $frontSectionNumbers) {
        Write-RomanFooter $doc $doc.Sections.Item($sectionNumber)
    }
    if ($null -ne $bodyIdx) {
        $bodySection = $doc.Sections.Item([int]$doc.Paragraphs.Item($bodyIdx).Range.Information(2))
        Write-BodyHeader $doc $bodySection
        Write-BodyFooter $doc $bodySection
    }

    $doc.Fields.Update() | Out-Null
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $doc.TablesOfContents.Item($i).Update() | Out-Null
    }
}

function Get-FieldErrorSignatures($doc) {
    try { $doc.Fields.Update() | Out-Null } catch {}
    $errors = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $doc.Fields.Count; $i++) {
        $code = Get-FieldCodeText $doc.Fields.Item($i)
        $result = Get-FieldResultText $doc.Fields.Item($i)
        if (Is-FieldErrorText $result) {
            $errors.Add("$code => $result") | Out-Null
        }
    }
    return $errors
}

function Get-FieldCodeText($field) {
    try { return Clean-String $field.Code.Text } catch { return '' }
}

function Get-FieldResultText($field) {
    try { return Clean-String $field.Result.Text } catch { return '' }
}

function Is-FieldErrorText([string]$text) {
    return ($text -match '错误|Error|引用源|Reference source|不能识别')
}

function Get-NormalizedFieldCode([string]$code) {
    $normalized = Clean-String $code
    $normalized = $normalized -replace '\s+\\\*\s+MERGEFORMAT\b', ''
    $normalized = $normalized -replace '\s+\\\*\s+CHARFORMAT\b', ''
    return (($normalized -replace '\s+', ' ').Trim())
}

function Is-CrossReferenceFieldCode([string]$code) {
    return ((Get-NormalizedFieldCode $code) -match '^(REF|PAGEREF|SEQ)\b')
}

function Normalize-FieldResultsForGuard($paragraph) {
    $normalized = ''
    try { $normalized = Clean-String $paragraph.Range.Text } catch { return '' }
    $fieldItems = @()

    try {
        for ($i = 1; $i -le $paragraph.Range.Fields.Count; $i++) {
            $field = $paragraph.Range.Fields.Item($i)
            $code = Get-NormalizedFieldCode (Get-FieldCodeText $field)
            if (-not (Is-CrossReferenceFieldCode $code)) { continue }
            $result = Get-FieldResultText $field
            if (-not $result) { continue }
            $fieldItems += [pscustomobject]@{
                Start = [int]$field.Result.Start
                Result = $result
                Token = "<WORD-FIELD:$code>"
            }
        }
    } catch {}

    $cursor = 0
    foreach ($item in @($fieldItems | Sort-Object Start)) {
        if ($cursor -lt 0 -or $cursor -gt $normalized.Length) { $cursor = 0 }
        $idx = $normalized.IndexOf($item.Result, $cursor, [System.StringComparison]::Ordinal)
        if ($idx -lt 0) {
            $idx = $normalized.IndexOf($item.Result, [System.StringComparison]::Ordinal)
        }
        if ($idx -lt 0) { continue }
        $normalized = $normalized.Substring(0, $idx) + $item.Token + $normalized.Substring($idx + $item.Result.Length)
        $cursor = $idx + $item.Token.Length
    }

    return $normalized
}

function Repair-SafeCrossReferenceFields($doc) {
    try { $doc.Fields.Update() | Out-Null } catch {}
    $repairs = New-Object System.Collections.Generic.List[string]

    for ($i = 1; $i -le $doc.Fields.Count; $i++) {
        $field = $doc.Fields.Item($i)
        $code = Get-FieldCodeText $field
        $result = Get-FieldResultText $field
        if (-not (Is-FieldErrorText $result)) { continue }
        if ($code -notmatch '^REF\s+(\S+)\s+\\r\s+\\h\s+\\\*\s+MERGEFORMAT$') { continue }

        $bookmarkName = $Matches[1]
        $bookmarkExists = $false
        try { $bookmarkExists = $doc.Bookmarks.Exists($bookmarkName) } catch {}
        if (-not $bookmarkExists) { continue }

        $newCode = Get-NormalizedFieldCode $code
        if ($newCode -notmatch '^REF\s+\S+\s+\\r\s+\\h$') { continue }

        $oldCodeText = $field.Code.Text
        try {
            $field.Code.Text = ' ' + $newCode + ' '
            $field.Update() | Out-Null
            $newResult = Get-FieldResultText $field
            if (Is-FieldErrorText $newResult) {
                $field.Code.Text = $oldCodeText
                try { $field.Update() | Out-Null } catch {}
                continue
            }
            $repairs.Add("field#${i}: $code -> $newCode => $newResult") | Out-Null
        } catch {
            try {
                $field.Code.Text = $oldCodeText
                $field.Update() | Out-Null
            } catch {}
        }
    }

    return $repairs
}

function Invoke-StructuralPass([string]$docxPath, [string]$headingPreference) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $null
    try {
        $doc = $word.Documents.Open($docxPath, $false, $false)
        foreach ($section in @($doc.Sections)) {
            $section.PageSetup.PageWidth = 595.3
            $section.PageSetup.PageHeight = 841.9
            $section.PageSetup.TopMargin = CmToPt 2.5
            $section.PageSetup.BottomMargin = CmToPt 2.0
            $section.PageSetup.LeftMargin = CmToPt 3.0
            $section.PageSetup.RightMargin = CmToPt 2.0
        }

        $system = Get-HeadingSystem $doc $headingPreference
        $preExistingFieldErrors = Get-FieldErrorSignatures $doc
        Remove-ExistingToc $doc $system

        $inRefs = $false
        for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
            $p = $doc.Paragraphs.Item($i)
            $text = Clean-Text $p.Range
            if (-not $text) { continue }

            if (Is-TocTitle $text) {
                Set-ParaFormat $p '黑体' 'Times New Roman' 16 -1 1 0 0 $true 0
                continue
            }
            if (Is-ReferencesTitle $text) { $inRefs = $true }
            if (Is-AckTitle $text) { $inRefs = $false }

            $level = Get-HeadingLevel $text $system
            if ($level -gt 0) {
                Apply-HeadingStyle $doc $p $level
                continue
            }

            if ($inRefs) {
                Set-ParaFormat $p '宋体' 'Times New Roman' 12 0 3 0 0 $true 0
            } else {
                Set-ParaFormat $p '宋体' 'Times New Roman' 12 0 3 0 0 $true 24
            }
        }

        Add-AutomaticToc $doc
        Normalize-YnuSectionsAndPageNumbers $doc $system
        Repair-SafeCrossReferenceFields $doc | Out-Null

        $postFieldErrors = Get-FieldErrorSignatures $doc
        $known = @{}
        foreach ($err in $preExistingFieldErrors) { $known[$err] = $true }
        $newFieldErrors = @()
        foreach ($err in $postFieldErrors) {
            if (-not $known.ContainsKey($err)) { $newFieldErrors += $err }
        }
        if ($newFieldErrors.Count -gt 0) {
            throw "New Word field error after one-shot pass: $($newFieldErrors -join '; ')"
        }

        $doc.Save()
    } finally {
        if ($doc) {
            $doc.Close($true) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

function Get-GuardFingerprint([string]$docxPath, [string]$headingPreference) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $null
    try {
        $doc = $word.Documents.Open($docxPath, $false, $true)
        try { $doc.Fields.Update() | Out-Null } catch {}
        $system = Get-HeadingSystem $doc $headingPreference
        $tocRanges = @()
        for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
            $tocRanges += $doc.TablesOfContents.Item($i).Range
        }
        $tocTitleIdx = Find-TocTitleIndex $doc
        $bodyStartIdx = $null
        if ($null -ne $tocTitleIdx) { $bodyStartIdx = Find-BodyStartAfterToc $doc $tocTitleIdx $system }

        $parts = New-Object System.Collections.Generic.List[string]
        for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
            $p = $doc.Paragraphs.Item($i)
            if ($tocRanges.Count -gt 0) {
                $inToc = $false
                foreach ($r in $tocRanges) {
                    if ($p.Range.Start -ge $r.Start -and $p.Range.End -le $r.End) { $inToc = $true; break }
                }
                if ($inToc) { continue }
            }
            if ($null -ne $tocTitleIdx -and $null -ne $bodyStartIdx -and $i -ge $tocTitleIdx -and $i -lt $bodyStartIdx) { continue }
            $text = Normalize-FieldResultsForGuard $p
            if (-not $text) { continue }
            if (Is-StaticTocLine $text) { continue }
            $parts.Add($text) | Out-Null
        }
        return ($parts -join "`n")
    } finally {
        if ($doc) {
            $doc.Close($false) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

function Assert-ContentPreserved([string]$beforeDocx, [string]$afterDocx, [string]$headingPreference) {
    $before = Get-GuardFingerprint $beforeDocx $headingPreference
    $after = Get-GuardFingerprint $afterDocx $headingPreference
    if ($before -ne $after) {
        throw "Content preservation guard failed: non-TOC thesis text changed. The adjusted DOCX was not accepted as format-only output."
    }
}

function Export-DocxPreviewImages([string]$docxPath, [string]$outputDir, [string]$label) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
    $pdfPath = Join-Path $outputDir ($label + '.pdf')

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $null
    try {
        $doc = $word.Documents.Open($docxPath, $false, $true)
        $doc.ExportAsFixedFormat($pdfPath, 17)
    } finally {
        if ($doc) {
            $doc.Close($false) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    $pdftoppm = Get-Command pdftoppm -ErrorAction SilentlyContinue
    if ($pdftoppm) {
        $prefix = Join-Path $outputDir $label
        & $pdftoppm.Source -png -r 160 $pdfPath $prefix | Out-Null
    }
    return $pdfPath
}

function Invoke-CoreFormatScript([string[]]$arguments) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $core @arguments | Out-Host
    return $LASTEXITCODE
}

$inputFull = Resolve-FullPath $InputDocx
$dir = [IO.Path]::GetDirectoryName($inputFull)
$base = [IO.Path]::GetFileNameWithoutExtension($inputFull)
if (-not $OutputDocx) { $OutputDocx = Join-Path $dir ($base + '_ynu_format_adjusted.docx') }
$outputFull = Resolve-FullPath $OutputDocx
if (-not $ReportJson) { $ReportJson = Join-Path $dir ($base + '_ynu_format_audit.json') }
if (-not $ReportMd) { $ReportMd = Join-Path $dir ($base + '_ynu_format_audit.md') }
$jsonFull = Resolve-FullPath $ReportJson
$mdFull = Resolve-FullPath $ReportMd
if (-not $ReportPdf) { $ReportPdf = [IO.Path]::ChangeExtension($mdFull, '.pdf') }
$pdfFull = Resolve-FullPath $ReportPdf
if (-not $BeforeReportJson) { $BeforeReportJson = Join-Path $dir ($base + '_ynu_format_audit_before.json') }
if (-not $BeforeReportMd) { $BeforeReportMd = Join-Path $dir ($base + '_ynu_format_audit_before.md') }
$beforeJsonFull = Resolve-FullPath $BeforeReportJson
$beforeMdFull = Resolve-FullPath $BeforeReportMd
if (-not $BeforeReportPdf) { $BeforeReportPdf = [IO.Path]::ChangeExtension($beforeMdFull, '.pdf') }
$beforePdfFull = Resolve-FullPath $BeforeReportPdf
if (-not $PreviewDir) { $PreviewDir = Join-Path $dir ($base + '_ynu_format_previews') }
$previewFull = Resolve-FullPath $PreviewDir

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$core = Join-Path $scriptDir 'Invoke-YnuThesisFormat.ps1'

if ($Mode -eq 'Audit') {
    $auditArgs = @('-InputDocx', $inputFull, '-Mode', 'Audit', '-HeadingSystem', $HeadingSystem, '-ReportJson', $jsonFull, '-ReportMd', $mdFull)
    if ($NoReportPdf) { $auditArgs += '-NoReportPdf' } else { $auditArgs += @('-ReportPdf', $pdfFull) }
    exit (Invoke-CoreFormatScript $auditArgs)
}

$beforeArgs = @('-InputDocx', $inputFull, '-Mode', 'Audit', '-HeadingSystem', $HeadingSystem, '-ReportJson', $beforeJsonFull, '-ReportMd', $beforeMdFull)
if ($NoReportPdf) { $beforeArgs += '-NoReportPdf' } else { $beforeArgs += @('-ReportPdf', $beforePdfFull) }
$exitCode = Invoke-CoreFormatScript $beforeArgs
if ($exitCode -ne 0) { exit $exitCode }

if ($ExportPreviewImages) {
    Export-DocxPreviewImages $inputFull $previewFull 'before' | Out-Null
}

$intermediateJson = [IO.Path]::ChangeExtension($jsonFull, '.intermediate.json')
$intermediateMd = [IO.Path]::ChangeExtension($mdFull, '.intermediate.md')
$intermediatePdf = [IO.Path]::ChangeExtension($mdFull, '.intermediate.pdf')
$intermediateArgs = @('-InputDocx', $inputFull, '-Mode', 'AuditAndFix', '-HeadingSystem', $HeadingSystem, '-OutputDocx', $outputFull, '-ReportJson', $intermediateJson, '-ReportMd', $intermediateMd)
if ($NoReportPdf) { $intermediateArgs += '-NoReportPdf' } else { $intermediateArgs += @('-ReportPdf', $intermediatePdf) }
$exitCode = Invoke-CoreFormatScript $intermediateArgs
if ($exitCode -ne 0) { exit $exitCode }

Invoke-StructuralPass $outputFull $HeadingSystem
Assert-ContentPreserved $inputFull $outputFull $HeadingSystem

if ($ExportPreviewImages) {
    Export-DocxPreviewImages $outputFull $previewFull 'after' | Out-Null
}

$finalArgs = @('-InputDocx', $outputFull, '-Mode', 'Audit', '-HeadingSystem', $HeadingSystem, '-ReportJson', $jsonFull, '-ReportMd', $mdFull)
if ($NoReportPdf) { $finalArgs += '-NoReportPdf' } else { $finalArgs += @('-ReportPdf', $pdfFull) }
$exitCode = Invoke-CoreFormatScript $finalArgs
if ($exitCode -ne 0) { exit $exitCode }

"YNU_STRICT_DONE beforeJson=$beforeJsonFull beforeMd=$beforeMdFull beforePdf=$(if ($NoReportPdf) { '' } else { $beforePdfFull }) output=$outputFull json=$jsonFull md=$mdFull pdf=$(if ($NoReportPdf) { '' } else { $pdfFull }) intermediateJson=$intermediateJson intermediateMd=$intermediateMd intermediatePdf=$(if ($NoReportPdf) { '' } else { $intermediatePdf }) previews=$previewFull"


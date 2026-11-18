[CmdletBinding()]
param(
    [switch]$ValidateOnly,
    [switch]$SmokeTest
)

$ErrorActionPreference = "Stop"
$contentPath = Join-Path $PSScriptRoot "presenter-content.json"
if (-not (Test-Path -LiteralPath $contentPath)) {
    throw "Missing presenter-content.json next to this script."
}

$script:content = Get-Content -LiteralPath $contentPath -Raw -Encoding UTF8 | ConvertFrom-Json
$script:expectedPptxPath = [System.IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot $script:content.presentationFileName)
)
if (-not (Test-Path -LiteralPath $script:expectedPptxPath)) {
    throw "Missing the matching SVM PPTX next to the presenter files."
}
if ($script:content.slides.Count -ne $script:content.expectedSlideCount) {
    throw "Expected $($script:content.expectedSlideCount) slide scripts, found $($script:content.slides.Count)."
}
if ($script:content.expectedSlideCount -ne 13) {
    throw "The presenter is locked to a 13-slide presentation."
}
for ($i = 0; $i -lt $script:content.slides.Count; $i++) {
    if ($script:content.slides[$i].index -ne ($i + 1)) {
        throw "Slide indexes must be consecutive and start at 1."
    }
}

foreach ($slide in $script:content.slides) {
    foreach ($field in @("title", "goal", "talk", "process", "code", "codeExplanation", "ask", "guide", "transition", "pitfall")) {
        if ($null -eq $slide.$field) {
            throw "Slide $($slide.index) is missing '$field'."
        }
    }
}

function Normalize-SlideTitle {
    param([string]$Value)
    return (($Value -replace '\s+', ' ').Trim())
}

function Test-PresentationContract {
    param(
        [string]$FullName,
        [int]$SlideCount,
        [string[]]$SlideTitles
    )
    $actualPath = ""
    try {
        $actualPath = [System.IO.Path]::GetFullPath($FullName)
    }
    catch {
        return @{ Matched = $false; Reason = "path" }
    }
    if (-not [string]::Equals($actualPath, $script:expectedPptxPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return @{ Matched = $false; Reason = "path" }
    }
    if ($SlideCount -ne $script:content.expectedSlideCount) {
        return @{ Matched = $false; Reason = "count" }
    }
    if ($SlideTitles.Count -ne $script:content.expectedSlideCount) {
        return @{ Matched = $false; Reason = "titles" }
    }
    for ($i = 0; $i -lt $script:content.slides.Count; $i++) {
        $actualTitle = Normalize-SlideTitle -Value $SlideTitles[$i]
        $expectedTitle = Normalize-SlideTitle -Value $script:content.slides[$i].title
        if ($actualTitle -cne $expectedTitle) {
            return @{ Matched = $false; Reason = "titles" }
        }
    }
    return @{ Matched = $true; Reason = "" }
}

if ($SmokeTest) {
    $testTitles = @($script:content.slides | ForEach-Object { $_.title })
    $valid = Test-PresentationContract -FullName $script:expectedPptxPath -SlideCount 13 -SlideTitles $testTitles
    $wrongPath = Test-PresentationContract -FullName ($script:expectedPptxPath + ".other") -SlideCount 13 -SlideTitles $testTitles
    $wrongCount = Test-PresentationContract -FullName $script:expectedPptxPath -SlideCount 11 -SlideTitles $testTitles
    $wrongTitles = @($testTitles)
    $wrongTitles[7] = "MISMATCH"
    $wrongTitleResult = Test-PresentationContract -FullName $script:expectedPptxPath -SlideCount 13 -SlideTitles $wrongTitles
    if (-not $valid.Matched -or $wrongPath.Matched -or $wrongCount.Matched -or $wrongTitleResult.Matched) {
        throw "Presentation matching smoke test failed."
    }
    Write-Output "Contract smoke test passed: exact path, 13-page count, and all titles required; mismatches rejected."
}

if ($ValidateOnly) {
    $talkChars = ($script:content.slides | ForEach-Object { $_.talk.Length } | Measure-Object -Sum).Sum
    Write-Output ("Content OK: {0} slides, {1} FAQ answers, {2} spoken-script characters. Target: {3}" -f $script:content.slides.Count, $script:content.questions.Count, $talkChars, $script:expectedPptxPath)
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:currentSlideIndex = 1
$script:pptApp = $null
$script:slideShow = $null
$script:matchingPresentation = $null
$script:matchFailure = $null
$script:notesSourcePresentation = $null
$script:updatingSlide = $false
$script:readerControls = New-Object System.Collections.ArrayList
$script:baseFontSize = 13.0

$colors = @{
    Ink = [System.Drawing.Color]::FromArgb(32, 42, 48)
    Muted = [System.Drawing.Color]::FromArgb(92, 103, 109)
    Paper = [System.Drawing.Color]::FromArgb(248, 249, 247)
    White = [System.Drawing.Color]::White
    Header = [System.Drawing.Color]::FromArgb(30, 42, 47)
    Teal = [System.Drawing.Color]::FromArgb(0, 126, 128)
    Coral = [System.Drawing.Color]::FromArgb(222, 94, 75)
    Rule = [System.Drawing.Color]::FromArgb(218, 224, 222)
    PaleTeal = [System.Drawing.Color]::FromArgb(229, 241, 239)
}

$fontName = "Microsoft YaHei UI"
try {
    $null = New-Object System.Drawing.Font($fontName, 12)
}
catch {
    $fontName = "Microsoft YaHei"
}

function New-Label {
    param(
        [string]$Text,
        [int]$Size = 10,
        [bool]$Bold = $false,
        [System.Drawing.Color]$Color = $colors.Ink
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.ForeColor = $Color
    $label.Font = New-Object System.Drawing.Font($fontName, $Size, $(if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }))
    return $label
}

function New-Reader {
    param(
        [string]$Text,
        [bool]$IsCode = $false
    )
    $reader = New-Object System.Windows.Forms.RichTextBox
    $reader.ReadOnly = $true
    $reader.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $reader.BackColor = $colors.Paper
    $reader.ForeColor = $colors.Ink
    $reader.DetectUrls = $false
    $reader.HideSelection = $false
    $reader.WordWrap = $true
    $reader.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    if ($IsCode) {
        $reader.Font = New-Object System.Drawing.Font("Consolas", 10.5)
    }
    else {
        $reader.Font = New-Object System.Drawing.Font($fontName, $script:baseFontSize)
        $null = $script:readerControls.Add($reader)
    }
    $reader.Text = $Text
    return $reader
}

function Add-Section {
    param(
        [System.Windows.Forms.Panel]$Panel,
        [string]$Title,
        [string]$Text,
        [int]$Top,
        [int]$Height,
        [bool]$IsCode = $false,
        [int]$Width = 520
    )
    $heading = New-Label -Text $Title -Size 10 -Bold $true -Color $colors.Teal
    $heading.Location = New-Object System.Drawing.Point(14, $Top)
    $Panel.Controls.Add($heading)

    $body = New-Reader -Text $Text -IsCode $IsCode
    $body.Location = New-Object System.Drawing.Point(14, ($Top + 24))
    $body.Size = New-Object System.Drawing.Size($Width, $Height)
    $Panel.Controls.Add($body)
    return ($Top + 24 + $Height + 20)
}

function New-TabPage {
    param([string]$Title)
    $page = New-Object System.Windows.Forms.TabPage
    $page.Text = $Title
    $page.BackColor = $colors.Paper
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.AutoScroll = $true
    $panel.BackColor = $colors.Paper
    $page.Controls.Add($panel)
    return @{ Page = $page; Panel = $panel }
}

function Get-SlideTitleMatch {
    param(
        [object]$Presentation,
        [int]$Index
    )
    $expectedTitle = Normalize-SlideTitle $script:content.slides[$Index - 1].title
    $slide = $Presentation.Slides.Item($Index)
    foreach ($shape in $slide.Shapes) {
        try {
            if ([int]$shape.HasTextFrame -ne 0 -and [int]$shape.TextFrame.HasText -ne 0) {
                $text = Normalize-SlideTitle ([string]$shape.TextFrame.TextRange.Text)
                if ($text -ceq $expectedTitle) {
                    return $expectedTitle
                }
            }
        }
        catch {
        }
    }
    return "<标题不匹配>"
}

function Get-SpeakerNotes {
    param(
        [object]$Presentation,
        [int]$Index
    )
    $parts = New-Object System.Collections.Generic.List[string]
    $notesPage = $Presentation.Slides.Item($Index).NotesPage
    foreach ($shape in $notesPage.Shapes) {
        try {
            if ([int]$shape.HasTextFrame -eq 0 -or [int]$shape.TextFrame.HasText -eq 0) {
                continue
            }
            $isPlaceholder = ([int]$shape.Type -eq 14)
            if ($isPlaceholder -and [int]$shape.PlaceholderFormat.Type -ne 2) {
                continue
            }
            $noteText = ([string]$shape.TextFrame.TextRange.Text).Trim()
            if ($noteText.Length -gt 0 -and -not $parts.Contains($noteText)) {
                $parts.Add($noteText)
            }
        }
        catch {
        }
    }
    return ($parts -join "`r`n`r`n")
}

function Get-MatchingSlideShow {
    $script:slideShow = $null
    $script:matchingPresentation = $null
    $script:matchFailure = $null
    try {
        if ($null -eq $script:pptApp) {
            $script:pptApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
        }
        $windows = $script:pptApp.SlideShowWindows
        if ($windows.Count -eq 0) {
            return $null
        }
        $script:matchFailure = "path"
        for ($windowIndex = 1; $windowIndex -le $windows.Count; $windowIndex++) {
            $candidate = $windows.Item($windowIndex)
            $presentation = $candidate.Presentation
            $contractPath = [string]$presentation.FullName
            try {
                $contractPath = [System.IO.Path]::GetFullPath($contractPath)
            }
            catch {
                continue
            }
            if (-not [string]::Equals($contractPath, $script:expectedPptxPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
            if ([int]$presentation.Slides.Count -ne $script:content.expectedSlideCount) {
                $script:matchFailure = "count"
                continue
            }
            $titles = New-Object System.Collections.Generic.List[string]
            for ($slideIndex = 1; $slideIndex -le $script:content.expectedSlideCount; $slideIndex++) {
                $titles.Add((Get-SlideTitleMatch -Presentation $presentation -Index $slideIndex))
            }
            $contract = Test-PresentationContract `
                -FullName $contractPath `
                -SlideCount ([int]$presentation.Slides.Count) `
                -SlideTitles @($titles.ToArray())
            if (-not $contract.Matched) {
                $script:matchFailure = $contract.Reason
                continue
            }
            $script:slideShow = $candidate
            $script:matchingPresentation = $presentation
            $script:matchFailure = $null
            return $candidate
        }
    }
    catch {
        $script:pptApp = $null
        $script:slideShow = $null
        $script:matchingPresentation = $null
        $script:matchFailure = "error"
    }
    return $null
}

function Refresh-SpeakerNotes {
    param(
        [int]$Index,
        [bool]$Force = $false
    )
    if ($null -eq $script:matchingPresentation) {
        $script:notesReader.Text = $script:content.ui.notesUnavailable
        $script:notesSourcePresentation = $null
        return
    }
    if (-not $Force -and [object]::ReferenceEquals($script:notesSourcePresentation, $script:matchingPresentation)) {
        return
    }
    try {
        $notes = Get-SpeakerNotes -Presentation $script:matchingPresentation -Index $Index
        if ([string]::IsNullOrWhiteSpace($notes)) {
            $notes = $script:content.ui.notesUnavailable
        }
        $script:notesReader.Text = $notes
        $script:notesSourcePresentation = $script:matchingPresentation
    }
    catch {
        $script:notesReader.Text = $script:content.ui.statusError
        $script:notesSourcePresentation = $null
    }
}

function Set-SlideContent {
    param(
        [int]$Index,
        [bool]$NavigatePowerPoint = $false
    )
    if ($Index -lt 1) { $Index = 1 }
    if ($Index -gt $script:content.slides.Count) { $Index = $script:content.slides.Count }

    if ($NavigatePowerPoint) {
        $show = Get-MatchingSlideShow
        if ($null -ne $show) {
            try {
                $show.View.GotoSlide($Index)
            }
            catch {
                $script:statusLabel.Text = $script:content.ui.statusError
            }
        }
    }

    $script:currentSlideIndex = $Index
    $slide = $script:content.slides[$Index - 1]
    $script:updatingSlide = $true
    $script:slidePicker.SelectedIndex = $Index - 1
    $script:updatingSlide = $false
    $script:slideNumberLabel.Text = ("{0:00} / {1:00}" -f $Index, $script:content.slides.Count)
    $script:talkReader.Text = $slide.talk
    Refresh-SpeakerNotes -Index $Index -Force $true
    $script:readerControls.Clear()
    $null = $script:readerControls.Add($script:talkReader)

    $script:processPanel.Controls.Clear()
    $top = 12
    $top = Add-Section -Panel $script:processPanel -Title $script:content.ui.sectionGoal -Text $slide.goal -Top $top -Height 74
    $processText = @()
    for ($i = 0; $i -lt $slide.process.Count; $i++) {
        $processText += ("{0}. {1}" -f ($i + 1), $slide.process[$i])
    }
    $top = Add-Section -Panel $script:processPanel -Title $script:content.ui.sectionProcess -Text ($processText -join "`r`n`r`n") -Top $top -Height 250
    $top = Add-Section -Panel $script:processPanel -Title $script:content.ui.sectionKey -Text $slide.key -Top $top -Height 76
    $null = Add-Section -Panel $script:processPanel -Title $script:content.ui.sectionPitfall -Text $slide.pitfall -Top $top -Height 92

    $script:codePanel.Controls.Clear()
    $top = 12
    $top = Add-Section -Panel $script:codePanel -Title $script:content.ui.sectionCode -Text $slide.code -Top $top -Height 245 -IsCode $true
    $null = Add-Section -Panel $script:codePanel -Title $script:content.ui.sectionCodeExplain -Text $slide.codeExplanation -Top $top -Height 260

    $script:audiencePanel.Controls.Clear()
    $top = 12
    $top = Add-Section -Panel $script:audiencePanel -Title $script:content.ui.sectionQuestion -Text $slide.ask -Top $top -Height 120
    $top = Add-Section -Panel $script:audiencePanel -Title $script:content.ui.sectionGuide -Text $slide.guide -Top $top -Height 130
    $top = Add-Section -Panel $script:audiencePanel -Title $script:content.ui.sectionTransition -Text $slide.transition -Top $top -Height 100
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:content.ui.appTitle
$form.Size = New-Object System.Drawing.Size(650, 840)
$form.MinimumSize = New-Object System.Drawing.Size(540, 680)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::SizableToolWindow
$form.ShowInTaskbar = $true
$form.TopMost = $true
$form.BackColor = $colors.Paper
$form.ForeColor = $colors.Ink
$form.Font = New-Object System.Drawing.Font($fontName, 9.5)
$form.KeyPreview = $true

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = [System.Windows.Forms.DockStyle]::Fill
$root.ColumnCount = 1
$root.RowCount = 5
$root.Margin = New-Object System.Windows.Forms.Padding(0)
$root.Padding = New-Object System.Windows.Forms.Padding(0)
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 88)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 66)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
$form.Controls.Add($root)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = [System.Windows.Forms.DockStyle]::Fill
$header.BackColor = $colors.Header
$titleLabel = New-Label -Text $script:content.title -Size 15 -Bold $true -Color $colors.White
$titleLabel.Location = New-Object System.Drawing.Point(16, 10)
$header.Controls.Add($titleLabel)
$hintLabel = New-Label -Text $script:content.ui.appHint -Size 8 -Color ([System.Drawing.Color]::FromArgb(202, 214, 213))
$hintLabel.Location = New-Object System.Drawing.Point(18, 36)
$header.Controls.Add($hintLabel)
$script:statusLabel = New-Label -Text $script:content.ui.statusConnecting -Size 8 -Bold $true -Color ([System.Drawing.Color]::FromArgb(166, 218, 197))
$script:statusLabel.AutoSize = $false
$script:statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:statusLabel.Location = New-Object System.Drawing.Point(16, 58)
$script:statusLabel.Size = New-Object System.Drawing.Size(610, 26)
$header.Controls.Add($script:statusLabel)
$root.Controls.Add($header, 0, 0)

$settings = New-Object System.Windows.Forms.Panel
$settings.Dock = [System.Windows.Forms.DockStyle]::Fill
$settings.BackColor = $colors.Paper
$displayLabel = New-Label -Text $script:content.ui.displayLabel -Size 9 -Bold $true -Color $colors.Muted
$displayLabel.Location = New-Object System.Drawing.Point(14, 8)
$settings.Controls.Add($displayLabel)
$script:displayPicker = New-Object System.Windows.Forms.ComboBox
$script:displayPicker.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:displayPicker.Location = New-Object System.Drawing.Point(116, 5)
$script:displayPicker.Width = 260
$script:displayPicker.Height = 28
$settings.Controls.Add($script:displayPicker)
$script:topMostCheck = New-Object System.Windows.Forms.CheckBox
$script:topMostCheck.Text = $script:content.ui.alwaysOnTop
$script:topMostCheck.Checked = $true
$script:topMostCheck.AutoSize = $true
$script:topMostCheck.Location = New-Object System.Drawing.Point(390, 7)
$settings.Controls.Add($script:topMostCheck)
$displayNotice = if ([System.Windows.Forms.Screen]::AllScreens.Count -gt 1) { $script:content.ui.displayNoticeExtended } else { $script:content.ui.displayNoticeMirror }
$noticeColor = if ([System.Windows.Forms.Screen]::AllScreens.Count -gt 1) { $colors.Teal } else { $colors.Coral }
$noticeLabel = New-Label -Text $displayNotice -Size 8 -Color $noticeColor
$noticeLabel.Location = New-Object System.Drawing.Point(14, 36)
$settings.Controls.Add($noticeLabel)
$root.Controls.Add($settings, 0, 1)

$navigation = New-Object System.Windows.Forms.FlowLayoutPanel
$navigation.Dock = [System.Windows.Forms.DockStyle]::Fill
$navigation.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$navigation.WrapContents = $false
$navigation.Padding = New-Object System.Windows.Forms.Padding(14, 7, 8, 4)
$navigation.BackColor = $colors.Paper
$script:previousButton = New-Object System.Windows.Forms.Button
$script:previousButton.Text = $script:content.ui.previous
$script:previousButton.Width = 90
$script:previousButton.Height = 34
$script:previousButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$script:previousButton.FlatAppearance.BorderColor = $colors.Rule
$script:previousButton.BackColor = $colors.White
$script:previousButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$navigation.Controls.Add($script:previousButton)
$script:slidePicker = New-Object System.Windows.Forms.ComboBox
$script:slidePicker.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:slidePicker.Width = 315
$script:slidePicker.Height = 32
$script:slidePicker.Margin = New-Object System.Windows.Forms.Padding(0, 2, 8, 0)
foreach ($slide in $script:content.slides) {
    $script:slidePicker.Items.Add(("{0:00}  {1}" -f $slide.index, $slide.title)) | Out-Null
}
$navigation.Controls.Add($script:slidePicker)
$script:nextButton = New-Object System.Windows.Forms.Button
$script:nextButton.Text = $script:content.ui.next
$script:nextButton.Width = 90
$script:nextButton.Height = 34
$script:nextButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$script:nextButton.FlatAppearance.BorderColor = $colors.Rule
$script:nextButton.BackColor = $colors.PaleTeal
$script:nextButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$navigation.Controls.Add($script:nextButton)
$script:slideNumberLabel = New-Label -Text "01 / 13" -Size 10 -Bold $true -Color $colors.Coral
$script:slideNumberLabel.Margin = New-Object System.Windows.Forms.Padding(4, 9, 0, 0)
$navigation.Controls.Add($script:slideNumberLabel)
$root.Controls.Add($navigation, 0, 2)

$script:tabs = New-Object System.Windows.Forms.TabControl
$script:tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:tabs.Font = New-Object System.Drawing.Font($fontName, 9)
$script:tabs.Padding = New-Object System.Drawing.Point(9, 5)

$talkTab = New-TabPage -Title $script:content.ui.tabTalk
$talkPage = $talkTab.Page
$talkPanel = $talkTab.Panel
$script:talkTab = $talkPage
$talkHeading = New-Label -Text $script:content.ui.sectionTalk -Size 10 -Bold $true -Color $colors.Teal
$talkHeading.Location = New-Object System.Drawing.Point(16, 15)
$talkPanel.Controls.Add($talkHeading)
$script:talkReader = New-Reader -Text ""
$script:talkReader.Location = New-Object System.Drawing.Point(16, 43)
$script:talkReader.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:talkReader.Size = New-Object System.Drawing.Size(580, 660)
$talkPanel.Controls.Add($script:talkReader)
$script:tabs.TabPages.Add($talkPage)

$notesTab = New-TabPage -Title $script:content.ui.tabNotes
$notesPage = $notesTab.Page
$notesPanel = $notesTab.Panel
$notesHeading = New-Label -Text $script:content.ui.sectionPptNotes -Size 10 -Bold $true -Color $colors.Teal
$notesHeading.Location = New-Object System.Drawing.Point(16, 15)
$notesPanel.Controls.Add($notesHeading)
$script:notesReader = New-Reader -Text $script:content.ui.notesUnavailable
$script:notesReader.Location = New-Object System.Drawing.Point(16, 43)
$script:notesReader.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:notesReader.Size = New-Object System.Drawing.Size(580, 660)
$notesPanel.Controls.Add($script:notesReader)
$script:tabs.TabPages.Add($notesPage)

$processTab = New-TabPage -Title $script:content.ui.tabProcess
$script:processPanel = $processTab["Panel"]
$script:tabs.TabPages.Add($processTab["Page"])

$codeTab = New-TabPage -Title $script:content.ui.tabCode
$script:codePanel = $codeTab["Panel"]
$script:tabs.TabPages.Add($codeTab["Page"])

$audienceTab = New-TabPage -Title $script:content.ui.tabAudience
$script:audiencePanel = $audienceTab["Panel"]
$script:tabs.TabPages.Add($audienceTab["Page"])

$qaTab = New-TabPage -Title $script:content.ui.tabQA
$qaTop = 12
foreach ($item in $script:content.questions) {
    $qaTop = Add-Section -Panel $qaTab["Panel"] -Title $item.question -Text $item.answer -Top $qaTop -Height 86
}
$script:tabs.TabPages.Add($qaTab["Page"])

$root.Controls.Add($script:tabs, 0, 3)

$footer = New-Object System.Windows.Forms.Panel
$footer.Dock = [System.Windows.Forms.DockStyle]::Fill
$footer.BackColor = $colors.White
$footer.Padding = New-Object System.Windows.Forms.Padding(12, 4, 10, 3)
$footerLabel = New-Label -Text $script:content.ui.footer -Size 8 -Color $colors.Muted
$footerLabel.Location = New-Object System.Drawing.Point(12, 12)
$footer.Controls.Add($footerLabel)
$buttonFontDown = New-Object System.Windows.Forms.Button
$buttonFontDown.Text = $script:content.ui.fontSmaller
$buttonFontDown.Size = New-Object System.Drawing.Size(52, 28)
$buttonFontDown.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonFontDown.FlatAppearance.BorderColor = $colors.Rule
$buttonFontDown.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonFontDown.Location = New-Object System.Drawing.Point(430, 7)
$footer.Controls.Add($buttonFontDown)
$buttonFontUp = New-Object System.Windows.Forms.Button
$buttonFontUp.Text = $script:content.ui.fontLarger
$buttonFontUp.Size = New-Object System.Drawing.Size(52, 28)
$buttonFontUp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonFontUp.FlatAppearance.BorderColor = $colors.Rule
$buttonFontUp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonFontUp.Location = New-Object System.Drawing.Point(485, 7)
$footer.Controls.Add($buttonFontUp)
$minimizeButton = New-Object System.Windows.Forms.Button
$minimizeButton.Text = $script:content.ui.minimize
$minimizeButton.Size = New-Object System.Drawing.Size(84, 28)
$minimizeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$minimizeButton.FlatAppearance.BorderColor = $colors.Rule
$minimizeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$minimizeButton.Location = New-Object System.Drawing.Point(540, 7)
$footer.Controls.Add($minimizeButton)
$root.Controls.Add($footer, 0, 4)

$script:displayPicker.Items.Clear()
$script:screens = [System.Windows.Forms.Screen]::AllScreens
for ($i = 0; $i -lt $script:screens.Count; $i++) {
    $screen = $script:screens[$i]
    $bounds = $screen.Bounds
    $displayName = "{0} {1}  ({2} x {3})" -f $script:content.ui.displayOther, ($i + 1), $bounds.Width, $bounds.Height
    if ($screen.Primary) {
        $displayName += "  [" + $script:content.ui.displayPrimary + "]"
    }
    $script:displayPicker.Items.Add($displayName) | Out-Null
}
if ($script:screens.Count -eq 0) {
    $script:displayPicker.Items.Add($script:content.ui.displayNone) | Out-Null
    $script:displayPicker.Enabled = $false
}
else {
    $primaryIndex = 0
    for ($i = 0; $i -lt $script:screens.Count; $i++) {
        if ($script:screens[$i].Primary) { $primaryIndex = $i }
    }
    $script:displayPicker.SelectedIndex = $primaryIndex
}

if ($script:screens.Count -lt 2) {
    $script:statusLabel.Text = $script:content.ui.statusMirrorWarning
    $script:statusLabel.ForeColor = $colors.Coral
}
else {
    $script:statusLabel.Text = $script:content.ui.statusExtendedHint
    $script:statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(166, 218, 197)
}

function Move-WindowToSelectedDisplay {
    if ($script:screens.Count -eq 0 -or $script:displayPicker.SelectedIndex -lt 0) { return }
    $screen = $script:screens[$script:displayPicker.SelectedIndex]
    $area = $screen.WorkingArea
    $form.Location = New-Object System.Drawing.Point(($area.Right - $form.Width - 24), ($area.Top + 28))
}

$script:displayPicker.add_SelectedIndexChanged({
    Move-WindowToSelectedDisplay
})
$script:topMostCheck.add_CheckedChanged({
    $form.TopMost = $script:topMostCheck.Checked
})
$script:slidePicker.add_SelectedIndexChanged({
    if (-not $script:updatingSlide -and $script:slidePicker.SelectedIndex -ge 0) {
        Set-SlideContent -Index ($script:slidePicker.SelectedIndex + 1) -NavigatePowerPoint $true
    }
})
$script:previousButton.add_Click({
    Set-SlideContent -Index ($script:currentSlideIndex - 1) -NavigatePowerPoint $true
})
$script:nextButton.add_Click({
    Set-SlideContent -Index ($script:currentSlideIndex + 1) -NavigatePowerPoint $true
})
$buttonFontDown.add_Click({
    $script:baseFontSize = [Math]::Max(10, $script:baseFontSize - 1)
    foreach ($reader in $script:readerControls) {
        $reader.Font = New-Object System.Drawing.Font($fontName, $script:baseFontSize)
    }
})
$buttonFontUp.add_Click({
    $script:baseFontSize = [Math]::Min(20, $script:baseFontSize + 1)
    foreach ($reader in $script:readerControls) {
        $reader.Font = New-Object System.Drawing.Font($fontName, $script:baseFontSize)
    }
})
$minimizeButton.add_Click({
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
})
$form.add_Resize({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
        $buttonFontDown.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 220), 7)
        $buttonFontUp.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 165), 7)
        $minimizeButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 110), 7)
        $footerLabel.MaximumSize = New-Object System.Drawing.Size(($form.ClientSize.Width - 240), 34)
        $script:talkReader.Width = [Math]::Max(360, ($talkPage.ClientSize.Width - 42))
        $script:talkReader.Height = [Math]::Max(360, ($talkPage.ClientSize.Height - 64))
        $script:notesReader.Width = [Math]::Max(360, ($notesPage.ClientSize.Width - 42))
        $script:notesReader.Height = [Math]::Max(360, ($notesPage.ClientSize.Height - 64))
        $script:statusLabel.Location = New-Object System.Drawing.Point(16, 58)
        $script:statusLabel.Width = [Math]::Max(360, ($form.ClientSize.Width - 34))
        $script:displayPicker.Width = [Math]::Max(180, [Math]::Min(260, ($form.ClientSize.Width - 410)))
        $script:topMostCheck.Location = New-Object System.Drawing.Point(($script:displayPicker.Right + 12), 7)
    }
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 900
$timer.add_Tick({
    $show = Get-MatchingSlideShow
    if ($null -ne $show) {
        try {
            $index = [int]$show.View.Slide.SlideIndex
            if ($index -ge 1 -and $index -le $script:content.expectedSlideCount) {
                if ($index -ne $script:currentSlideIndex) {
                    Set-SlideContent -Index $index
                }
                else {
                    Refresh-SpeakerNotes -Index $index
                }
            }
            $slide = $script:content.slides[$script:currentSlideIndex - 1]
            $fileName = [System.IO.Path]::GetFileName($script:expectedPptxPath)
            $script:statusLabel.Text = (($script:content.ui.statusConnectedFormat -f $fileName, $script:currentSlideIndex, $script:content.expectedSlideCount) + "`r`n" + $slide.title)
            $script:statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(166, 218, 197)
        }
        catch {
            $script:slideShow = $null
            $script:pptApp = $null
            $script:statusLabel.Text = $script:content.ui.statusError
            $script:statusLabel.ForeColor = $colors.Coral
        }
    }
    else {
        Refresh-SpeakerNotes -Index $script:currentSlideIndex -Force $true
        if ($script:matchFailure) {
            $script:statusLabel.Text = ($script:content.ui.statusRejected + " · " + $script:matchFailure)
            $script:statusLabel.ForeColor = $colors.Coral
        }
        else {
            $script:statusLabel.Text = $script:content.ui.statusNotFound
            $script:statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(202, 143, 74)
        }
    }
})

Set-SlideContent -Index 1
$script:talkTab.Select()
Move-WindowToSelectedDisplay
if ($SmokeTest) {
    $form.PerformLayout()
    if ($script:tabs.TabPages.Count -ne 6 -or $script:slidePicker.Items.Count -ne 13 -or $null -eq $script:notesReader) {
        throw "Presenter UI smoke test failed."
    }
    Write-Output ("UI smoke test passed: {0} tabs, {1} slides, actual-notes tab initialized, {2} display(s)." -f $script:tabs.TabPages.Count, $script:slidePicker.Items.Count, $script:screens.Count)
    $form.Dispose()
    exit 0
}
$timer.Start()
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
$timer.Stop()

#Requires -Version 5.1
<#
.SYNOPSIS
  Builds a client-ready 16-slide PDF presentation from a brief JSON.

.DESCRIPTION
  Reads a brief JSON file, generates a 16-slide presentation with:
  - Slide 1: Product direction with hero image
  - Slide 2: Client self-summary (4 paragraphs, first person)
  - Slide 3: Meeting summary (7 paragraphs, consultant)
  - Slide 4: Personalized pricing table
  - Slide 5: Diagnostic page (from reference PDF)
  - Slide 6: Consultant intro (photo + competencies from config)
  - Slides 7-15: Static company/proof pages (from reference PDF)
  - Slide 16: Final offer / next step

  All company-specific values (brand text, consultant name, competencies,
  slide titles) are read from config.json in the project root.

.PARAMETER BriefPath
  Path to the brief JSON file.

.PARAMETER OutputPath
  Path for the intermediate PPTX file.

.PARAMETER PdfOutputPath
  Path for the final PDF file.

.PARAMETER ReferencePdfPath
  Path to the reference PDF for static slides (5, 7-15).
  Default: references/reference-presentation.pdf

.PARAMETER ConsultantPhotoPath
  Path to the consultant photo for slide 6.
  Default: references/consultant-photo.png

.PARAMETER PythonPath
  Path to Python executable for render_pdf_pages.py.
  Default: python (looks in PATH)

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File build_presentation.ps1 \
    -BriefPath briefs/Alexey_2026-06-15_brief.json \
    -OutputPath reports/Alexey_2026-06-15.pptx \
    -PdfOutputPath reports/Alexey_2026-06-15.pdf
#>

param(
  [Parameter(Mandatory = $true)]
  [string]$BriefPath,

  [Parameter(Mandatory = $true)]
  [string]$OutputPath,

  [Parameter(Mandatory = $true)]
  [string]$PdfOutputPath,

  [string]$ReferencePdfPath = 'references/reference-presentation.pdf',

  [string]$ConsultantPhotoPath = 'references/consultant-photo.png',

  [string]$PythonPath = 'python'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Load config.json ──────────────────────────────────────────────────────────
$ProjectRoot = $PSScriptRoot | Split-Path -Parent
$ConfigPath = Join-Path $ProjectRoot 'config.json'
if (-not (Test-Path -LiteralPath $ConfigPath)) {
  # Fall back to config.example.json if config.json doesn't exist
  $ConfigPath = Join-Path $ProjectRoot 'config.example.json'
}
$ConfigBytes = [System.IO.File]::ReadAllBytes($ConfigPath)
$ConfigJson = [System.Text.Encoding]::UTF8.GetString($ConfigBytes)
$Config = $ConfigJson | ConvertFrom-Json

$CompanyBrandText = if ($Config.company.brandText) { $Config.company.brandText } else { $Config.company.site -replace '^https?://', '' }
$ConsultantName = $Config.company.consultant.name
$ConsultantCompetencies = @($Config.company.consultant.competencies)
$DiscountPercent = if ($Config.defaults.pricing.oneTimeDiscountPercent) { $Config.defaults.pricing.oneTimeDiscountPercent } else { 10 }
$InstallmentMonths = if ($Config.defaults.pricing.installmentMonths) { $Config.defaults.pricing.installmentMonths } else { 24 }
$SlideTitleLine1 = if ($Config.defaults.presentation.slide1_title_line1) { $Config.defaults.presentation.slide1_title_line1 } else { 'ВАШЕ НАПРАВЛЕНИЕ ПО ИТОГАМ' }
$SlideTitleLine2 = if ($Config.defaults.presentation.slide1_title_line2) { $Config.defaults.presentation.slide1_title_line2 } else { 'ЛИЧНОЙ КОНСУЛЬТАЦИИ С ЭКСПЕРТОМ' }
$AiBadgeText = if ($Config.defaults.presentation.slide1_ai_badge_text) { $Config.defaults.presentation.slide1_ai_badge_text } else { 'AI' }
$AiBadgeDesc = if ($Config.defaults.presentation.slide1_ai_badge_description) { $Config.defaults.presentation.slide1_ai_badge_description } else { '' }

# Make paths relative to project root
if (-not [System.IO.Path]::IsPathRooted($ReferencePdfPath)) {
  $ReferencePdfPath = Join-Path $ProjectRoot $ReferencePdfPath
}
if (-not [System.IO.Path]::IsPathRooted($ConsultantPhotoPath)) {
  $ConsultantPhotoPath = Join-Path $ProjectRoot $ConsultantPhotoPath
}

Add-Type -AssemblyName System.Drawing

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RenderScript = Join-Path $ScriptDir 'render_pdf_pages.py'
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-Prop {
  param(
    [Parameter(Mandatory = $true)] $Object,
    [Parameter(Mandatory = $true)] [string]$Name,
    $Default = $null
  )
  $prop = $Object.PSObject.Properties[$Name]
  if ($null -ne $prop) { return $prop.Value }
  return $Default
}

function Read-BriefObject {
  param([Parameter(Mandatory = $true)] [string]$Path)
  $bytes = [System.IO.File]::ReadAllBytes($Path)
  $json = [System.Text.Encoding]::UTF8.GetString($bytes)
  return ($json | ConvertFrom-Json)
}

function Parse-Amount {
  param([Parameter(Mandatory = $true)] $Value)
  $digits = ([string]$Value) -replace '[^\d]', ''
  if ([string]::IsNullOrWhiteSpace($digits)) {
    throw "Could not parse amount from value: $Value"
  }
  return [int]$digits
}

function Format-Amount {
  param([Parameter(Mandatory = $true)] [int]$Value)
  $text = [string]$Value
  $chunks = @()
  while ($text.Length -gt 3) {
    $chunks = @($text.Substring($text.Length - 3)) + $chunks
    $text = $text.Substring(0, $text.Length - 3)
  }
  if ($text.Length -gt 0) { $chunks = @($text) + $chunks }
  return ($chunks -join ' ')
}

function Get-InstallmentLabel {
  param([Parameter(Mandatory = $true)] [int]$Months)
  switch ($Months) {
    12 { return '1 год' }
    24 { return '2 года' }
    default { return "$Months мес." }
  }
}

function Get-StringArray {
  param(
    $Value,
    [int]$Count,
    [string]$Name
  )
  $items = @()
  foreach ($item in @($Value)) {
    $items += [string]$item
  }
  if ($items.Count -ne $Count) {
    throw "$Name must contain exactly $Count items."
  }
  return [string[]]$items
}

function RGB {
  param([int]$R, [int]$G, [int]$B)
  return ($R + ($G * 256) + ($B * 65536))
}

function Resolve-EdgePath {
  $candidates = @(
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
  )
  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) { return $candidate }
  }
  throw 'Microsoft Edge was not found.'
}

function Save-CourseHero {
  param(
    [Parameter(Mandatory = $true)] [string]$CourseUrl,
    [Parameter(Mandatory = $true)] [string]$DestinationPath
  )
  $edgePath = Resolve-EdgePath
  $fullShotPath = [System.IO.Path]::ChangeExtension($DestinationPath, '.full.png')
  & $edgePath '--headless=new' '--disable-gpu' '--window-size=1600,1000' "--screenshot=$fullShotPath" $CourseUrl | Out-Null

  $bitmap = [System.Drawing.Bitmap]::FromFile($fullShotPath)
  try {
    # The product page usually keeps the clean hero visual on the right.
    # Crop that area first so slide 1 does not show broken page text or buttons.
    $cropLeft = [int][Math]::Floor($bitmap.Width * 0.625)
    $cropTop = [int][Math]::Floor($bitmap.Height * 0.03)
    $cropWidth = [int][Math]::Min(($bitmap.Width - $cropLeft - 30), ($bitmap.Width * 0.34))
    $cropHeight = [int][Math]::Min(($bitmap.Height * 0.68), ($bitmap.Height - $cropTop - 80))
    if ($cropWidth -lt 220 -or $cropHeight -lt 220) {
      $cropLeft = 0
      $cropTop = 0
      $cropWidth = [Math]::Min(1600, $bitmap.Width)
      $cropHeight = [Math]::Min(620, $bitmap.Height)
    }
    $rect = New-Object System.Drawing.Rectangle $cropLeft, $cropTop, $cropWidth, $cropHeight
    $cropped = $bitmap.Clone($rect, $bitmap.PixelFormat)
    try {
      $cropped.Save($DestinationPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
      $cropped.Dispose()
    }
  }
  finally {
    $bitmap.Dispose()
  }
}

function Save-CoverCrop {
  param(
    [Parameter(Mandatory = $true)] [string]$SourcePath,
    [Parameter(Mandatory = $true)] [string]$DestinationPath,
    [Parameter(Mandatory = $true)] [double]$TargetRatio,
    [double]$FocusX = 0.5,
    [double]$FocusY = 0.5
  )

  $bitmap = [System.Drawing.Bitmap]::FromFile($SourcePath)
  try {
    $sourceRatio = $bitmap.Width / $bitmap.Height
    if ($sourceRatio -gt $TargetRatio) {
      $cropHeight = $bitmap.Height
      $cropWidth = [int][Math]::Round($cropHeight * $TargetRatio)
      $cropLeft = [int][Math]::Floor(($bitmap.Width - $cropWidth) * $FocusX)
      $cropTop = 0
    }
    else {
      $cropWidth = $bitmap.Width
      $cropHeight = [int][Math]::Round($cropWidth / $TargetRatio)
      $cropLeft = 0
      $cropTop = [int][Math]::Floor(($bitmap.Height - $cropHeight) * $FocusY)
    }
    $rect = New-Object System.Drawing.Rectangle $cropLeft, $cropTop, $cropWidth, $cropHeight
    $cropped = $bitmap.Clone($rect, $bitmap.PixelFormat)
    try {
      $cropped.Save($DestinationPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
      $cropped.Dispose()
    }
  }
  finally {
    $bitmap.Dispose()
  }
}

function Add-RoundedPicture {
  param(
    [Parameter(Mandatory = $true)] $Slide,
    [Parameter(Mandatory = $true)] [string]$ImagePath,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [double]$BorderColor = -1
  )
  if (-not (Test-Path -LiteralPath $ImagePath)) {
    $placeholder = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height)
    $placeholder.Fill.ForeColor.RGB = RGB 38 42 55
    $placeholder.Line.ForeColor.RGB = RGB 169 177 190
    $placeholder.Line.Weight = 1
    return $placeholder
  }
  $pic = $Slide.Shapes.AddPicture($ImagePath, 0, -1, $Left, $Top, $Width, $Height)
  if ($BorderColor -ge 0) {
    $pic.Line.ForeColor.RGB = $BorderColor
    $pic.Line.Weight = 2.5
  }
  else {
    $pic.Line.Visible = 0
  }
  return $pic
}

function Add-Text {
  param(
    [Parameter(Mandatory = $true)] $Slide,
    [Parameter(Mandatory = $true)] [string]$Text,
    [Parameter(Mandatory = $true)] [double]$Left,
    [Parameter(Mandatory = $true)] [double]$Top,
    [Parameter(Mandatory = $true)] [double]$Width,
    [Parameter(Mandatory = $true)] [double]$Height,
    [double]$Size = 12,
    [double]$Color = -1,
    [string]$Font = 'Montserrat',
    [bool]$Bold = $false,
    [int]$Align = 1,
    [bool]$Wrap = $true
  )
  $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
  $tb.TextFrame.WordWrap = $Wrap
  if ($Align -eq 2) { $tb.TextFrame.TextRange.ParagraphFormat.Alignment = 2 }
  elseif ($Align -eq 0) { $tb.TextFrame.TextRange.ParagraphFormat.Alignment = 0 }
  $tr = $tb.TextFrame.TextRange
  $tr.Text = $Text
  $tr.Font.Size = $Size
  $tr.Font.Name = $Font
  $tr.Font.Bold = $Bold
  if ($Color -ge 0) { $tr.Font.Color.RGB = $Color }
  return $tb
}

function Add-Panel {
  param(
    [Parameter(Mandatory = $true)] $Slide,
    [Parameter(Mandatory = $true)] [double]$Left,
    [Parameter(Mandatory = $true)] [double]$Top,
    [Parameter(Mandatory = $true)] [double]$Width,
    [Parameter(Mandatory = $true)] [double]$Height,
    [double]$FillColor = (RGB 31 34 44),
    [double]$BorderColor = -1
  )
  $shape = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $FillColor
  if ($BorderColor -ge 0) {
    $shape.Line.ForeColor.RGB = $BorderColor
    $shape.Line.Weight = 1
  }
  else {
    $shape.Line.Visible = 0
  }
  return $shape
}

function New-BlankSlide {
  param(
    [Parameter(Mandatory = $true)] $Presentation,
    [int]$Number
  )
  $slide = $Presentation.Slides.Add($Number, 12)  # ppLayoutBlank = 12
  return $slide
}

function Add-Logo {
  param(
    [Parameter(Mandatory = $true)] $Slide,
    [ValidateSet('TopRight', 'BottomRight', 'None')] [string]$IconPosition = 'TopRight'
  )

  $lineColor = RGB 140 143 150
  $leftLine = $Slide.Shapes.AddLine(346, 498, 424, 498)
  $leftLine.Line.ForeColor.RGB = $lineColor
  $leftLine.Line.Transparency = 0.3
  $leftLine.Line.Weight = 1
  $rightLine = $Slide.Shapes.AddLine(535, 498, 612, 498)
  $rightLine.Line.ForeColor.RGB = $lineColor
  $rightLine.Line.Transparency = 0.3
  $rightLine.Line.Weight = 1
  [void](Add-Text -Slide $Slide -Text $CompanyBrandText -Left 432 -Top 488 -Width 96 -Height 18 -Size 13 -Color $lineColor -Font 'Arial' -Align 2)

  if ($IconPosition -eq 'None') { return }

  $x = 874
  $y = if ($IconPosition -eq 'BottomRight') { 458 } else { 35 }
  $outer = $Slide.Shapes.AddShape(10, $x, $y, 40, 50)
  $outer.Fill.ForeColor.RGB = $Global:Green
  $outer.Line.Visible = 0
  $cutout = $Slide.Shapes.AddShape(1, $x + 23, $y + 10, 12, 30)
  $cutout.Fill.ForeColor.RGB = $Global:Bg
  $cutout.Line.Visible = 0
}

function Add-Background {
  param([Parameter(Mandatory = $true)] $Slide)
  $bg = $Slide.Shapes.AddShape(1, 0, 0, $Global:W, $Global:H)
  $bg.Fill.ForeColor.RGB = $Global:Bg
  $bg.Line.Visible = 0
}

function Add-StaticPage {
  param(
    [Parameter(Mandatory = $true)] $Slide,
    [Parameter(Mandatory = $true)] [string]$ImagePath
  )
  [void]$Slide.Shapes.AddPicture($ImagePath, 0, -1, 0, 0, $Global:W, $Global:H)
}

# ── Main ──────────────────────────────────────────────────────────────────────

$brief = Read-BriefObject -Path $BriefPath
$studentName = [string](Get-Prop $brief 'clientName' (Get-Prop $brief 'studentName'))
$courseUrl = [string](Get-Prop $brief 'productUrl' (Get-Prop $brief 'courseUrl'))
$courseTitle = [string](Get-Prop $brief 'productTitle' (Get-Prop $brief 'courseTitle'))
$courseTitlePrefix = [string](Get-Prop $brief 'productTitlePrefix' (Get-Prop $brief 'courseTitlePrefix'))
$slide2Paragraphs = Get-StringArray -Value (Get-Prop $brief 'slide2Paragraphs') -Count 4 -Name 'slide2Paragraphs'
$slide3Paragraphs = Get-StringArray -Value (Get-Prop $brief 'slide3Paragraphs') -Count 7 -Name 'slide3Paragraphs'
$pricing = $brief.pricing
$fullPrice = Parse-Amount (Get-Prop $pricing 'fullPrice')
$specialPrice = Parse-Amount (Get-Prop $pricing 'specialPrice')
$oneTimeDiscountPercent = [int](Get-Prop $pricing 'oneTimeDiscountPercent' $DiscountPercent)
$oneTimePrice = [int][Math]::Floor($specialPrice * (1 - $oneTimeDiscountPercent / 100))
$productName = [string](Get-Prop $pricing 'productName')
$installmentMonths = [int](Get-Prop $pricing 'installmentMonths' $InstallmentMonths)
$installment = [int][Math]::Floor($specialPrice / $installmentMonths)
$installmentLabel = Get-InstallmentLabel $installmentMonths
$specialPriceHeader = [string](Get-Prop $pricing 'specialPriceHeader' (Get-Prop $Config.defaults.pricing 'specialPriceHeader' 'Цена в рассрочку'))
$oneTimePriceHeader = [string](Get-Prop $pricing 'oneTimePriceHeader' (Get-Prop $Config.defaults.pricing 'oneTimePriceHeader' 'Цена при единовременной оплате'))

$Global:W = 960
$Global:H = 540
$Global:Bg = RGB 18 20 26
$Global:Panel = RGB 31 34 44
$Global:Panel2 = RGB 38 42 55
$Global:Green = RGB 71 236 132
$Global:Purple = RGB 115 67 255
$Global:White = RGB 246 248 255
$Global:Muted = RGB 169 177 190
$Global:Orange = RGB 255 168 77

$workRoot = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'tri-bloka-presentation-' + [guid]::NewGuid().ToString('N'))
$staticDir = Join-Path $workRoot 'static'
$heroPath = Join-Path $workRoot 'product_hero.png'
$heroCoverPath = Join-Path $workRoot 'product_hero_cover.png'
$photoCoverPath = Join-Path $workRoot 'consultant_photo_cover.png'

try {
  New-Item -ItemType Directory -Force -Path $workRoot, $staticDir | Out-Null

  Save-CourseHero -CourseUrl $courseUrl -DestinationPath $heroPath
  Save-CoverCrop -SourcePath $heroPath -DestinationPath $heroCoverPath -TargetRatio (352 / 246) -FocusX 0.5 -FocusY 0.42
  if (Test-Path -LiteralPath $ConsultantPhotoPath) {
    Save-CoverCrop -SourcePath $ConsultantPhotoPath -DestinationPath $photoCoverPath -TargetRatio (300 / 392) -FocusX 0.5 -FocusY 0.18
  }
  & $PythonPath -X utf8 $RenderScript $ReferencePdfPath $staticDir 5 7 8 9 10 11 12 13 14 15 | Out-Null

  if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
  if (Test-Path -LiteralPath $PdfOutputPath) { Remove-Item -LiteralPath $PdfOutputPath -Force }

  $powerPoint = $null
  $presentation = $null
  try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add()
    $presentation.PageSetup.SlideWidth = $Global:W
    $presentation.PageSetup.SlideHeight = $Global:H

    # 1. Direction
    $s = New-BlankSlide $presentation 1
    Add-Background $s
    Add-Logo $s
    [void](Add-Text $s $SlideTitleLine1 122 36 716 40 30 $Global:Green 'Montserrat' $true 2)
    [void](Add-Text $s $SlideTitleLine2 122 78 716 38 30 $Global:Green 'Montserrat' $true 2)
    [void](Add-Panel $s 20 126 920 302 (RGB 48 24 125))
    $labelText = if ([string]::IsNullOrWhiteSpace($courseTitlePrefix)) { 'ПРОДУКТ' } else { $courseTitlePrefix.ToUpper() }
    [void](Add-Text $s $labelText 32 162 330 26 20 (RGB 226 255 47) 'Montserrat' $true)
    $courseDisplay = $courseTitle.ToUpper()
    $titleSize = if ($courseDisplay.Length -gt 24) { 34 } else { 38 }
    $courseTitleShape = Add-Text $s $courseDisplay 32 204 444 112 $titleSize $Global:White 'Montserrat' $true
    $courseTitleShape.ActionSettings.Item(1).Hyperlink.Address = $courseUrl
    [void](Add-Panel $s 36 342 442 68 $Global:Panel)
    [void](Add-Panel $s 48 348 76 58 (RGB 48 54 63) (RGB 226 255 47))
    [void](Add-Text $s $AiBadgeText 62 360 48 26 24 (RGB 226 255 47) 'Montserrat' $true 2)
    if ($AiBadgeDesc) {
      [void](Add-Text $s $AiBadgeDesc 154 350 300 46 14 $Global:White 'Montserrat')
    }
    $heroShape = Add-RoundedPicture $s $heroCoverPath 548 150 352 246 $Global:Purple
    $heroShape.ActionSettings.Item(1).Hyperlink.Address = $courseUrl

    # 2. Client self-summary
    $s = New-BlankSlide $presentation 2
    Add-Background $s
    Add-Logo $s
    [void](Add-Text $s 'Имя:' 30 34 78 28 20 $Global:Green 'Montserrat' $true)
    [void](Add-Text $s $studentName 112 32 270 32 24 $Global:White 'Montserrat' $true)
    $slide2Text = ($slide2Paragraphs -join "`r`n`r`n")
    [void](Add-Text $s $slide2Text 26 94 886 330 15.8 $Global:White 'Montserrat')

    # 3. Meeting summary
    $s = New-BlankSlide $presentation 3
    Add-Background $s
    Add-Logo $s 'BottomRight'
    [void](Add-Text $s 'Резюме встречи' 0 14 960 42 28 $Global:Green 'Montserrat' $true 2)
    $slide3Text = ($slide3Paragraphs -join "`r`n`r`n")
    [void](Add-Text $s $slide3Text 28 74 904 388 13.6 $Global:White 'Montserrat')

    # 4. Personalized pricing
    $s = New-BlankSlide $presentation 4
    Add-Background $s
    Add-Logo $s 'BottomRight'
    [void](Add-Text $s 'Персональное предложение после консультации' 38 70 870 46 29 $Global:Green 'Montserrat' $true)
    $headers = @(
      ('Рекомендации эксперта' + "`r`n" + 'по направлению' + "`r`n" + 'развития'),
      'Полная цена',
      $specialPriceHeader,
      $oneTimePriceHeader,
      ('Размер' + "`r`n" + 'платежа' + "`r`n" + 'в рассрочку' + "`r`n" + "на $installmentLabel")
    )
    $values = @($productName.ToUpper(), ((Format-Amount $fullPrice) + ' руб.'), ((Format-Amount $specialPrice) + ' руб.'), ((Format-Amount $oneTimePrice) + ' руб.'), ((Format-Amount $installment) + ' руб.'))
    $tableLeft = 12
    $tableTop = 196
    $tableWidth = 936
    $headerHeight = 112
    $valueHeight = 104
    $colWidths = @(284, 164, 158, 176, 154)
    $lineColor = RGB 154 157 165
    $outline = $s.Shapes.AddShape(1, $tableLeft, $tableTop, $tableWidth, ($headerHeight + $valueHeight))
    $outline.Fill.Visible = 0
    $outline.Line.ForeColor.RGB = $lineColor
    $outline.Line.Weight = 0.8
    $mid = $s.Shapes.AddLine($tableLeft, ($tableTop + $headerHeight), ($tableLeft + $tableWidth), ($tableTop + $headerHeight))
    $mid.Line.ForeColor.RGB = $lineColor
    $mid.Line.Weight = 0.8
    $x = $tableLeft
    for ($i = 0; $i -lt ($colWidths.Count - 1); $i++) {
      $x += $colWidths[$i]
      $vline = $s.Shapes.AddLine($x, $tableTop, $x, ($tableTop + $headerHeight + $valueHeight))
      $vline.Line.ForeColor.RGB = $lineColor
      $vline.Line.Weight = 0.8
    }
    $x = $tableLeft
    for ($i = 0; $i -lt 5; $i++) {
      [void](Add-Text $s $headers[$i] ($x + 14) ($tableTop + 24) ($colWidths[$i] - 28) 72 15 (RGB 170 128 255) 'Montserrat' $true 2 $true)
      $valueSize = if ($i -eq 0) { 16 } else { 19 }
      [void](Add-Text $s $values[$i] ($x + 12) ($tableTop + $headerHeight + 42) ($colWidths[$i] - 24) 34 $valueSize $Global:Green 'Montserrat' $true 2)
      $x += $colWidths[$i]
    }

    # 5. Diagnostic page from reference
    $s = New-BlankSlide $presentation 5
    Add-StaticPage $s (Join-Path $staticDir 'page_05.png')

    # 6. Consultant intro
    $s = New-BlankSlide $presentation 6
    Add-Background $s
    Add-Logo $s
    $rightGlow = $s.Shapes.AddShape(1, 562, 0, 398, 540)
    $rightGlow.Fill.ForeColor.RGB = RGB 48 28 126
    $rightGlow.Fill.Transparency = 0.38
    $rightGlow.Line.Visible = 0
    if (Test-Path -LiteralPath $ConsultantPhotoPath) {
      [void](Add-RoundedPicture $s $photoCoverPath 20 70 300 392 $Global:Green)
    }
    else {
      [void](Add-Text $s 'Место для фото' 78 236 190 30 18 $Global:Muted 'Montserrat' $true 2)
    }
    [void](Add-Text $s 'Давайте знакомиться' 352 20 430 38 30 $Global:Green 'Montserrat' $true)
    [void](Add-Text $s 'Ваш личный карьерный консультант:' 352 72 520 34 22 $Global:Green 'Montserrat' $true)
    [void](Add-Text $s $ConsultantName 352 106 410 42 31 (RGB 170 128 255) 'Montserrat' $true)
    [void](Add-Text $s 'Основные компетенции:' 540 158 310 30 20 $Global:White 'Montserrat' $true 2)
    $y = 204
    foreach ($item in $ConsultantCompetencies) {
      [void](Add-Text $s '✓' 354 ($y - 4) 46 34 29 $Global:Green 'Montserrat' $true 2)
      [void](Add-Text $s $item 408 $y 476 42 14.5 $Global:White 'Montserrat')
      $y += 56
    }

    # 7-15. Static proof/company/path/special-offer pages from reference style
    $staticPages = @(7, 8, 9, 10, 11, 12, 13, 14, 15)
    for ($idx = 0; $idx -lt $staticPages.Count; $idx++) {
      $s = New-BlankSlide $presentation (7 + $idx)
      Add-StaticPage $s (Join-Path $staticDir ('page_{0:00}.png' -f $staticPages[$idx]))
    }

    # 16. Final offer / next step
    $s = New-BlankSlide $presentation 16
    Add-Background $s
    Add-Logo $s
    [void](Add-Text $s 'Итоговое предложение после консультации' 58 58 560 34 22 $Global:White 'Montserrat' $true)
    [void](Add-Text $s $courseTitle.ToUpper() 58 112 520 28 16 $Global:Green 'Montserrat' $true)
    [void](Add-Panel $s 58 178 330 184 $Global:Panel $Global:Purple)
    [void](Add-Text $s ((Format-Amount $specialPrice) + ' руб.') 86 210 270 46 30 $Global:White 'Montserrat' $true 2)
    [void](Add-Text $s 'цена в рассрочку' 86 260 270 24 13 $Global:Muted 'Montserrat' $false 2)
    [void](Add-Text $s ((Format-Amount $installment) + ' руб./мес') 86 304 270 32 22 $Global:Green 'Montserrat' $true 2)
    [void](Add-Text $s "платеж на $installmentLabel" 86 338 270 22 12 $Global:Muted 'Montserrat' $false 2)
    [void](Add-Panel $s 430 178 430 184 $Global:Panel2)
    [void](Add-Text $s 'Что фиксируем как следующий шаг' 462 204 360 28 17 $Global:Green 'Montserrat' $true)
    [void](Add-Text $s 'Посмотрите программу, отметьте вопросы по модулям и формату, а на следующем касании спокойно обсудим, насколько маршрут попадает в вашу цель.' 462 248 340 82 14 $Global:White 'Montserrat')
    [void](Add-Text $s 'Без давления - задача презентации в том, чтобы у вас перед глазами был понятный маршрут и условия.' 462 346 350 46 12 $Global:Muted 'Montserrat')

    $presentation.SaveAs($OutputPath, 24)
    $presentation.SaveAs($PdfOutputPath, 32)
  }
  finally {
    if ($null -ne $presentation) {
      $presentation.Close()
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($null -ne $powerPoint) {
      $powerPoint.Quit()
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }

  Write-Output "Built presentation: $OutputPath"
  Write-Output "Exported PDF: $PdfOutputPath"
}
finally {
  if (Test-Path -LiteralPath $workRoot) {
    Remove-Item -LiteralPath $workRoot -Recurse -Force
  }
}
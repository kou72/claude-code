# pptx_read.ps1
# Convert PowerPoint slides to PNG images and extract text content.
#
# Usage:
#   powershell.exe -ExecutionPolicy Bypass -File pptx_read.ps1 -PptxPath "path\to\file.pptx"
#   powershell.exe -ExecutionPolicy Bypass -File pptx_read.ps1 -PptxPath "path\to\file.pptx" -OutDir "path\to\output"
#
# Output:
#   {OutDir}/slides/*.PNG   - slide images
#   {OutDir}/content.md     - extracted text with image references

param(
    [Parameter(Mandatory=$true)]
    [string]$PptxPath,

    [string]$OutDir = ""
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$PptxPath = (Resolve-Path $PptxPath).Path

if (-not (Test-Path $PptxPath)) {
    Write-Error "File not found: $PptxPath"
    exit 1
}

if ($OutDir -eq "") {
    $OutDir = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($PptxPath),
        [System.IO.Path]::GetFileNameWithoutExtension($PptxPath)
    )
}
$SlidesDir = Join-Path $OutDir "slides"

if (-not (Test-Path $SlidesDir)) {
    New-Item -ItemType Directory -Path $SlidesDir | Out-Null
}

Write-Output "=== pptx_read ==="
Write-Output "Input : $PptxPath"
Write-Output "Output: $OutDir"
Write-Output ""

# --- Step 1: Export slides as PNG via PowerPoint COM ---
Write-Output "[Step 1] Exporting slides as PNG..."
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $presentation = $ppt.Presentations.Open($PptxPath, $false, $false, $true)
    $presentation.SaveAs($SlidesDir, 18)  # 18 = ppSaveAsPNG
    $presentation.Close()
    $ppt.Quit()
    $pngFiles = Get-ChildItem $SlidesDir -Filter "*.PNG" | Sort-Object { [int]($_.BaseName -replace '[^0-9]', '') }
    $count = $pngFiles.Count
    Write-Output "  -> $count PNG files saved to: $SlidesDir"
} catch {
    Write-Error "PNG export failed (make sure PowerPoint is closed): $_"
    exit 1
}

# --- Step 2: Extract text via ZIP + XML ---
Write-Output ""
Write-Output "[Step 2] Extracting text..."

Add-Type -Assembly 'System.IO.Compression.FileSystem'
$zip = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
$slideFiles = $zip.Entries |
    Where-Object { $_.FullName -match '^ppt/slides/slide[0-9]+\.xml$' } |
    Sort-Object { [int]($_.Name -replace '[^0-9]', '') }

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($PptxPath)
$mdLines = @()
$mdLines += "# $baseName"
$mdLines += ""
$mdLines += "> Auto-extracted by pptx_read.ps1"
$mdLines += ""

$slideNum = 1
foreach ($entry in $slideFiles) {
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $xml = [xml]$reader.ReadToEnd()
    $reader.Close()

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $texts = $xml.SelectNodes('//a:t', $ns) |
        ForEach-Object { $_.InnerText } |
        Where-Object { $_.Trim() -ne "" }

    $pngName = $pngFiles[$slideNum - 1].Name
    $mdLines += "## Slide $slideNum"
    $mdLines += ""
    $mdLines += "![Slide $slideNum](slides/$pngName)"
    $mdLines += ""
    if ($texts) {
        $mdLines += ($texts -join " ")
    } else {
        $mdLines += "(no text)"
    }
    $mdLines += ""
    $slideNum++
}

$zip.Dispose()

$contentPath = Join-Path $OutDir "content.md"
$mdLines | Out-File -FilePath $contentPath -Encoding UTF8
Write-Output "  -> Text saved to: $contentPath"

Write-Output ""
Write-Output "=== Done ==="
Write-Output "PNG : $SlidesDir"
Write-Output "Text: $contentPath"

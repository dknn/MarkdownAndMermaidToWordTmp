<#
.SYNOPSIS
  End-to-end: Render Mermaid i Markdown til billeder og konverter til DOCX via Pandoc.

.REQUIREMENTS
  - pandoc i PATH
  - mmdc i PATH
#>

[CmdletBinding()]
param(
  [Parameter()]
  [string] $InputMarkdown,

  [Parameter()]
  [string] $OutputDocx,

  # Hvor vi lægger genererede assets og midlertidig md
  [string] $WorkDir = ".\_build",

  # Brug png for Word
  [ValidateSet('png','svg')]
  [string] $ImageFormat = 'png',

  # Valgfrit: Word template (reference docx) for styling
  [string] $ReferenceDocx = ""
)

function Assert-Command([string]$Name) {
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Mangler '$Name' i PATH."
  }
}

Assert-Command "pandoc"
Assert-Command "mmdc"

if (-not (Test-Path $InputMarkdown)) { throw "InputMarkdown findes ikke: $InputMarkdown" }

New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
$assets = Join-Path $WorkDir "_assets\mermaid"
New-Item -ItemType Directory -Force -Path $assets | Out-Null

$renderedMd = Join-Path $WorkDir "doc.rendered.md"

# 1) Render Mermaid blocks -> billeder + md med image refs
.\Render-MermaidInMarkdown.ps1 `
  -InputMarkdown $InputMarkdown `
  -OutMarkdown   $renderedMd `
  -AssetsDir     $assets `
  -ImageFormat   $ImageFormat `
  -BackgroundColor "white"

# 2) Pandoc -> DOCX
$pandocArgs = @(
  $renderedMd,
  "-o", $OutputDocx,
  "--from", "gfm+pipe_tables+task_lists",
  "--resource-path", "$WorkDir;."
)

# Valgfrit reference doc (Word template)
if ($ReferenceDocx -and (Test-Path $ReferenceDocx)) {
  $pandocArgs += @("--reference-doc", $ReferenceDocx)
}

Write-Host "pandoc -> $OutputDocx"
& pandoc @pandocArgs
if ($LASTEXITCODE -ne 0) { throw "pandoc fejlede (exit $LASTEXITCODE)" }

Write-Host "Done: $OutputDocx"

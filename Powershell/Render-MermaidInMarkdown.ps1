<#
.SYNOPSIS
  Finder ```mermaid``` code fences i en Markdown-fil, renderer hver til PNG via mermaid-cli (mmdc),
  og erstatter blokken med en Markdown image-reference.

.REQUIREMENTS
  - Node + npm
  - @mermaid-js/mermaid-cli (mmdc) i PATH
  - (Valgfrit) mermaid config json, fx til tema

.EXAMPLE
  .\Render-MermaidInMarkdown.ps1 -InputMarkdown .\doc.md -OutMarkdown .\doc.rendered.md -AssetsDir .\_assets\mermaid
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string] $InputMarkdown,

  [Parameter(Mandatory)]
  [string] $OutMarkdown,

  [Parameter(Mandatory)]
  [string] $AssetsDir,

  # PNG er bedst til Word (DOCX). SVG kan give problemer i Word.
  [ValidateSet('png','svg')]
  [string] $ImageFormat = 'png',

  # Hvis du vil styre tema, fonts, etc: peg på en Mermaid config JSON
  # https://mermaid.js.org/config/theming.html
  [string] $MermaidConfigJson = "",

  # Baggrund: Word håndterer typisk bedre en solid baggrund end transparent.
  [string] $BackgroundColor = "white"
)

function Assert-Command([string]$Name) {
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Mangler '$Name' i PATH. Installer det (fx: npm i -g @mermaid-js/mermaid-cli) og prøv igen."
  }
}

Assert-Command "mmdc"

if (-not (Test-Path $InputMarkdown)) { throw "InputMarkdown findes ikke: $InputMarkdown" }

New-Item -ItemType Directory -Force -Path $AssetsDir | Out-Null

$md = Get-Content -Raw -LiteralPath $InputMarkdown

# Regex: ```mermaid ... ```
$pattern = '(?ms)```mermaid\s*(?<code>.*?)\s*```'

$index = 0
$updated = [System.Text.StringBuilder]::new()

$pos = 0
$matches = [regex]::Matches($md, $pattern)

foreach ($m in $matches) {
  # append alt før match
  $updated.Append($md.Substring($pos, $m.Index - $pos)) | Out-Null

  $code = $m.Groups['code'].Value.Trim()

  # Stabilt filnavn: hash af indholdet (så samme diagram genbruges)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($code)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  $hash = ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join ''
  $short = $hash.Substring(0, 12)

  $index++
  $baseName = "diagram-$index-$short"
  $imgPath = Join-Path $AssetsDir "$baseName.$ImageFormat"
  $tmpMmd  = Join-Path $AssetsDir "$baseName.mmd"

  # Skriv midlertidig .mmd
  Set-Content -LiteralPath $tmpMmd -Value $code -Encoding UTF8

  # Render kun hvis billedet ikke findes
  if (-not (Test-Path $imgPath)) {
    $args = @(
      "-i", $tmpMmd,
      "-o", $imgPath,
      "-e", $ImageFormat,
      "-b", $BackgroundColor
    )

    if ($MermaidConfigJson -and (Test-Path $MermaidConfigJson)) {
      $args += @("-c", $MermaidConfigJson)
    }

    Write-Host "Rendering Mermaid -> $imgPath"
    & mmdc @args
    if ($LASTEXITCODE -ne 0) { throw "mmdc fejlede for $tmpMmd (exit $LASTEXITCODE)" }
  }

  # Relativ sti fra OutMarkdown til AssetsDir (simplere: skriv det som den relative path du giver)
  # Her indsætter vi den sti du gav (AssetsDir) som tekst. Hvis du vil være helt korrekt ift. relativ path,
  # så giv AssetsDir relativt (fx .\_assets\mermaid).
  $imgRef = Join-Path $AssetsDir "$baseName.$ImageFormat"
  $imgRef = $imgRef -replace '\\','/'  # Markdown foretrækker /

  # Erstat Mermaid blok med billede
  $updated.AppendLine("![Mermaid diagram $index]($imgRef)") | Out-Null
  $updated.AppendLine() | Out-Null

  # næste position efter match
  $pos = $m.Index + $m.Length
}

# append resten
$updated.Append($md.Substring($pos)) | Out-Null

Set-Content -LiteralPath $OutMarkdown -Value $updated.ToString() -Encoding UTF8
Write-Host "Wrote: $OutMarkdown"
Write-Host "Assets: $AssetsDir"

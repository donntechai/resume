param(
  [Parameter(Mandatory = $true)]
  [string]$Path
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path)) {
  throw "File not found: $Path"
}

$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false

  # Open read-only, do not add to recent files.
  $doc = $word.Documents.Open($Path, $false, $true, $false)

  $text = $doc.Content.Text

  # Normalize line endings and trim trailing whitespace.
  $text = ($text -replace "`r`n", "`n" -replace "`r", "`n")
  $lines = $text.Split("`n") | ForEach-Object { $_.TrimEnd() }

  # Collapse excessive blank lines.
  $out = New-Object System.Collections.Generic.List[string]
  $blankRun = 0
  foreach ($line in $lines) {
    if ([string]::IsNullOrWhiteSpace($line)) {
      $blankRun++
      if ($blankRun -le 1) { $out.Add("") }
    } else {
      $blankRun = 0
      $out.Add($line.Trim())
    }
  }

  ($out -join "`n").Trim() + "`n"
}
finally {
  if ($doc -ne $null) { $doc.Close($false) | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
}


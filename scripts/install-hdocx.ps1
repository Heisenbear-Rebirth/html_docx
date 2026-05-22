param(
  [switch]$All,
  [switch]$Tool,
  [switch]$Codex,
  [switch]$Claude,
  [switch]$Force,
  [switch]$AddToUserPath,
  [switch]$DryRun,
  [string]$ToolHome,
  [string]$CodexHome,
  [string]$ClaudeHome,
  [string]$Python = "python"
)

$ErrorActionPreference = "Stop"

function Write-Step($Message) {
  Write-Host "[hdocx] $Message"
}

function Resolve-RepoRoot {
  $scriptPath = Split-Path -Parent $PSCommandPath
  return (Resolve-Path (Join-Path $scriptPath "..")).Path
}

function Ensure-Directory($Path) {
  if ($DryRun) {
    Write-Step "Would create directory: $Path"
    return
  }
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

function Copy-Skill($Source, $Destination) {
  if (-not (Test-Path -LiteralPath $Source)) {
    throw "Missing skill source: $Source"
  }

  if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) {
      throw "Destination exists: $Destination. Re-run with -Force to replace it."
    }
    if ($DryRun) {
      Write-Step "Would remove existing skill: $Destination"
    } else {
      Remove-Item -LiteralPath $Destination -Recurse -Force
    }
  }

  Ensure-Directory (Split-Path -Parent $Destination)
  if ($DryRun) {
    Write-Step "Would copy skill: $Source -> $Destination"
  } else {
    Copy-Item -LiteralPath $Source -Destination $Destination -Recurse
  }
}

function Install-Tool($RepoRoot, $ResolvedToolHome) {
  $venvPath = Join-Path $ResolvedToolHome "venv"
  $binPath = Join-Path $ResolvedToolHome "bin"
  $tmpPath = Join-Path $ResolvedToolHome "tmp"
  $cachePath = Join-Path $ResolvedToolHome "pip-cache"
  $pythonExe = Join-Path $venvPath "Scripts\python.exe"
  $htmlDocxExe = Join-Path $venvPath "Scripts\html-docx.exe"
  $cmdShim = Join-Path $binPath "html-docx.cmd"
  $psShim = Join-Path $binPath "html-docx.ps1"

  Ensure-Directory $ResolvedToolHome
  Ensure-Directory $binPath
  Ensure-Directory $tmpPath
  Ensure-Directory $cachePath

  if ($DryRun) {
    Write-Step "Would create venv: $venvPath"
    Write-Step "Would install package from: $RepoRoot"
  } else {
    if (-not (Test-Path -LiteralPath $pythonExe)) {
      & $Python -m venv $venvPath
    }
    $env:TMP = $tmpPath
    $env:TEMP = $tmpPath
    $env:PIP_CACHE_DIR = $cachePath
    $env:PIP_DISABLE_PIP_VERSION_CHECK = "1"
    & $pythonExe -m pip install --force-reinstall --no-deps $RepoRoot
    if ($LASTEXITCODE -ne 0) {
      throw "pip install failed with exit code $LASTEXITCODE"
    }

    $cmdContent = "@echo off`r`n""%~dp0..\venv\Scripts\html-docx.exe"" %*`r`n"
    Set-Content -LiteralPath $cmdShim -Value $cmdContent -Encoding ASCII

    $psContent = @'
$exe = Join-Path $PSScriptRoot "..\venv\Scripts\html-docx.exe"
& $exe @args
exit $LASTEXITCODE
'@
    Set-Content -LiteralPath $psShim -Value $psContent -Encoding UTF8
  }

  Write-Step "CLI shim directory: $binPath"

  if ($AddToUserPath) {
    $currentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $parts = @()
    if ($currentUserPath) {
      $parts = $currentUserPath -split ";"
    }
    $alreadyPresent = $parts | Where-Object {
      $_.TrimEnd("\") -ieq $binPath.TrimEnd("\")
    }

    if ($alreadyPresent) {
      Write-Step "User PATH already contains: $binPath"
    } elseif ($DryRun) {
      Write-Step "Would add to user PATH: $binPath"
    } else {
      $newUserPath = if ($currentUserPath) { "$currentUserPath;$binPath" } else { $binPath }
      [Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")
      Write-Step "Added to user PATH. Open a new terminal before running html-docx."
    }
  } else {
    Write-Step "To use without editing PATH in this terminal:"
    Write-Host "`$env:PATH = `"$binPath;`$env:PATH`""
  }
}

if (-not ($All -or $Tool -or $Codex -or $Claude)) {
  $All = $true
}

if ($All) {
  $Tool = $true
  $Codex = $true
  $Claude = $true
}

$repoRoot = Resolve-RepoRoot

if (-not $ToolHome) {
  $ToolHome = Join-Path $HOME ".hdocx"
}
if (-not $CodexHome) {
  if ($env:CODEX_HOME) {
    $CodexHome = $env:CODEX_HOME
  } else {
    $CodexHome = Join-Path $HOME ".codex"
  }
}
if (-not $ClaudeHome) {
  $ClaudeHome = Join-Path $HOME ".claude"
}

$ToolHome = [System.IO.Path]::GetFullPath($ToolHome)
$CodexHome = [System.IO.Path]::GetFullPath($CodexHome)
$ClaudeHome = [System.IO.Path]::GetFullPath($ClaudeHome)

Write-Step "Repository: $repoRoot"

if ($Tool) {
  Install-Tool $repoRoot $ToolHome
}

if ($Codex) {
  $codexSource = Join-Path $repoRoot "skills\hdocx-agent"
  $codexDest = Join-Path $CodexHome "skills\hdocx-agent"
  Copy-Skill $codexSource $codexDest
  if ($DryRun) {
    Write-Step "Would install Codex skill: $codexDest"
  } else {
    Write-Step "Installed Codex skill: $codexDest"
  }
}

if ($Claude) {
  $claudeSource = Join-Path $repoRoot ".claude\skills\hdocx-agent"
  $claudeDest = Join-Path $ClaudeHome "skills\hdocx-agent"
  Copy-Skill $claudeSource $claudeDest
  if ($DryRun) {
    Write-Step "Would install Claude Code skill: $claudeDest"
  } else {
    Write-Step "Installed Claude Code skill: $claudeDest"
  }
}

Write-Step "Done."
if ($Codex) {
  Write-Step "Restart Codex so it discovers newly installed skills."
}
if ($Claude) {
  Write-Step "Restart Claude Code if it was already running before this install."
}

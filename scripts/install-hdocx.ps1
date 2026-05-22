param(
  [switch]$All,
  [switch]$Tool,
  [switch]$Codex,
  [switch]$Claude,
  [switch]$Force,
  [switch]$AddToUserPath,
  [switch]$CodexFallbackAgent,
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

function Test-Skill($Path) {
  $skillFile = Join-Path $Path "SKILL.md"
  if (-not (Test-Path -LiteralPath $skillFile)) {
    return @{
      ok = $false
      path = $Path
      reason = "missing SKILL.md"
    }
  }

  $content = Get-Content -LiteralPath $skillFile -Raw
  if ($content -notmatch "(?ms)^---\s*\r?\n.*?^name:\s*hdocx-agent\s*$.*?^description:\s*.+?\r?\n---") {
    return @{
      ok = $false
      path = $Path
      reason = "invalid or missing YAML frontmatter"
    }
  }

  return @{
    ok = $true
    path = $Path
    reason = $null
  }
}

function Update-CodexFallbackAgent($ResolvedCodexHome, $ResolvedToolHome) {
  $agentFile = Join-Path $ResolvedCodexHome "AGENTS.md"
  $skillPath = Join-Path $ResolvedCodexHome "skills\hdocx-agent\SKILL.md"
  $cliPath = Join-Path $ResolvedToolHome "bin\html-docx.cmd"
  $startMarker = "<!-- hdocx-agent-install:start -->"
  $endMarker = "<!-- hdocx-agent-install:end -->"
  $block = @"
$startMarker

## H-DOCX Agent Fallback

When a task asks to inspect, edit, round-trip, validate, or pressure-test DOCX
files through H-DOCX/html_docx, use the installed H-DOCX skill and CLI:

- Skill: ``$skillPath``
- CLI: ``$cliPath``

If the Codex skill list does not show `hdocx-agent`, manually read the skill
file above and follow it as the workflow. Keep all task files inside the active
workspace unless the user explicitly approves otherwise.

$endMarker
"@

  Ensure-Directory $ResolvedCodexHome

  if (Test-Path -LiteralPath $agentFile) {
    $existing = Get-Content -LiteralPath $agentFile -Raw -ErrorAction Stop
  } else {
    $existing = ""
  }

  $pattern = "(?s)\r?\n?$([regex]::Escape($startMarker)).*?$([regex]::Escape($endMarker))\r?\n?"
  $clean = [regex]::Replace($existing, $pattern, "`r`n")
  $newContent = ($clean.TrimEnd() + "`r`n`r`n" + $block.TrimEnd() + "`r`n")

  if ($DryRun) {
    Write-Step "Would update Codex fallback AGENTS.md: $agentFile"
    return
  }

  if ((Test-Path -LiteralPath $agentFile) -and ($existing -notmatch [regex]::Escape($startMarker))) {
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backup = "$agentFile.hdocx-backup-$stamp"
    Copy-Item -LiteralPath $agentFile -Destination $backup -Force
    Write-Step "Backed up Codex AGENTS.md: $backup"
  }

  Set-Content -LiteralPath $agentFile -Value $newContent -Encoding UTF8
  Write-Step "Updated Codex fallback AGENTS.md: $agentFile"
}

function Write-InstallReport($ResolvedToolHome, $ResolvedCodexHome, $ResolvedClaudeHome) {
  $binPath = Join-Path $ResolvedToolHome "bin"
  $htmlDocxCmd = Join-Path $binPath "html-docx.cmd"
  $codexSkill = Join-Path $ResolvedCodexHome "skills\hdocx-agent"
  $claudeSkill = Join-Path $ResolvedClaudeHome "skills\hdocx-agent"
  $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
  $pathContainsBin = $false
  if ($userPath) {
    $pathContainsBin = @(($userPath -split ";") | Where-Object {
      $_.TrimEnd("\") -ieq $binPath.TrimEnd("\")
    }).Count -gt 0
  }

  $cliDoctor = @{
    ok = $false
    exitCode = $null
    output = $null
  }

  if ((Test-Path -LiteralPath $htmlDocxCmd) -and (-not $DryRun)) {
    $doctorOutput = & $htmlDocxCmd doctor 2>&1
    $cliDoctor.exitCode = $LASTEXITCODE
    $cliDoctor.ok = ($LASTEXITCODE -eq 0)
    $cliDoctor.output = ($doctorOutput | Out-String).Trim()
  }

  $report = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    toolHome = $ResolvedToolHome
    cli = [ordered]@{
      binPath = $binPath
      command = $htmlDocxCmd
      exists = Test-Path -LiteralPath $htmlDocxCmd
      userPathContainsBin = $pathContainsBin
      doctor = $cliDoctor
    }
    codex = [ordered]@{
      home = $ResolvedCodexHome
      skill = Test-Skill $codexSkill
      fallbackAgent = Join-Path $ResolvedCodexHome "AGENTS.md"
      fallbackAgentContainsHdocx = if (Test-Path -LiteralPath (Join-Path $ResolvedCodexHome "AGENTS.md")) {
        ((Get-Content -LiteralPath (Join-Path $ResolvedCodexHome "AGENTS.md") -Raw) -match "hdocx-agent-install:start")
      } else {
        $false
      }
    }
    claude = [ordered]@{
      home = $ResolvedClaudeHome
      skill = Test-Skill $claudeSkill
    }
  }

  Ensure-Directory $ResolvedToolHome
  $reportPath = Join-Path $ResolvedToolHome "install-report.json"
  if ($DryRun) {
    Write-Step "Would write install report: $reportPath"
    Write-Host ($report | ConvertTo-Json -Depth 8)
  } else {
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Step "Wrote install report: $reportPath"
  }
}

if (-not ($All -or $Tool -or $Codex -or $Claude -or $CodexFallbackAgent)) {
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

if ($CodexFallbackAgent) {
  Update-CodexFallbackAgent $CodexHome $ToolHome
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

Write-InstallReport $ToolHome $CodexHome $ClaudeHome

Write-Step "Done."
if ($Codex) {
  Write-Step "Restart Codex so it discovers newly installed skills."
}
if ($Claude) {
  Write-Step "Restart Claude Code if it was already running before this install."
}

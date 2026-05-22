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

function Write-Text($Path, $Value, $Encoding = "UTF8") {
  if ($DryRun) {
    Write-Step "Would write file: $Path"
    return
  }
  Set-Content -LiteralPath $Path -Value $Value -Encoding $Encoding
}

function To-TomlString($Value) {
  $escaped = $Value.Replace("\", "\\").Replace('"', '\"')
  return '"' + $escaped + '"'
}

function Install-Tool($RepoRoot, $ResolvedToolHome) {
  $venvPath = Join-Path $ResolvedToolHome "venv"
  $binPath = Join-Path $ResolvedToolHome "bin"
  $tmpPath = Join-Path $ResolvedToolHome "tmp"
  $cachePath = Join-Path $ResolvedToolHome "pip-cache"
  $pythonExe = Join-Path $venvPath "Scripts\python.exe"
  $htmlDocxExe = Join-Path $venvPath "Scripts\html-docx.exe"
  $mcpExe = Join-Path $venvPath "Scripts\html-docx-mcp.exe"

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
  }

  $htmlCmd = Join-Path $binPath "html-docx.cmd"
  $htmlPs1 = Join-Path $binPath "html-docx.ps1"
  $mcpCmd = Join-Path $binPath "html-docx-mcp.cmd"
  $mcpPs1 = Join-Path $binPath "html-docx-mcp.ps1"

  Write-Text $htmlCmd "@echo off`r`n""%~dp0..\venv\Scripts\html-docx.exe"" %*`r`n" "ASCII"
  Write-Text $mcpCmd "@echo off`r`n""%~dp0..\venv\Scripts\html-docx-mcp.exe"" %*`r`n" "ASCII"
  Write-Text $htmlPs1 @'
$exe = Join-Path $PSScriptRoot "..\venv\Scripts\html-docx.exe"
& $exe @args
exit $LASTEXITCODE
'@
  Write-Text $mcpPs1 @'
$exe = Join-Path $PSScriptRoot "..\venv\Scripts\html-docx-mcp.exe"
& $exe @args
exit $LASTEXITCODE
'@

  Write-Step "CLI shim directory: $binPath"
  Write-Step "MCP command: $mcpCmd"

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

  return @{
    binPath = $binPath
    htmlDocxCmd = $htmlCmd
    mcpCmd = $mcpCmd
    htmlDocxExe = $htmlDocxExe
    mcpExe = $mcpExe
  }
}

function Update-CodexMcpConfig($ResolvedCodexHome, $McpCommand) {
  Ensure-Directory $ResolvedCodexHome
  $configPath = Join-Path $ResolvedCodexHome "config.toml"
  $startMarker = "# hdocx-mcp-install:start"
  $endMarker = "# hdocx-mcp-install:end"
  $block = @"
$startMarker
[mcp_servers.hdocx]
command = $(To-TomlString $McpCommand)
args = []
$endMarker
"@

  if (Test-Path -LiteralPath $configPath) {
    $existing = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop
  } else {
    $existing = ""
  }

  $hasManualHdocx = ($existing -match "(?m)^\[mcp_servers\.hdocx\]") -and ($existing -notmatch [regex]::Escape($startMarker))
  if ($hasManualHdocx -and (-not $Force)) {
    throw "Codex config already contains [mcp_servers.hdocx]. Re-run with -Force to replace the managed block manually."
  }

  $pattern = "(?s)\r?\n?$([regex]::Escape($startMarker)).*?$([regex]::Escape($endMarker))\r?\n?"
  $clean = [regex]::Replace($existing, $pattern, "`r`n")
  $newContent = ($clean.TrimEnd() + "`r`n`r`n" + $block.TrimEnd() + "`r`n")

  if ($DryRun) {
    Write-Step "Would update Codex MCP config: $configPath"
    Write-Host $block
    return
  }

  if ((Test-Path -LiteralPath $configPath) -and ($existing -notmatch [regex]::Escape($startMarker))) {
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backup = "$configPath.hdocx-backup-$stamp"
    Copy-Item -LiteralPath $configPath -Destination $backup -Force
    Write-Step "Backed up Codex config: $backup"
  }

  Set-Content -LiteralPath $configPath -Value $newContent -Encoding UTF8
  Write-Step "Updated Codex MCP config: $configPath"
}

function Configure-ClaudeMcp($McpCommand) {
  $claude = Get-Command claude -ErrorAction SilentlyContinue
  $manual = "claude mcp add --transport stdio --scope user hdocx -- `"$McpCommand`""

  if (-not $claude) {
    Write-Step "Claude Code CLI was not found. Run this manually after installing Claude Code:"
    Write-Host $manual
    return
  }

  if ($DryRun) {
    Write-Step "Would configure Claude Code MCP:"
    Write-Host $manual
    return
  }

  & $claude.Source mcp add --transport stdio --scope user hdocx -- $McpCommand
  if ($LASTEXITCODE -ne 0) {
    throw "claude mcp add failed with exit code $LASTEXITCODE. You can run manually: $manual"
  }
  Write-Step "Configured Claude Code MCP server: hdocx"
}

function Write-InstallReport($ResolvedToolHome, $ResolvedCodexHome, $ToolInfo) {
  $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
  $pathContainsBin = $false
  if ($userPath) {
    $pathContainsBin = @(($userPath -split ";") | Where-Object {
      $_.TrimEnd("\") -ieq $ToolInfo.binPath.TrimEnd("\")
    }).Count -gt 0
  }

  $doctor = @{
    ok = $false
    exitCode = $null
    output = $null
  }
  if ((Test-Path -LiteralPath $ToolInfo.htmlDocxCmd) -and (-not $DryRun)) {
    $doctorOutput = & $ToolInfo.htmlDocxCmd doctor 2>&1
    $doctor.exitCode = $LASTEXITCODE
    $doctor.ok = ($LASTEXITCODE -eq 0)
    $doctor.output = ($doctorOutput | Out-String).Trim()
  }

  $codexConfig = Join-Path $ResolvedCodexHome "config.toml"
  $report = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    toolHome = $ResolvedToolHome
    cli = [ordered]@{
      htmlDocx = $ToolInfo.htmlDocxCmd
      htmlDocxMcp = $ToolInfo.mcpCmd
      userPathContainsBin = $pathContainsBin
      doctor = $doctor
    }
    mcp = [ordered]@{
      command = $ToolInfo.mcpCmd
      codexConfig = $codexConfig
      codexConfigContainsHdocx = if (Test-Path -LiteralPath $codexConfig) {
        ((Get-Content -LiteralPath $codexConfig -Raw) -match "hdocx-mcp-install:start")
      } else {
        $false
      }
      claudeAddCommand = "claude mcp add --transport stdio --scope user hdocx -- `"$($ToolInfo.mcpCmd)`""
    }
  }

  Ensure-Directory $ResolvedToolHome
  $reportPath = Join-Path $ResolvedToolHome "mcp-install-report.json"
  if ($DryRun) {
    Write-Step "Would write install report: $reportPath"
    Write-Host ($report | ConvertTo-Json -Depth 8)
  } else {
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Step "Wrote install report: $reportPath"
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

$ToolHome = [System.IO.Path]::GetFullPath($ToolHome)
$CodexHome = [System.IO.Path]::GetFullPath($CodexHome)

Write-Step "Repository: $repoRoot"

$toolInfo = @{
  binPath = Join-Path $ToolHome "bin"
  htmlDocxCmd = Join-Path $ToolHome "bin\html-docx.cmd"
  mcpCmd = Join-Path $ToolHome "bin\html-docx-mcp.cmd"
}

if ($Tool) {
  $toolInfo = Install-Tool $repoRoot $ToolHome
}

if ($Codex) {
  Update-CodexMcpConfig $CodexHome $toolInfo.mcpCmd
}

if ($Claude) {
  Configure-ClaudeMcp $toolInfo.mcpCmd
}

Write-InstallReport $ToolHome $CodexHome $toolInfo

Write-Step "Done."
if ($Codex) {
  Write-Step "Restart Codex so it loads the hdocx MCP server."
}
if ($Claude) {
  Write-Step "Restart Claude Code, or run /mcp inside Claude Code to verify the hdocx server."
}

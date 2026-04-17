param(
    [string]$SkillName = 'office-desktop-assistant'
)

$ErrorActionPreference = 'Stop'

if ($env:CODEX_HOME) {
    $skillRoot = Join-Path $env:CODEX_HOME 'skills'
} else {
    $skillRoot = Join-Path $HOME '.codex\skills'
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $repoRoot "skills\$SkillName"
$target = Join-Path $skillRoot $SkillName

if (-not (Test-Path -LiteralPath $source)) {
    throw "Skill source not found: $source"
}

New-Item -ItemType Directory -Force -Path $skillRoot | Out-Null
Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -LiteralPath $source -Destination $target -Recurse -Force

Write-Output "Installed $SkillName to $target"

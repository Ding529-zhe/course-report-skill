<#
.SYNOPSIS
    把 WorkBuddy 里的 skill 同步进本仓库，并提交推送。

.DESCRIPTION
    日常只在 ~/.workbuddy/skills/<skill> 里改技能，改完跑这个脚本，
    仓库里的 skills/<skill> 会被镜像成一致的内容，然后提交并推送。
    推上去之后 CC Switch 点「刷新」就能看到更新。

.PARAMETER SkillName
    技能目录名，默认 course-report-docx。必须与 SKILL.md 里的 name 一致。

.PARAMETER Source
    源目录，默认 %USERPROFILE%\.workbuddy\skills\<SkillName>。

.PARAMETER Message
    提交信息。留空则自动生成带时间戳的默认信息。

.PARAMETER NoPush
    只提交不推送。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/sync-from-workbuddy.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/sync-from-workbuddy.ps1 -Message "fix: 修正题注字号"
#>
[CmdletBinding()]
param(
    [string]$SkillName = 'course-report-docx',
    [string]$Source = '',
    [string]$Message = '',
    [switch]$NoPush
)

$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
if ($Source -eq '') {
    $Source = Join-Path $env:USERPROFILE ".workbuddy\skills\$SkillName"
}
$dest = Join-Path $repo "skills\$SkillName"

if (-not (Test-Path -LiteralPath $Source)) {
    throw "源目录不存在：$Source"
}
if (-not (Test-Path -LiteralPath (Join-Path $repo '.git'))) {
    throw "当前目录不是 git 仓库：$repo"
}

Write-Output "源  ：$Source"
Write-Output "目标：$dest"
Write-Output ''

# 镜像复制。robocopy 退出码 0-7 都算成功，>=8 才是真错误。
robocopy $Source $dest /MIR /NFL /NDL /NJH /NJS /NP /XD .git | Out-Null
if ($LASTEXITCODE -ge 8) {
    throw "robocopy 失败，退出码 $LASTEXITCODE"
}

$status = git -C $repo status --porcelain
if (-not $status) {
    Write-Output '内容没有变化，无需提交。'
    exit 0
}

Write-Output '变更文件：'
$status | ForEach-Object { Write-Output "  $_" }
Write-Output ''

if ($Message -eq '') {
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $Message = "chore: 同步 $SkillName（$stamp）"
}

git -C $repo add -A
git -C $repo commit -m $Message | Out-Null
Write-Output "已提交：$Message"

if ($NoPush) {
    Write-Output '按 -NoPush 要求，未推送。'
    exit 0
}

git -C $repo push
Write-Output ''
Write-Output '已推送。CC Switch 点「刷新」即可看到更新。'

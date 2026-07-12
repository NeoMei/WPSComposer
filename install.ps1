# 将调试好的 WPSComposer 插件安装到 Codex skills 目录
# 用法: pwsh ./install.ps1   或   pwsh ./install.ps1 -Force
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = $PSScriptRoot
$SkillName   = 'WPSComposer'
$DestRoot    = Join-Path $env:USERPROFILE '.codex\skills'
$Dest        = Join-Path $DestRoot $SkillName

if (Test-Path $Dest) {
    if (-not $Force) {
        Write-Host "目标已存在: $Dest" -ForegroundColor Yellow
        Write-Host "使用 -Force 覆盖安装，或先手动删除。" -ForegroundColor Yellow
        return
    }
    Remove-Item $Dest -Recurse -Force
    Write-Host "已删除旧版本: $Dest" -ForegroundColor DarkGray
}

# 拷贝 .codex-plugin 和 skills 子目录（排除 __pycache__）
New-Item -ItemType Directory -Path $Dest -Force | Out-Null
Copy-Item (Join-Path $ProjectRoot '.codex-plugin') (Join-Path $Dest '.codex-plugin') -Recurse -Force
Copy-Item (Join-Path $ProjectRoot 'skills')        (Join-Path $Dest 'skills')        -Recurse -Force

# 清理目标里的 __pycache__
Get-ChildItem -Path $Dest -Recurse -Directory -Filter '__pycache__' |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "已安装到: $Dest" -ForegroundColor Green
Write-Host "插件文件清单:" -ForegroundColor DarkGray
Get-ChildItem -Path $Dest -Recurse -File |
    ForEach-Object { Write-Host "  $($_.FullName.Substring($Dest.Length + 1))" -ForegroundColor DarkGray }

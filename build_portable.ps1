$ErrorActionPreference = "Stop"
$sourceRoot = $PSScriptRoot
$projectRoot = Split-Path -Parent $sourceRoot
$python = Join-Path $sourceRoot ".venv_build311\Scripts\python.exe"
$buildRoot = Join-Path $sourceRoot ".build"
$mcpDistRoot = Join-Path $buildRoot "mcp-dist"
$guiDistRoot = Join-Path $buildRoot "gui-dist"
$stagingRoot = Join-Path $buildRoot "portable-dist"
$bundleRoot = Join-Path $stagingRoot "DnS Auto"
$portableBundle = Join-Path $sourceRoot "DnS Auto"

if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw "빌드 Python을 찾을 수 없습니다. 먼저 setup_build_environment.ps1을 실행하세요: $python"
}
New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null

foreach ($path in @($mcpDistRoot, $guiDistRoot, $stagingRoot)) {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
}

function Merge-Directory {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    foreach ($item in Get-ChildItem -LiteralPath $Source -Force) {
        $target = Join-Path $Destination $item.Name
        if ($item.PSIsContainer) {
            Merge-Directory -Source $item.FullName -Destination $target
        } else {
            Copy-Item -LiteralPath $item.FullName -Destination $target -Force
        }
    }
}

& $python -m PyInstaller --noconfirm --clean --distpath $mcpDistRoot --workpath (Join-Path $buildRoot "mcp-work") (Join-Path $sourceRoot "DnS_Auto_MCP.spec")
if ($LASTEXITCODE -ne 0) { throw "MCP PyInstaller 빌드 실패 (종료 코드: $LASTEXITCODE)" }

& $python -m PyInstaller --noconfirm --clean --distpath $guiDistRoot --workpath (Join-Path $buildRoot "gui-work") (Join-Path $sourceRoot "DnS_Auto.spec")
if ($LASTEXITCODE -ne 0) { throw "GUI PyInstaller 빌드 실패 (종료 코드: $LASTEXITCODE)" }
$mcpBundle = Join-Path $mcpDistRoot "DnS Auto"
$guiBundle = Join-Path $guiDistRoot "DnS Auto"
if (-not (Test-Path -LiteralPath $mcpBundle -PathType Container)) { throw "MCP one-dir 산출 폴더가 없습니다: $mcpBundle" }
if (-not (Test-Path -LiteralPath $guiBundle -PathType Container)) { throw "GUI one-dir 산출 폴더가 없습니다: $guiBundle" }
Merge-Directory -Source $mcpBundle -Destination $bundleRoot
Merge-Directory -Source $guiBundle -Destination $bundleRoot

Copy-Item -LiteralPath (Join-Path $sourceRoot "mcp_policy.json") -Destination $bundleRoot -Force
New-Item -ItemType Directory -Force -Path (Join-Path $bundleRoot "inputs"), (Join-Path $bundleRoot "outputs"), (Join-Path $bundleRoot "profiles\sheet"), (Join-Path $bundleRoot "profiles\pdf") | Out-Null
foreach ($name in @(
    "DnS_Auto_MCP_사용자_설명서(AI_연동용).html",
    "DnS_Auto_사용자_설명서(직접_실행용).html",
    "DnS_Auto_빠른_시작_가이드(직접_실행용).html",
    "DnS_Auto_통합_설명서.html",
    "DnS_Auto_통합_설명서(빠른_시작).html",
    "CODEX_MCP_CONFIG.example.json",
    "CODEX_MCP_CONFIG.example.toml"
)) {
    Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination $bundleRoot -Force
}
$internalRoot = Join-Path $bundleRoot "_internal"
New-Item -ItemType Directory -Force -Path $internalRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $sourceRoot "assets") -Destination $internalRoot -Recurse -Force

# 병합된 staging의 복제본에서 GUI와 MCP를 검증한 뒤에만 기존 포터블을 교체합니다.
& $python -B (Join-Path $sourceRoot "tests\portable_bundle_smoke.py") --bundle $bundleRoot
if ($LASTEXITCODE -ne 0) { throw "통합 포터블 스모크 테스트 실패 (종료 코드: $LASTEXITCODE)" }

if (Test-Path -LiteralPath $portableBundle) { Remove-Item -LiteralPath $portableBundle -Recurse -Force }
Copy-Item -LiteralPath $bundleRoot -Destination $sourceRoot -Recurse -Force
Write-Host "통합 포터블 빌드 및 검증 완료: $portableBundle"

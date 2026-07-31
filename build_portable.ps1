$ErrorActionPreference = "Stop"
$sourceRoot = $PSScriptRoot
$projectRoot = Split-Path -Parent $sourceRoot
$python = Join-Path $sourceRoot ".venv_build311\Scripts\python.exe"
$buildRoot = Join-Path $sourceRoot ".build"
$stagingRoot = Join-Path $buildRoot "portable-dist"
$guiDistRoot = Join-Path $buildRoot "gui-dist"
$bundleRoot = Join-Path $stagingRoot "DnS Auto"
$releaseRoot = Join-Path $projectRoot "release"
$releaseBundle = Join-Path $releaseRoot "DnS Auto"
$archive = Join-Path $releaseRoot "DnS_Auto_Portable.zip"

if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw "빌드 Python을 찾을 수 없습니다. 먼저 setup_build_environment.ps1을 실행하세요: $python"
}
New-Item -ItemType Directory -Force -Path $buildRoot, $releaseRoot | Out-Null

& $python -m PyInstaller --noconfirm --clean --distpath $stagingRoot --workpath (Join-Path $buildRoot "mcp-work") (Join-Path $sourceRoot "DnS_Auto_MCP.spec")
if ($LASTEXITCODE -ne 0) { throw "MCP PyInstaller 빌드 실패 (종료 코드: $LASTEXITCODE)" }

& $python -m PyInstaller --noconfirm --clean --distpath $guiDistRoot --workpath (Join-Path $buildRoot "gui-work") (Join-Path $sourceRoot "DnS_Auto.spec")
if ($LASTEXITCODE -ne 0) { throw "GUI PyInstaller 빌드 실패 (종료 코드: $LASTEXITCODE)" }
Copy-Item -LiteralPath (Join-Path $guiDistRoot "DnS Auto.exe") -Destination $bundleRoot -Force

Copy-Item -LiteralPath (Join-Path $sourceRoot "mcp_policy.json") -Destination $bundleRoot -Force
New-Item -ItemType Directory -Force -Path (Join-Path $bundleRoot "inputs"), (Join-Path $bundleRoot "outputs"), (Join-Path $bundleRoot "profiles\sheet"), (Join-Path $bundleRoot "profiles\pdf") | Out-Null
Copy-Item -LiteralPath (Join-Path $sourceRoot "PORTABLE_README.txt") -Destination (Join-Path $bundleRoot "QUICK_START.txt") -Force
foreach ($name in @("USER_GUIDE.html", "GUI_GUIDE.html", "GUI_QUICK_START.html", "MCP_GUIDE.html", "CODEX_MCP_CONFIG.example.json", "CODEX_MCP_CONFIG.example.toml")) {
    Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination $bundleRoot -Force
}
Copy-Item -LiteralPath (Join-Path $sourceRoot "assets") -Destination $bundleRoot -Recurse -Force

# 완성된 staging만 release에 반영합니다.
if (Test-Path -LiteralPath $releaseBundle) { Remove-Item -LiteralPath $releaseBundle -Recurse -Force }
Copy-Item -LiteralPath $bundleRoot -Destination $releaseRoot -Recurse -Force
if (Test-Path -LiteralPath $archive) { Remove-Item -LiteralPath $archive -Force }
Compress-Archive -LiteralPath $releaseBundle -DestinationPath $archive -CompressionLevel Optimal
Write-Host "통합 포터블 빌드 완료: $releaseBundle"
Write-Host "압축 파일: $archive"
$ErrorActionPreference = "Stop"
$sourceRoot = $PSScriptRoot
$venv = Join-Path $sourceRoot ".venv_build311"
if (-not (Test-Path -LiteralPath $venv)) {
    py -3.11 -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw "Python 3.11 가상환경 생성 실패" }
}
$python = Join-Path $venv "Scripts\python.exe"
& $python -m pip install --upgrade pip
& $python -m pip install --no-cache-dir -r (Join-Path $sourceRoot "requirements.txt") -r (Join-Path $sourceRoot "requirements-build.txt")
if ($LASTEXITCODE -ne 0) { throw "빌드 의존성 설치 실패" }
Write-Host "빌드 환경 준비 완료: $venv"
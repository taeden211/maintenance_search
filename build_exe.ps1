$ErrorActionPreference = 'Stop'

$python = ".\.venv\Scripts\python.exe"

if (-not (Test-Path $python)) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -m venv .venv
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv .venv
    }
    else {
        throw "Python을 찾을 수 없습니다. Python을 설치한 뒤 다시 실행하세요."
    }

    if ($LASTEXITCODE -ne 0) {
        throw "가상환경 생성에 실패했습니다."
    }
}

& $python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw "pip 업그레이드에 실패했습니다."
}

& $python -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    throw "필수 패키지 설치에 실패했습니다."
}

& $python -m PyInstaller --noconfirm --clean --onefile --windowed --name "유지보수사례검색기" app.py
if ($LASTEXITCODE -ne 0) {
    throw "실행 파일 생성에 실패했습니다."
}

Write-Host "완료: dist\유지보수사례검색기.exe"

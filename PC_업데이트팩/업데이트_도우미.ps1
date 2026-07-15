
param()

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Pause-End {
    Write-Host ""
    Read-Host "Enter 키를 누르면 창이 닫힙니다"
}

try {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $payloadDir = Join-Path $scriptDir "payload"

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " 太炅 Lotto Lab 업데이트 도우미 v1.0" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "이 도우미는 최신 Windows/Android/.github 파일을"
    Write-Host "컴퓨터의 GitHub 저장소 폴더에 덮어씁니다."
    Write-Host ""

    if (!(Test-Path $payloadDir)) {
        throw "payload 폴더를 찾을 수 없습니다."
    }

    $repoPath = Read-Host "GitHub 저장소가 있는 컴퓨터 폴더 경로를 입력하세요"
    $repoPath = $repoPath.Trim('"').Trim()

    if ([string]::IsNullOrWhiteSpace($repoPath)) {
        throw "저장소 폴더 경로가 입력되지 않았습니다."
    }

    if (!(Test-Path $repoPath)) {
        $create = Read-Host "폴더가 없습니다. 새로 만들까요? (Y/N)"
        if ($create -match '^[Yy]$') {
            New-Item -ItemType Directory -Path $repoPath -Force | Out-Null
        } else {
            throw "업데이트를 취소했습니다."
        }
    }

    Write-Host ""
    Write-Host "파일 복사 중..." -ForegroundColor Cyan

    $items = @("Windows", "Android", ".github", "README.md", "v6.0_업로드방법.txt")
    foreach ($item in $items) {
        $src = Join-Path $payloadDir $item
        if (!(Test-Path $src)) { continue }

        $dst = Join-Path $repoPath $item

        if (Test-Path $src -PathType Container) {
            if (Test-Path $dst) {
                Remove-Item $dst -Recurse -Force
            }
            Copy-Item $src $dst -Recurse -Force
        } else {
            Copy-Item $src $dst -Force
        }
    }

    Write-Host "파일 복사 완료" -ForegroundColor Green

    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($git -and (Test-Path (Join-Path $repoPath ".git"))) {
        Write-Host ""
        $push = Read-Host "GitHub에 자동 Commit/Push까지 할까요? (Y/N)"
        if ($push -match '^[Yy]$') {
            Push-Location $repoPath
            try {
                git add .
                $message = "Update Taegyeong Lotto Lab Ultimate v6.0"
                git commit -m $message
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "변경사항이 없거나 Commit이 필요하지 않을 수 있습니다." -ForegroundColor Yellow
                }
                git push
                if ($LASTEXITCODE -ne 0) {
                    throw "git push 실패. GitHub 로그인을 확인하세요."
                }
                Write-Host "GitHub 업로드 완료" -ForegroundColor Green
            }
            finally {
                Pop-Location
            }
        } else {
            Write-Host ""
            Write-Host "파일 준비만 완료했습니다." -ForegroundColor Green
            Write-Host "GitHub 웹에서 Add file → Upload files로 올리면 됩니다."
        }
    } else {
        Write-Host ""
        Write-Host "Git 또는 로컬 저장소(.git)가 없어 자동 Push는 생략했습니다." -ForegroundColor Yellow
        Write-Host "파일 준비는 완료되었습니다."
        Write-Host "GitHub 웹에서 Add file → Upload files로 올리면 됩니다."
    }

    Write-Host ""
    Write-Host "완료되었습니다." -ForegroundColor Green
    Pause-End
}
catch {
    Write-Host ""
    Write-Host "오류: $($_.Exception.Message)" -ForegroundColor Red
    Pause-End
    exit 1
}

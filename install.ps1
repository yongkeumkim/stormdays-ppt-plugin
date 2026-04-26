# install.ps1 — 스톰데이즈 PPT 플러그인 설치

$PLUGIN_DIR = $PSScriptRoot
$SKILLS_DIR = "$env:USERPROFILE\.claude\skills"

Write-Host "=== 스톰데이즈 PPT 플러그인 설치 ===" -ForegroundColor Cyan
Write-Host "플러그인 경로: $PLUGIN_DIR"

# 스킬 디렉토리 생성
if (-not (Test-Path $SKILLS_DIR)) {
    New-Item -ItemType Directory -Path $SKILLS_DIR -Force | Out-Null
    Write-Host "  ✅ 스킬 디렉토리 생성: $SKILLS_DIR"
}

# 스킬 파일 내 PLUGIN_DIR 경로 치환 후 복사
$skillContent = Get-Content "$PLUGIN_DIR\skills\generate-ppt.md" -Raw -Encoding UTF8
$pluginDirEscaped = $PLUGIN_DIR -replace '\\', '/'
$skillContent = $skillContent -replace '<PLUGIN_DIR>', $pluginDirEscaped
$skillContent | Set-Content "$SKILLS_DIR\generate-ppt.md" -Encoding UTF8

Write-Host "  ✅ 스킬 설치: $SKILLS_DIR\generate-ppt.md"

# Python 의존성 설치
Write-Host "`n의존성 설치 중..." -ForegroundColor Yellow
python -m pip install python-pptx openpyxl lxml -q
Write-Host "  ✅ Python 패키지 설치 완료"

# 설치 확인
Write-Host "`n=== 설치 확인 ===" -ForegroundColor Green
python -c "from pptx import Presentation; import openpyxl; print('  ✅ 모든 패키지 정상')"

Write-Host @"

=== 사용 방법 ===
1. Claude Code에서 /generate-ppt 명령 실행
2. Excel 파일 경로를 제공하면 자동으로 PPT 생성

Excel 형식은 /generate-ppt 실행 시 안내됩니다.
샘플 Excel: $PLUGIN_DIR\sample\sample_input.xlsx

"@ -ForegroundColor Cyan

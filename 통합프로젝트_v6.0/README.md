# 太炅 Lotto Lab Ultimate v6.0

Windows와 Android를 한 저장소에서 관리하는 통합 프로젝트입니다.

## 구조
- `Windows/` : Windows EXE
- `Android/` : Android APK
- `.github/workflows/windows-build.yml` : EXE 자동 빌드
- `.github/workflows/android-build.yml` : APK 자동 빌드

## 자동 빌드
GitHub에 업로드 후 `Actions` 탭에서 아래 두 작업이 자동 실행됩니다.

- Windows EXE 자동 빌드
- Android APK 자동 빌드

성공하면 각 실행 화면의 `Artifacts`에서 다운로드합니다.

## 주요 기능
- 역대 로또 Excel 분석
- 사진 OCR 및 직접 번호 입력
- 추천조합
- 나온횟수
- 동반수
- 트리플
- 최근패턴
- 자체추천
- 자동 100조합
- 역대 1등·2등 동일 조합 제외

> 분석 결과는 과거 통계 기반이며 당첨을 보장하지 않습니다.

# OSP 재고 하루 1회 자동 동기화 대시보드

Google Sheets의 `월재고현황`을 기준으로 이카운트 N001/N004 재고와 품고/네이버 재고를 매일 정오에 동기화하는 Python 자동화 프로젝트입니다.

대상 시트: https://docs.google.com/spreadsheets/d/1l_2A9NEX8e1Oii0njdXafTWwILY5MKi9_Is1biZUkEM/edit

## 설치 방법

```bash
cd inventory-dashboard
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Playwright 로그인 세션 저장 기능을 사용할 때만 추가로 실행합니다.

```bash
playwright install
```

## .env 설정 방법

`.env.example`을 `.env`로 복사한 뒤 값을 채웁니다.

```env
GOOGLE_APPLICATION_CREDENTIALS=./credentials/google-service-account.json
GOOGLE_SHEET_ID=1l_2A9NEX8e1Oii0njdXafTWwILY5MKi9_Is1biZUkEM
GOOGLE_SHEET_NAME=월재고현황
APPS_SCRIPT_WEBAPP_URL=
APPS_SCRIPT_TOKEN=osp-sync-642768
SYNC_INTERVAL_MINUTES=1440
BACKUP_POLICY=daily
BACKUP_BEFORE_EACH_SYNC=false
HEADLESS=false
SESSION_DIR=./data/session
```

## Apps Script 방식

`APPS_SCRIPT_WEBAPP_URL`이 설정되어 있으면 `python src/sync_inventory.py`는 서비스 계정 JSON 대신 Apps Script 웹 앱으로 전체 대시보드 데이터를 전송합니다.

## Google Service Account 설정 방법

1. Google Cloud에서 Service Account를 만듭니다.
2. JSON 키를 발급합니다.
3. 파일을 `credentials/google-service-account.json`으로 저장합니다.
4. 이 파일은 `.gitignore`에 포함되어 Git에 올라가지 않습니다.

## Google Sheets 공유 권한 설정 방법

Google Sheets 문서에서 `공유`를 누른 뒤, Service Account JSON 안의 `client_email` 주소를 편집자 권한으로 초대합니다.

## 이카운트 로그인 정보 설정 방법

`.env`에 아래 값을 입력합니다.

```env
ECOUNT_ID=
ECOUNT_PASSWORD=
ECOUNT_COMPANY_CODE=
```

기본 수집 방식은 저장된 로그인 세션을 사용한 자동 export입니다. 자동화는 `재고 I > 출력물 > 재고현황` 화면을 열고, N001 `제품자재창고`와 N004 `하은물류`를 각각 선택한 뒤 Excel 파일을 `data/raw/ecount`에 저장합니다.

## 품고 로그인 정보 설정 방법

`.env`에 아래 값을 입력합니다.

```env
POOMGO_ID=
POOMGO_PASSWORD=
```

품고에서 네이버 재고 또는 출고 export를 내려받아 `data/raw/poomgo`에 저장합니다.

## 수동 실행 방법

```bash
python src/sync_inventory.py
```

실행하면 먼저 `월재고현황`을 백업하고, 필요한 경우 네이버재고 열을 추가한 뒤 원본 탭과 Dashboard를 갱신합니다.

더블클릭으로 실행하려면 프로젝트 폴더의 `run_sync.bat`을 사용합니다.
export 다운로드만 먼저 갱신하려면 `refresh_exports.bat`, 준비 상태만 확인하려면 `check_setup.bat`을 사용합니다.

## 자동 실행 설정 방법

자동 실행은 윈도우 작업 스케줄러 사용을 권장합니다. 프로젝트 내부의 `src/scheduler.py`도 제공하지만, 운영 환경에서는 작업 스케줄러가 더 안정적입니다.

## 윈도우 작업 스케줄러 등록 방법

- 실행 주기: 평일 하루 1회, 오전 9:15
- 프로그램/스크립트: `C:\Users\CHOIIH\Documents\New project\inventory-dashboard\.venv\Scripts\python.exe`
- 인수 추가: `src/sync_inventory.py`
- 시작 위치: `C:\Users\CHOIIH\Documents\New project\inventory-dashboard`

프로젝트 폴더에서 아래 명령으로 작업 스케줄러를 바로 등록할 수도 있습니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\install_hourly_task.ps1
```

등록을 해제하려면 아래 명령을 사용합니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall_hourly_task.ps1
```

## 오류 발생 시 확인할 로그 위치

로그는 날짜별로 저장됩니다.

```text
logs/sync_YYYYMMDD.log
```

로그에는 실행 시작/종료, 수집 성공/실패, SKU 수, 업데이트 수, 미매핑 수, 오류 메시지가 기록됩니다.

## 매핑 실패 시 처리 방법

매핑 우선순위는 아래와 같습니다.

1. Google Sheets의 `상품매핑` 시트
2. `config/product_mapping.csv`
3. Google Sheets의 `SKU마스터` 바코드/품목코드 매핑
4. 코드가 완전히 동일한 경우 자동 매칭

매핑되지 않은 외부 상품은 `미매핑상품` 시트에 기록됩니다. 확인 후 `상품매핑` 시트나 `config/product_mapping.csv`에 `sheet_code, ecount_code, poomgo_code, naver_code`를 추가하세요.

## 자동 export와 수동 fallback

API나 브라우저 자동화가 막히면 export 파일만 바꿔서 계속 운영할 수 있습니다.

```env
ECOUNT_EXPORT_DIR=./data/raw/ecount
POOMGO_EXPORT_DIR=./data/raw/poomgo
ECOUNT_USE_PLAYWRIGHT=false
POOMGO_USE_PLAYWRIGHT=false
AUTO_DOWNLOAD_EXPORTS=true
DOWNLOAD_TIMEOUT_SECONDS=120
```

`AUTO_DOWNLOAD_EXPORTS=true`이면 `python src/sync_inventory.py` 실행 시 저장된 로그인 세션으로 이카운트 N001/N004와 품고/네이버 최신 export 다운로드를 먼저 시도합니다. 다운로드가 실패하면 각 폴더에 있는 최신 CSV/XLSX 파일을 자동 선택해 계속 진행합니다.

저장된 로그인 세션으로 export 다운로드용 브라우저를 열려면 아래 명령을 사용합니다.

```bash
python src/open_export_browser.py ecount
python src/open_export_browser.py poomgo
```

다운로드가 감지되면 각각 `data/raw/ecount`, `data/raw/poomgo`에 자동 저장됩니다.

모든 export 자동 다운로드만 먼저 실행하려면 아래 명령을 사용합니다.

```bash
python src/refresh_exports.py
```

## 준비 상태 점검

```bash
python src/check_setup.py
```

Apps Script URL 또는 Google 인증 파일, 로그인 세션, export 파일 준비 여부를 한 번에 확인합니다.

## Playwright 세션 저장

로그인 차단, 2FA, CAPTCHA가 있으면 아래처럼 켜서 직접 로그인 세션을 저장합니다.

```env
HEADLESS=false
ECOUNT_USE_PLAYWRIGHT=true
POOMGO_USE_PLAYWRIGHT=true
```

첫 실행에서 브라우저가 열리면 직접 로그인한 뒤 `세션 저장` 탭의 버튼을 누릅니다. 자동 감지가 애매하면 `현재 브라우저 세션 강제 저장` 버튼을 누르면 됩니다. 이후 세션은 `data/session/ecount_storage_state.json`, `data/session/poomgo_storage_state.json`에 저장됩니다.

로그인 세션만 먼저 저장하려면 아래 명령을 사용합니다.

```bash
python src/capture_sessions.py ecount
python src/capture_sessions.py poomgo
```

## 테스트

```bash
pytest
```

테스트 범위:

- 코드 기준 병합
- N001/N004 재고 합산
- 낱개에서 박스/팔렛트 변환
- 기준값 누락 처리
- 외부 시스템 미조회 표시
- 미매핑 상품 분리
- 네이버재고를 재고합계에 더하지 않는지 검증

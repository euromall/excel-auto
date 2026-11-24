# ExcelAuto

주문·적하 엑셀 반복작업 자동화. 인보이스 프로그램(InvoiceAuto, `invoice_2610`)과 **별도로** 만들어 테스트한 뒤 합칠 예정.

## 현재 기능 (v0.1)

| 기능 | 대체하는 작업 | 상태 |
|---|---|---|
| 적하 작성 (`ExcelAuto.exe`) | 6~8단계: 샵마인 수령인 정보 → 적하, 품명 조회(주문재고변환 `변환` 시트), 품명·수량 입력, Shipping 파일 만들기 | 2026-11-24 실제 파일로 결과 일치 확인 |
| 출고 체크 매크로 (`vba/ShipScan.bas`) | 14단계: 송장 스캔 → 출고일 기록, 출고현황 시트 | 엑셀 실행 검증 전 |

## 적하 작성 동작

1. 샵마인 다운로드 폴더에서 가장 최근 `발송대기*.xlsx`를 읽음
2. 주문마다 `상품주문번호(AD)`, 없으면 `대표주문번호(AC)`로 `Base.xlsx` → `newPrice` F열(주문번호), 없으면 G열(바코드)을 찾아 I열 품명을 가져옴 (같은 번호가 여러 번이면 맨 아래 행)
3. 샵마인 F~V열 → Shipping C~S열, 품명 → AC, 수량(AI) → AF, 고정값(T~Z, AH, AM)과 수식(AB, AP, AQ) 입력
4. 첫 행 B열에 시작 송장번호(HAWB) 입력 (InvoiceAuto가 이 번호부터 부여)
5. `InvoiceAuto\ShippingInvoice` 폴더에 `Shipping_YYYYMMDD_HHMMSS.xlsx`로 저장
6. **품명을 못 찾은 주문이 하나라도 있으면 파일을 만들지 않고 목록을 보여줌** (InvoiceAuto는 품명 없는 행을 건너뛰므로 누락 방지)

## 받기 / 실행

- GitHub → Actions → build-exe → 최근 실행 → Artifacts `ExcelAuto` 다운로드 → `ExcelAuto.exe`를 원하는 폴더에 두고 실행 (Python 설치 불필요)
- 처음 실행 시 경로 4개 지정 후 `설정 저장` (같은 폴더에 `ExcelAuto.ini` 생성)

## 개인정보

실제 주문 파일(xlsx/xlsm)은 저장소에 올리지 않는다(`.gitignore`). 테스트는 가짜 데이터만 사용.

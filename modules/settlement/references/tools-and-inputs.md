# 실행과 입력

명령은 해당 모듈 폴더(`modules/transfer`, `modules/plan`, `modules/settlement` 중 대상 업무 폴더)를 현재 디렉터리로 하여 실행한다.
RHWP는 필요하지 않다. Python에서 `../../requirements.txt`의 패키지를 준비한다. HWP5 읽기/변환은 `../../scripts/hangul_io.py` 및 `../../scripts/normalize_hangul.py`, HWPX 구조 조사는 `hangul_io.py inspect`를 사용한다. Windows 한글이 있으면 `../../scripts/finalize_hancom.ps1`로 HWPX 재저장·PDF 출력 후 검토한다. 스킬만으로 Python이나 한글이 설치되는 것은 아니다.

`python scripts/settlement_totals.py --ledger 원장.xlsx --inputs 정산입력.json --output 검산.json`

검산 입력 JSON:
- as_of: YYYY-MM-DD 보고 기준일.
- funding_budget, reported_spending: 해당 재원 예산액·정산총괄 집행액(원).
- interest, disallowed: 확인된 이자·불인정액(원). 미확인은 null로 두며 계산기를 실행하면 오류로 알려준다.
- spending_by_category: unit, category, spent를 가진 비용구분별 배열. 확인된 0은 명시하고 누락은 오류로 취급한다.
- approval_events: date, order, document, unit, category, before, after, reason. 승인 이력별 행 배열. 변경후 총액은 각 항목의 직전금액에서 이어져야 하며 결재별 증감은 0이어야 한다. 지원금 총액 자체 변경은 별도로 매핑해야 한다.

여러 재원은 재원별로 분리한다. 현재 계산기는 재원 1종을 검산하며, 모든 학교의 파일을 즉시 제출용 HWPX로 바꾸는 범용 생성기는 아니다.

HWP/HWPX 편집 후 스키마 검사만으로 완료 처리하지 않는다. 한글 등 편집기의 PDF 출력·페이지 검토가 필요하다. 암호·배포용 HWP는 보호 해제된 사본, HWP97·이미지 문서는 별도 변환/OCR과 금액 대조가 필요하다.

기본 원장 시트 보호 암호는 Q1W2E3이다. AI가 의견이나 수식을 갱신해야 하면 해당 사본에서 보호를 해제하고 작업 후 재설정한다. 입력열만 평소 편집 가능하게 유지한다.

최종 출력 규칙: 공단 기본양식의 수치 칸만 채운 HWPX.

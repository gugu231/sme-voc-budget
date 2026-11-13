# RHWP 없는 실행 구성

명령은 해당 모듈 폴더(`modules/transfer`, `modules/plan`, `modules/settlement` 중 대상 업무 폴더)를 현재 디렉터리로 하여 실행한다.
Python 3 환경에서 `python -m pip install -r ../../requirements.txt`로 의존성을 준비한다. 스킬은 실행 파일 자체가 아니므로 Python과 해당 패키지가 필요하다.

1. HWP 내용 추출: `python ../../scripts/hangul_io.py read 참고계획.hwp --output 참고계획.json`
2. HWP5/HWPX 정규화: `python ../../scripts/normalize_hangul.py 참고계획.hwp 참고계획_변환.hwpx`
3. 표 생성: `python scripts/create_transfer_request.py --ledger 원장.xlsx --template assets/전용신청서_공단예시.hwpx --content 문안.json --stage 3 --output 신청서_초안.hwpx` (`--stage` 생략 시 증감이 입력된 마지막 차수)
4. Windows 한글 설치 환경에서: `powershell -File ../../scripts/finalize_hancom.ps1 -InputPath 신청서_초안.hwpx -OutputPath 신청서.hwpx -PdfPath 신청서.pdf`
5. PDF 모든 페이지를 열어 잘림·누락·금액·글꼴을 확인한다. XML 검증 통과만으로 제출 가능 판정을 하지 않는다.

문안 JSON은 `school`(학교명), `year`(연도), `bullets`(총괄·세부 사유 문자열 배열), `reasons`(차수 문자열 → 원장 세부항목명 → 사유 문자열)로 구성한다. 증감이 있는 상세행마다 사유가 필요하다. 같은 항목명이 여러 행에 있으면 `row:5`처럼 원장 행 번호를 키로 지정한다. 행 번호 키는 항목명 키보다 우선한다. 최소 형식은 `examples/content.example.json`을 참조한다. 학교별 실제 항목명에 맞게 채운다.

생성기는 표준 원장 머리글(3행)을 엄격히 확인하며 1차~신청 차수 표를 출력한다. 각 차수는 프로그램·세부항목·원장 행 번호로 구별해 증감을 적으며 상쇄되는 내부 이동도 보존한다. 표에는 어느 차수에서든 증감이 있는 비용구분만 공단 양식 순서로 싣는다. 단위내전용 별도 열은 지원하지 않으므로 해당 차수 증감 열에 합산해 둔다. 과거 표는 최신 원장에서 재구성하므로 승인본과 같다는 뜻이 아니며, 별도 승인 이력과 대조한다.

## 여러 학교의 HWP 대응

파일 확장자 대신 실제 OLE/ZIP 구조로 HWP5/HWPX를 분류한다. 원본은 보존하고 변환 사본만 만든다. HWP5는 python-hwpx로 HWPX 변환하고 원문 텍스트·표·금액을 대조한 뒤 공단 양식에 필요한 내용을 옮긴다. 학교마다 임의로 만든 원장 열을 자동 추정하지 말고 명시적 매핑을 먼저 작성한다.

암호·배포용 HWP는 보호 해제된 사본이 필요하다. HWP97·이미지 문서는 별도 변환/OCR 경로가 필요하며, OCR 금액은 원본 대조 후 사용한다. 복잡한 도형·수식·특수 표의 외형까지 모든 HWP에서 동일하게 보존된다고 보장하지 않는다. 변환 오류·본문 누락은 실패로 보고하고 미리보기 텍스트를 전체 본문으로 취급하지 않는다.

RHWP 없이 문서 생성·변환은 가능하다. Windows 한글 COM으로 최종 레이아웃을 검증할 수 있다. 한글이 없는 환경에서는 호환 편집기에서 열기·PDF 출력·페이지 확인 단계가 별도로 필요하며, 이 패키지에 독립적인 렌더러는 포함하지 않았다.

- python-hwpx: https://github.com/airmang/python-hwpx
- 한컴 형식 공개: https://www2.hancom.com/support/downloadCenter/hwpOwpml

본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

차수 열 이름은 승인 여부의 증거가 아니다. 기준일별 정산에서는 별도의 승인 이력을 사용한다.

HWP/HWPX는 RHWP로 우선 렌더링하여 모든 페이지의 표 잘림·겹침·공란·금액 줄바꿈을 확인한다. 별도 RHWP 스킬 없이 CLI/MCP를 사용하며 준비·실행 절차는 스킬 루트 `scripts/RHWP_RENDERING.md`를 따른다. 구조 검사만으로 시각 검증을 대체하지 않는다. 실행할 수 없으면 구체적 사유와 미검증 범위를 알린다. 최종 제출 화면은 작업자가 확인한다.
공식 도구 안내: https://github.com/edwardkim/rhwp . 실행 가능하면 `rhwp capabilities`로 해당 버전의 명령을 확인한 뒤 렌더링한다.

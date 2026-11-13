# RHWP 화면 검증

별도 RHWP 스킬 없이 RHWP CLI 또는 MCP로 HWP/HWPX를 렌더링한다.

1. 사용 가능한 RHWP 도구를 확인한다. CLI에서는 `rhwp capabilities`와 `rhwp export-svg --help`로 설치 버전의 기능을 확인한다.
2. 도구가 없으면 [공식 Releases](https://github.com/edwardkim/rhwp/releases/latest)의 운영체제별 실행 파일을 작업용 도구 폴더에 준비하고 배포된 SHA256SUMS로 확인한다. 관리자 설치나 전역 설정 변경은 필요하지 않다. 실행 정책 차단을 해제하거나 우회하지 않는다.
3. 작업 사본에서 `rhwp export-svg "작성본.hwpx" -o "검토폴더"`를 실행한다. PDF가 지원되면 `rhwp export-pdf "작성본.hwpx" -o "검토.pdf"`를 사용할 수 있다. MCP에서는 노출된 도구 목록과 입력 스키마를 확인해 대응 렌더링 기능을 사용한다.
4. 생성된 모든 페이지를 실제 이미지·뷰어로 열어 표 잘림, 겹침, 빈 표·빈 쪽, 금액 줄바꿈과 글꼴 누락을 확인한다. 생성 성공이나 텍스트 추출만으로 화면 검증 완료로 표시하지 않는다. 문제를 수정하면 다시 렌더링한다.
5. 결과에는 사용한 RHWP 버전, 확인한 페이지 범위, 남은 문제를 간략히 기록한다. 실행이 불가능하면 이유를 명시하고 시각 검증 미완료로 반환한다. RHWP 결과가 한컴 한글 출력과 완전히 같다고 보장하지 않으며 최종 제출본은 작업자가 확인한다.

RHWP는 외부 실행 도구이며 이 스킬에 바이너리를 포함하지 않는다. 별도 스킬 미설치는 렌더링 생략 사유가 아니다.

공식 명령·MCP 안내: https://github.com/edwardkim/rhwp/blob/main/README_EN.md

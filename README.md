# sme-voc-budget

중소기업 특성화고 인력양성사업의 예산 원장을 만들고 검증한 뒤 사업계획서·전용신청서·정산 서식을 작성하는 통합 스킬입니다.

FAQ 171문항을 18개 묶음으로 구성해 스킬 본체 102개 파일, 저장소 전체 104개 파일로 배포합니다. 질문별 근거와 검색 위치는 유지합니다.

## 다운로드

**Code → Download ZIP**으로 내려받아 한 번만 압축을 풉니다. `sme-voc-budget-main` 폴더의 이름을 `sme-voc-budget`으로 바꿉니다. 폴더 바로 안에 `SKILL.md`, `manuals`, `modules`, `scripts`, `requirements.txt`가 있어야 합니다. 개별 파일만 옮기지 말고 폴더 전체를 사용합니다.

## 1. GPT·Codex에서 설치

이 항목은 로컬 파일과 스크립트를 다룰 수 있는 **Codex** 기준입니다. 일반 ChatGPT 채팅에 파일을 첨부하는 것과 스킬을 설치하는 것은 구분합니다.

### Codex 설치 도우미 사용

Codex의 `$skill-installer`에서 이 저장소 주소와 스킬 경로 `.`(저장소 루트)를 지정합니다.

- 저장소: https://github.com/gugu231/sme-voc-budget
- 스킬 경로: `.`
- 스킬 이름: `sme-voc-budget`

### 직접 설치

1. 압축을 푼 `sme-voc-budget` 폴더 전체를 사용 환경의 스킬 디렉터리에 복사합니다.
2. 현재 공식 문서의 사용자 공통 경로는 `~/.agents/skills/sme-voc-budget/`입니다. Windows에서는 사용자 폴더 아래 `.agents\skills\sme-voc-budget\`입니다.
3. `.codex/skills`를 사용하는 데스크톱 배포판은 그 환경에서 지정한 스킬 경로에 설치합니다. 동일 스킬을 여러 경로에 중복 설치하지 않습니다.
4. 새 대화에서 `$sme-voc-budget`을 선택하거나 입력합니다. 표시되지 않으면 Codex를 재시작합니다.

일반 ChatGPT에서 이 폴더를 첨부하는 것만으로 영구 설치되거나 한글 COM이 실행된다고 보장하지 않습니다. 해당 환경의 스킬·플러그인 지원과 파일 실행 도구를 확인합니다.

근거: [OpenAI 공식 스킬 안내](https://learn.chatgpt.com/docs/build-skills), [OpenAI Codex 스킬 설치 경로 예시](https://developers.openai.com/cookbook/examples/codex/iterating-development-workflows-with-codex).

## 2. Claude에서 설치

### Claude 웹·Desktop·Cowork의 계정 스킬

1. 다운로드한 폴더 이름을 `sme-voc-budget`으로 맞춥니다.
2. 이 폴더를 ZIP으로 압축합니다. ZIP 내부는 `sme-voc-budget/SKILL.md` 구조여야 하며, 다른 ZIP 파일을 넣지 않습니다. 이 압축은 Claude 업로드용 포장입니다.
3. **Customize → Skills → + → Create skill → Upload a skill**에서 ZIP을 올리고 활성화합니다.
4. 새 대화에 예산자료를 첨부하고 `sme-voc-budget` 스킬 사용을 지정합니다.
5. 업로드 메뉴가 없으면 조직의 사용자 스킬 허용 여부와 코드 실행·파일 생성 설정을 확인합니다.

웹·클라우드 실행 환경에서는 로컬 Windows 한글 COM을 직접 실행할 수 없습니다. 생성된 HWPX는 한글 등 호환 편집기에서 열어 최종 배치를 확인합니다.

근거: [Claude 공식 설치 안내](https://support.claude.com/en/articles/12512180-use-skills-in-claude).

### Claude Code의 로컬 스킬

1. 압축을 푼 폴더 전체를 `~/.claude/skills/sme-voc-budget/`에 복사합니다. 프로젝트 한정 설치는 `.claude/skills/sme-voc-budget/`입니다.
2. Claude Code에서 `/sme-voc-budget`으로 호출합니다.
3. 파일 접근 권한과 실행 도구를 확인한 뒤 학교 예산자료를 처리합니다. Cowork는 이 로컬 폴더를 직접 읽는 방식이 아니므로 위 계정 스킬 방식으로 설치합니다.

근거: [Claude Code 공식 스킬 안내](https://code.claude.com/docs/en/skills).

## 준비할 원자료

Google 스프레드시트 링크보다 **학교에서 실제 사용하는 자체 예산편성 엑셀(.xlsx) 파일 업로드를 권장**합니다. 대상 시트·기준일·반영 차수를 함께 지정합니다. Google 자료도 처리할 수 있으나 변환된 수식·누락 탭·계산값을 추가 확인해야 합니다.

원장의 사용 중인 행은 **목·세목·사업단위·비용구분·세부항목·산출근거·본예산을 모두 근거에 따라 입력한 뒤** 다음 양식을 작성합니다. 빈 예비행은 채우지 않습니다. 모르는 값을 추정해 채우거나 미확인을 0원으로 바꾸지 않습니다.

본예산과 적용 증감의 합이 **0원인 항목의 최종 산출근거는 `0원`**으로 씁니다. 양수 예산의 산출근거를 맞추려고 임의의 `금액 × 1식`을 만들지 않습니다.

## 작업 순서와 산출물

학교 예산자료 → 표준 원장 작성 → 검산·오류 수정 → 작업자 원장 검토 → 필요한 서식 작성 → 작업자 최종 검증.

| 업무 | 산출물 |
|---|---|
| 예산 원장 작성·검증 | XLSX, 확인사항 |
| 사업계획서 예산 | HWPX, 확인사항 |
| 전용신청서 | HWPX, 확인사항 — 표준 원장 1~4차 생성 지원 |
| 예산 및 전용편성내역 입력자료 | XLSX, 복붙위치, 확인사항 |
| 중간정산보고서 | 기본양식의 대상 수치를 채운 HWPX, 확인사항 |

중간정산에는 실제 집행자료와 기준일별 승인 이력이 별도로 필요합니다. 예산을 집행액으로 대신하지 않습니다.

## 공통 실행 조건

- Python 스크립트를 실행하는 환경은 `requirements.txt`의 패키지가 필요합니다. 스킬 설치만으로 Python이나 한글이 설치되지는 않습니다.
- RHWP는 필수가 아닙니다. Windows 한글 COM 보조 스크립트는 한글이 설치된 Windows에서 실행합니다.
- HWPX 구조 검사를 통과해도 표 잘림·글꼴·쪽 배치는 편집기에서 별도로 확인합니다.
- 매뉴얼은 관련 주제별 추출문이며 원본 PDF는 포함하지 않습니다. 표·이미지·해석이 불명확하면 해당 연도 공식 원문과 대조합니다.
- 생성 결과는 작업자가 원자료·승인 문서·집행 증빙과 최종 대조합니다. 자동 검산은 최종 제출 승인을 대신하지 않습니다.

## 업데이트

현재 버전: **2026.11.14.1**. GitHub 변경은 이미 설치한 ZIP에 자동 반영되지 않습니다. 스킬은 네트워크 접근이 가능한 작업 시작 시 배포본 버전을 확인하도록 구성돼 있습니다. 이 확인은 환경과 에이전트의 지침 실행 여부에 의존하며, 강제 실행되는 업데이트 프로그램은 아닙니다.

- **GPT/Codex 로컬 폴더 설치:** 변경 시 최신 파일로 설치 폴더를 갱신합니다. Git으로 설치한 경우 작업 시작 전에 해당 설치 저장소에서 `git pull --ff-only`로 갱신할 수 있습니다. 로컬 변경이 있으면 덮어쓰지 말고 먼저 보존합니다. 변경을 인식하지 못하면 새 대화나 앱 재시작을 사용합니다.
- **Claude 개인 ZIP 설치:** GitHub 변경과 별개이므로 최신 ZIP으로 설치본을 갱신합니다. 기존 버전이 함께 활성화되지 않게 확인합니다.
- **Claude Team/Enterprise 공유:** 관리자가 공유 스킬을 갱신하면 수신자는 다음 사용 시 갱신본을 받는 방식을 활용할 수 있습니다. 조직 게시 스킬은 조직의 재게시·승인 절차를 따릅니다. GitHub 변경만으로 공유본이 바뀌지는 않습니다.

[OpenAI 스킬 문서](https://learn.chatgpt.com/docs/build-skills) · [Claude 설치·공유 안내](https://support.claude.com/en/articles/12512180-use-skills-in-claude)

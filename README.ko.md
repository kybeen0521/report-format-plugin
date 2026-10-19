[English](README.md) | 한국어

# report-format

A4 한글 기술 보고서를 일정한 형식의 Word(.docx)로 만드는 Claude Code 스킬이다. PDF 가 필요하면 Word 에서 "PDF 로 저장"한다.

## 설치

### 방법 1. Claude Code 터미널(대화창)에서

```
/plugin marketplace add kybeen0521/report-format-plugin
/plugin install report-format@report-format
```

### 방법 2. 셸 명령으로 (VS Code 확장 등 `/plugin` 을 쓸 수 없는 환경)

VS Code 확장의 채팅창에서는 `/plugin` 이 "isn't available in this environment" 로 막힌다. 아무 터미널에서 아래를 실행한다.

```
claude plugin marketplace add kybeen0521/report-format-plugin
claude plugin install report-format@report-format
```

설치 후 Claude Code 세션(VS Code 창 포함)을 새로 열면 스킬이 잡힌다.

### 갱신·삭제

```
claude plugin marketplace update report-format
claude plugin update report-format@report-format
claude plugin uninstall report-format@report-format
```

### 처음 한 번: 환경 준비

python-docx 를 설치하고, Word 가 없는 PC 를 위한 근사 미리보기 도구(typst·한글 폰트)를 준비한다. Claude 에게 "보고서 환경 준비해줘"라고 하거나 직접:

```
python3 <설치 위치>/skills/report-format/scripts/report_kit.py setup
python3 <설치 위치>/skills/report-format/scripts/report_kit.py doctor --build
```

## 쓰기

"이 결과를 보고서로 정리해줘"처럼 요청하면 스킬이 로드되어 원고(markdown)를 쓰고 `.docx` 를 만든 뒤 미리보기로 검수한다.

## 필요한 것

- Python 3
- pip 패키지 `python-docx`(setup 이 설치)
- 선택: `matplotlib`(그림 스크립트), `Pillow`(그림 여백 자르기·미리보기), `PyMuPDF`(미리보기), `typst`(근사 미리보기)
- Word 미리보기: Microsoft Word(Windows, 또는 WSL 에서 Windows 의 Word) 또는 LibreOffice. 없으면 typst 근사 미리보기

## 구성

```
.claude-plugin/marketplace.json   마켓플레이스 정의
.claude-plugin/plugin.json        플러그인 정의
skills/report-format/
  SKILL.md                        절차·내용 범위 규칙·검수 체크리스트
  references/rules.md             형식(F 부)·작성(W 부) 규칙
  scripts/report_kit.py           빌더(markdown → docx, 미리보기, 점검)
  examples/                       예제 원고와 그림 스크립트
```

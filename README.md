# report-format

A4 한글 기술 보고서를 일정한 형식의 Word(.docx)로 만드는 Claude Code 스킬이다. 같은 원고로 PDF 도 만든다.

## 설치

Claude Code 에서:

```
/plugin marketplace add kybeen0521/report-format-plugin
/plugin install report-format@report-format
```

로컬 폴더에서 설치할 때는 `/plugin marketplace add /경로/report-format-plugin`.

설치 후 처음 한 번 환경을 준비한다(python-docx·typst 설치, PDF 용 한글 폰트 내려받기). Claude 에게 "보고서 환경 준비해줘"라고 하거나 직접:

```
python3 <설치 위치>/skills/report-format/scripts/report_kit.py setup
python3 <설치 위치>/skills/report-format/scripts/report_kit.py doctor --build
```

## 쓰기

"이 결과를 보고서로 정리해줘"처럼 요청하면 스킬이 로드되어 원고(markdown)를 쓰고 `.docx` 를 만든 뒤 미리보기로 검수한다. PDF 가 필요하면 "PDF 로" 라고 말한다.

## 필요한 것

- Python 3
- pip 패키지 `python-docx`, `typst`(setup 이 설치)
- 선택: `matplotlib`(그림 스크립트), `Pillow`(그림 여백 자르기·미리보기), `PyMuPDF`(미리보기)
- Word 미리보기: Microsoft Word(Windows, 또는 WSL 에서 Windows 의 Word) 또는 LibreOffice. 없으면 typst 근사 미리보기

## 구성

```
.claude-plugin/marketplace.json   마켓플레이스 정의
.claude-plugin/plugin.json        플러그인 정의
skills/report-format/
  SKILL.md                        절차·내용 범위 규칙·검수 체크리스트
  references/rules.md             형식(F 부)·작성(W 부) 규칙
  scripts/report_kit.py           빌더(markdown → docx / pdf, 미리보기, 점검)
  examples/                       예제 원고와 그림 스크립트
```

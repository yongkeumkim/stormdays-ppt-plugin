# stormdays-ppt-plugin

Claude Code 플러그인 — Excel 파일을 입력받아 스톰데이즈 디자인 템플릿으로 PowerPoint를 자동 생성합니다.

## 특징

- **픽셀 퍼펙트**: 배경, 폰트, 색상, 레이아웃을 템플릿과 완전히 동일하게 유지
- **멀티 에이전트**: Content Analyst → Layout Planner → PPT Builder 3단계 AI 팀
- **9가지 슬라이드 타입**: cover, toc, stat, two_col, steps, metrics, chart, table, closing
- **Excel 입력**: 시트 구조에 맞게 데이터만 채우면 자동 생성

## 설치

```powershell
# PowerShell에서 실행
.\install.ps1
```

설치 후 `~/.claude/skills/generate-ppt.md` 가 생성됩니다.

## 사용법

Claude Code에서:

```
/generate-ppt C:\path\to\your\data.xlsx
```

또는 자연어로:

```
이 엑셀 파일로 PPT 만들어줘: C:\data\business_plan.xlsx
```

## Excel 입력 형식

### 필수: `config` 시트

| key | value |
|-----|-------|
| title | 프레젠테이션 제목 |
| subtitle | 부제목 |
| date | 2025.10.24 |
| company | 회사명 |
| year | 2025 |

### 선택 시트

| 시트명 | 슬라이드 타입 | 설명 |
|--------|-------------|------|
| `stat` | 큰 숫자 강조 | 핵심 지표 1개 + 설명 bullet |
| `two_col` | 두 컬럼 | 좌우 비교 텍스트 |
| `steps` | 3단계 카드 | 프로세스/전략 순서 |
| `metrics` | 성과 지표 | KPI 4개 바/진행률 |
| `chart` | 파이 차트 | 비율/구성 데이터 |
| `table` | 표 | 일정, 항목 비교 |
| `closing` | 결론 | 마무리 메시지 |

샘플 Excel: `sample/sample_input.xlsx`

## 파일 구조

```
stormdays-ppt-plugin/
├── install.ps1              # 설치 스크립트
├── skills/
│   └── generate-ppt.md      # Claude Code 스킬 정의
├── scripts/
│   ├── analyze_excel.py     # Excel → content.json
│   └── build_pptx.py        # slide_plan.json → .pptx
├── agents/
│   ├── content-analyst.md   # Content Analyst 역할 정의
│   ├── layout-planner.md    # Layout Planner 역할 정의
│   └── ppt-builder.md       # PPT Builder 역할 정의
├── template/
│   └── stormdays.pptx       # 스톰데이즈 디자인 템플릿
└── sample/
    └── sample_input.xlsx    # 샘플 입력 Excel
```

## 요구사항

- Python 3.8+
- python-pptx
- openpyxl
- lxml
- Claude Code CLI

## 에이전트 팀 구조

```
사용자 요청 (Excel 파일)
      │
      ▼
오케스트레이터 (/generate-ppt 스킬)
      │
      ├── Phase 1: Content Analyst
      │           Excel 파싱 → content.json
      │
      ├── Phase 2: Layout Planner
      │           맥락 분석 → slide_plan.json
      │
      ├── Phase 3: PPT Builder
      │           slide_plan.json → output.pptx
      │
      └── Phase 4: QA
                  PNG 변환 → 시각 검증
```

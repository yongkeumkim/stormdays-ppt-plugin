# generate-ppt — 스톰데이즈 PPT 자동 생성 스킬

## 트리거
사용자가 `/generate-ppt <excel_파일>` 을 실행하거나
"엑셀 파일로 PPT 만들어줘", "PPT 자동 생성" 등을 요청할 때

## 플러그인 경로 설정
이 스킬이 실행되는 시점에 PLUGIN_DIR을 결정한다:

```bash
# Windows
PLUGIN_DIR="C:/Temp/환경구성/stormdays-ppt-plugin"
# (설치 후 변경되면 install.ps1 실행 결과로 업데이트)
```

---

## 에이전트 팀 역할

```
사용자 요청 (Excel 파일)
      │
      ▼
┌─────────────────────────────────────────┐
│  오케스트레이터 (이 스킬)                  │
│  - Excel 경로 수신 및 작업 폴더 생성       │
│  - 3개 서브 에이전트 순차 조율             │
│  - 최종 QA 및 결과 보고                   │
└─────────────────────────────────────────┘
      │
      ├─ Phase 1 ──▶ [Content Analyst]
      │               Excel 파싱 → content.json
      │
      ├─ Phase 2 ──▶ [Layout Planner]
      │               content.json → slide_plan.json
      │               (맥락 기반 슬라이드 배치 결정)
      │
      ├─ Phase 3 ──▶ [PPT Builder]
      │               slide_plan.json → output.pptx
      │
      └─ Phase 4 ──▶ QA (이미지 변환 + 시각 확인)
```

---

## 실행 절차

### 0. 사전 확인
```bash
python -c "from pptx import Presentation; print('OK')" 2>/dev/null || pip install python-pptx -q
python -c "import openpyxl; print('OK')" 2>/dev/null || pip install openpyxl -q
```

### 1. 작업 폴더 생성
```bash
WORK_DIR="$(dirname '<EXCEL_PATH>')/ppt_work_$(date +%Y%m%d_%H%M%S)"
mkdir -p "$WORK_DIR"
echo "작업 폴더: $WORK_DIR"
```

### 2. Phase 1 — Content Analyst (서브 에이전트 위임)

다음 프롬프트로 서브 에이전트를 스폰한다:

```
너는 Content Analyst야. Excel 파일을 파싱해서 content.json을 생성해야 해.

실행 명령:
python "<PLUGIN_DIR>/scripts/analyze_excel.py" "<EXCEL_PATH>" "<WORK_DIR>/content.json"

완료 후 content.json 경로를 보고해줘.
에러가 나면 오류 메시지와 함께 보고해줘.

참고 파일: <PLUGIN_DIR>/agents/content-analyst.md
```

### 3. Phase 2 — Layout Planner (서브 에이전트 위임)

content.json을 읽어 슬라이드 배치 계획을 수립하는 에이전트:

```
너는 Layout Planner야. content.json을 분석해서 최적의 슬라이드 배치를 결정해야 해.

입력 파일: <WORK_DIR>/content.json

작업:
1. content.json을 읽어라
2. 각 데이터의 성격을 파악해라:
   - 숫자 하나를 강조할 데이터 → "stat" 타입
   - 두 가지 비교 데이터 → "two_col" 타입
   - 3단계 프로세스 → "steps" 타입
   - KPI/성과 지표 → "metrics" 타입
   - 비율/구성 차트 → "chart" 타입
   - 일정/표 데이터 → "table" 타입
   - 비전/미션 선언 → "vision" 타입
   - 마지막은 항상 → "closing" 타입
3. 슬라이드 흐름이 자연스럽도록 순서를 결정해라
4. TOC에 들어갈 항목 목록을 만들어라
5. 아래 형식으로 slide_plan.json을 생성해라

출력 형식 (slide_plan.json):
{
  "config": { ... content.json의 config 그대로 ... },
  "slide_plan": [
    { "order": 1, "type": "cover", "section_num": "", "title": "", "data": { "config에서 복사" } },
    { "order": 2, "type": "toc", "section_num": "", "title": "목차", "items": [{"num":"01","text":"제목"},...] },
    { "order": 3, "type": "stat", "section_num": "01", "title": "제목", "data": { ... } },
    ...
    { "order": N, "type": "closing", "section_num": "0X", "title": "결론", "data": { ... } }
  ]
}

출력 파일: <WORK_DIR>/slide_plan.json

참고 파일: <PLUGIN_DIR>/agents/layout-planner.md
```

### 4. Phase 3 — PPT Builder (서브 에이전트 위임)

```
너는 PPT Builder야. slide_plan.json으로 PPTX를 생성해야 해.

실행 명령:
python "<PLUGIN_DIR>/scripts/build_pptx.py" \
  "<WORK_DIR>/slide_plan.json" \
  "<WORK_DIR>/output.pptx" \
  --template "<PLUGIN_DIR>/template/stormdays.pptx"

생성 후 슬라이드 수와 각 슬라이드 제목을 확인해:
python -c "
from pptx import Presentation
prs = Presentation('<WORK_DIR>/output.pptx')
print(f'총 {len(prs.slides)}개 슬라이드')
for i, s in enumerate(prs.slides):
    texts = [sh.text_frame.text[:40] for sh in s.shapes if sh.has_text_frame]
    print(f'  Slide {i+1}: {texts[:2]}')
"

오류 없이 완료되면 성공을 보고해줘.
에러가 나면 traceback 전체를 보고해줘.

참고 파일: <PLUGIN_DIR>/agents/ppt-builder.md
```

### 5. Phase 4 — QA (이미지 변환 + 시각 검증)

```bash
python - << 'EOF'
import comtypes.client, os
output_pptx = "<WORK_DIR>/output.pptx"
qa_dir = "<WORK_DIR>/qa"
os.makedirs(qa_dir, exist_ok=True)

ppt = comtypes.client.CreateObject("PowerPoint.Application")
ppt.Visible = True
prs = ppt.Presentations.Open(output_pptx, ReadOnly=True, WithWindow=False)
for i in range(1, prs.Slides.Count + 1):
    prs.Slides(i).Export(f"{qa_dir}/slide_{i:02d}.png", "PNG", 1920, 1080)
    print(f"  QA: slide_{i:02d}.png")
prs.Close()
ppt.Quit()
print("QA 이미지 생성 완료")
EOF
```

QA 이미지를 순서대로 읽어 시각 검사:
- 배경 디자인(노란 바, 점 패턴, 삼각형)이 모든 슬라이드에 있는가
- 텍스트가 잘리지 않고 박스 안에 들어가는가
- 섹션 번호가 올바른 순서인가
- 빈 내용 슬라이드가 없는가

---

## 오케스트레이터 최종 보고 형식

```
✅ PPT 생성 완료

📁 출력 파일: <WORK_DIR>/output.pptx
📊 슬라이드 구성:
  - Slide 1: [Cover] 제목
  - Slide 2: [TOC] 목차
  - Slide 3: [01] 현황 분석
  ...

🎨 디자인 검증:
  - 스톰데이즈 배경 템플릿: ✅
  - 폰트 (Noto Sans/Serif KR): ✅
  - 색상 체계 (#FFBD0B 옐로): ✅

⚠️ 주의사항: (있을 경우)
```

---

## Excel 입력 형식 안내

사용자가 Excel 형식을 물어보면 아래 내용을 안내:

### 필수 시트: `config`
| key | value |
|-----|-------|
| title | 프레젠테이션 제목 |
| subtitle | 부제목 |
| date | 2025.10.24 |
| company | 회사명 |
| year | 2025 |

### 선택 시트들 (하나 이상 필요)

**`stat` 시트** (큰 숫자 강조):
section_num, title, big_number, number_label, bullet1, bullet2, bullet3

**`two_col` 시트** (두 컬럼):
section_num, title, left_header, left1~left5, right_header, right1~right5

**`steps` 시트** (단계별 전략):
section_num, title, step1_label, step1_title, step1_items, step2_label, step2_title, step2_items, step3_label, step3_title, step3_items

**`metrics` 시트** (성과 지표):
section_num, title, metric1_label, metric1_value, metric2_label, metric2_value, metric3_label, metric3_value, metric4_label, metric4_value

**`chart` 시트** (파이 차트):
section_num, title, total_label, cat1, val1, pct1, cat2, val2, pct2, cat3, val3, pct3, cat4, val4, pct4

**`table` 시트** (표):
각 행이 테이블의 한 행. 첫 행 = 헤더. section_num과 title은 별도 열로.

**`closing` 시트** (결론):
section_num, title, body (결론 본문 텍스트)

---

## 관련 파일
- 스크립트: `<PLUGIN_DIR>/scripts/`
- 에이전트 역할 정의: `<PLUGIN_DIR>/agents/`
- 템플릿: `<PLUGIN_DIR>/template/stormdays.pptx`

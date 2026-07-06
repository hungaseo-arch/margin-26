# 마진 분석 (margin_26)

월별 **판매 자료**와 **매입(AP Invoice) 자료**를 결합해 업체·상품·아이템·브랜드별 마진을 집계하고,
`Excel` + `PDF` 보고서를 자동 생성하는 도구입니다.

분석 코드는 [margin.ipynb](margin.ipynb) 노트북에 있습니다.

---

## 1. 입력 파일

| 구분 | 파일명 | 시트/헤더 | 사용 컬럼 |
|---|---|---|---|
| **매출자료** | `6.2.1 Delivery Status(List Only).xlsx` | `skiprows=6` | `Delivery Date.1`, `Buyer`, `Brand`, `Item Code`, `Description`, `Type 2`, `Type 3`, `Q'ty`, `Discounted Amount (IDR)` |
| **구매자료(원가)** | `9.2 Inventory_Ledger_For_Costing.xlsx` | `skiprows=4` | `Item Code`, `Unit Cost`(=`Unnamed: 10`) → `P_Price(IDR)` |
| (로고) | `ASCENDO_Blue.png` | — | PDF 헤더 이미지 |

### 구매단가(원가) 산정 — 재고원장 Unit Cost

원가는 **재고원장(9.2)의 `Unit Cost` 컬럼**(ERP가 산정한 품목별 이동평균 단가)을 사용합니다.
`Item Code`가 **품목당 1건**으로 유일하고 깨끗해, 별도 가공 없이 그대로 씁니다.

1. `skiprows=4`로 읽어 `Unnamed: 10`(= Unit Cost) → `P_Price(IDR)`로 rename
2. **Unit Cost가 0/공란이면 원장 내 대체 원가열로 재계산(보완)**:
   `Ending(20÷19)` → `Beginning(4÷3)` → `Good Receipt(6÷5)` → `Landed(7÷5)` 순
3. **그래도 원가가 없는 품목**은 `cost_overrides.csv`(있으면)로 수동 보정 —
   `Item Code,P_Price(IDR),P_Price(USD)` 로 지정(예: 판매 `E4`=구매 `E3` 명칭변경).
   **fill-only**: 원장에 실제 원가가 생기면 자동 무시(수동보정은 금회 한정). *공급단가라 git 제외.*
4. 그래도 없는 품목만 제외 → **미커버 목록**(`uncovered_items.xlsx`)으로 관리

> **참고**: 초기에는 매입세금계산서(`9.1 AP Invoice Status`)를 원가로 썼으나, 한 품목에
> 부대비용(≈15%)·취소전표·0원 라인이 섞여 **마진율이 80~100%로 과대 계산**되는 문제가
> 있었습니다(→ [마진 재검증](#7-마진-재검증-margin--50)). 재고원장 단가로 전환해 해소했습니다.

### 매출 정합성 — 미매칭 품목 처리

판매자료에 `Item Code`로 left merge 하되 **매칭 안 되는 행도 매출에는 포함**합니다
(→ 월 매출합계가 원본 피벗과 **정확히 일치**). 다만 원가가 없으므로 **마진은 N/A**로 두고,
각 표에 **`Coverage(%)`(마진산출 커버율 = 원가매칭 매출 ÷ 전체 매출)** 를 함께 표기합니다.

---

## 2. 출력 파일

**보고 대상 월(단월)** 기준으로 아래 파일이 생성됩니다. `.env`의 `START_MONTH`/`END_MONTH`로 지정하며,
**매월 그 달로 1회 실행**합니다(과거 1~5월은 기존 보고서 사용 → 재생성하지 않음).

| 파일 | 내용 |
|---|---|
| `sales analysis report YYYY-MM.xlsx` / `.pdf` | **IDR 보고서** — 5개 시트 (Row_Data / 업체 / 상품 / 아이템 / 브랜드) |
| `sales analysis report YYYY-MM USD.xlsx` / `.pdf` | **USD 보고서** — 동일 구조. 매출=`Discounted Amount (USD)`, 원가=USD 재고원장 Unit Cost |
| `uncovered_items.xlsx` | **원가 미매칭 품목** 목록 3시트 (월별 요약 / 코드별 합계 / 월별·코드별 상세) — 구매팀 원장 등록 검토용 |

> **IDR·USD 마진율이 다른 이유(환율)**: 매출 USD는 **판매시점 환율**, 원가 USD는 재고원장의
> **매입/원가계상 시점 환율**로 환산됩니다. 두 시점 환율이 다르면(예: 루피아 약세) 마진율이
> 통화별로 달라집니다. 같은 환율이면 두 마진율은 동일 — **환차 효과이지 오류가 아닙니다.**
> (2026-06 예: IDR 14.8% vs USD 12.5%, 매입~판매 사이 루피아 약 2.8% 약세)

**각 보고서에 기록되는 항목**
- **환율(IDR/USD)**: Excel `0.FX Rate` 시트 + PDF 상단에 매출·원가 환율 표기
  (2026-06: 매출 17,898 / 원가 17,415)
- **Coverage(%)**: 원가매칭 커버율 = 원가매칭 매출 ÷ 전체 매출
  - **100% 도달 시** → Coverage 컬럼 **전체 삭제**(모두 100%라 불필요)
  - **100% 미달 시** → 표에 컬럼 유지 + **보고서 상단에 경고 표시**(원가 미매칭 매출 포함)
- 로고는 헤더에 **85% 크기**로 표시

> 두 파일은 **동일한 파일명 형태**를 사용하며, 월은 `06`처럼 **숫자 2자리**로 표시됩니다.
> PDF 본문 헤더의 월 표기(`SALES ANALYSIS for Jun 2026`)만 영문 월 이름을 사용합니다.

---

## 3. uv 관리환경 구축

이 프로젝트는 [uv](https://docs.astral.sh/uv/)로 파이썬 버전과 의존성을 관리합니다.
의존성은 [pyproject.toml](pyproject.toml)에 정의되어 있고, 정확한 버전은 `uv.lock`에 고정됩니다.

### 3-1. uv 설치 (최초 1회)

이미 설치되어 있으면 건너뜁니다. (`uv --version`으로 확인)

```powershell
# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 3-2. 환경 구축 (프로젝트 폴더에서)

```powershell
cd "g:\내 드라이브\00-Working\_margin_26"

# pyproject.toml + uv.lock 기준으로 .venv 생성 및 의존성 설치
uv sync
```

- `uv sync`가 Python 3.14를 자동으로 내려받고 `.venv`를 만든 뒤 모든 패키지를 설치합니다.
- 이후 패키지를 추가하려면 `uv add <패키지명>` (예: `uv add matplotlib`).

### 3-3. 의존성 목록 ([pyproject.toml](pyproject.toml))

| 패키지 | 용도 |
|---|---|
| `pandas` | 데이터 가공·집계 |
| `openpyxl` | Excel 읽기/쓰기 |
| `fpdf2` | PDF 생성 |
| `pillow` | PDF 로고 이미지 처리 (fpdf2 의존) |
| `ipykernel` | VS Code / Jupyter 노트북 커널 |
| `jupyter`, `nbconvert` | 노트북 실행 (CLI `--execute` 포함) |

---

## 4. 실행 방법

### ① 여러 달 일괄 생성 — `generate_report.py` (권장)

`.env`의 월 범위대로 각 달의 Excel·PDF를 한 번에 생성합니다.

```powershell
uv run python generate_report.py
```

```
[done] sales analysis report 2026-01.xlsx / .pdf  (rows=1560)
...
[done] sales analysis report 2026-06.xlsx / .pdf  (rows=1279)
완료: 모든 월 보고서 생성
```

### ② VS Code 노트북 (단일 기간 인터랙티브 분석)

1. VS Code에서 이 폴더 열기
2. `margin.ipynb` 열기 → 우측 상단 **커널 선택** → **`.venv` (Python 3.14)** 선택
3. **Run All** 또는 셀마다 `Shift+Enter`

> 코드 수정 후에는 **커널 재시작(Restart Kernel)** 후 재실행하세요 (이전 정의 캐시 방지).

---

## 5. 보고 기간 변경

### 스크립트: `.env` 수정 (매월 그 달로)

```dotenv
REPORT_YEAR=2026
START_MONTH=6
END_MONTH=6      # 해당 월 단월. 매월 그 달로 지정
```

> 과거 1~5월은 각 월의 원장으로 이미 생성된 기존 보고서를 사용하며 **재생성하지 않습니다**
> (원장이 6월 스냅샷이라 과거월에 적용하면 원가·마진이 왜곡됨 → [§6](#6-6월-검증-기준-보고서-대비) 참고).
> `.env`는 git 미커밋. 최초에는 `Copy-Item .env.example .env`.

### 노트북: 두 번째 셀 수정

```python
YEAR, MONTH = 2026, 6
```

현재 데이터는 **2026-01 ~ 2026-06** 범위가 존재합니다.

---

## 6. 6월 검증 (기준 보고서 대비)

이 도구는 **6월(현행 월)** 을 생성하며, 6월 재고원장으로 계산합니다.
결과를 기존 `sales analysis report 2026-06`(기준)과 대조:

| 항목 | 기준 보고서 | 본 도구 | 비고 |
|---|---|---|---|
| 매출(Sales) | 29,955,065,588 | 29,955,065,588 | **원본 피벗과 완전 일치** |
| 마진율 | 15.1% | **14.8%** | 기준이 과대(아래) |
| Coverage | (표기없음) | **100.0%** | 원가 보정 후 전량 매칭 |

**0.3%p 차이의 원인 = 단일 품목 `OTR1-180025E44059`** (ASC 18.00-25 E4 40PR 59):
- 재고원장에 원가가 **0**(미계상)이라, 기준 보고서는 이를 **원가 0 → 마진 100%**(순이익 1.1억)로 잡아 총마진을 15.1%로 부풀림.
- 본 도구는 `cost_overrides.csv`에 **실매입가 22,509,000**(= 동일규격 구매품 `E3`, 9.1 AP Invoice 2026-04-21)을 지정 →
  해당 품목 마진 **19.1%**(정상), 6월 총마진 **14.8%**(정확).

> 결론: 6월 매출·마진 모두 정상이며, 본 도구의 **14.8%가 정확**합니다(기준 15.1%는 미계상 원가를 100% 마진으로 오처리한 값).

---

## 7. 마진 재검증 (Margin > 50%)

`9.1 AP Invoice`를 원가로 쓰던 시기에 마진율 50%를 넘는 이상치가 다수 발생했습니다.
원인은 **AP Invoice 데이터 구조**였습니다(코드 오류 아님):

- 한 `Item Code`에 정상 매입가(예 2,713,000)와 **부대비용 라인(≈405,000)** 이 공존
- **취소 전표**(음수)와 **0원 라인**이 최신 건으로 잡히면 원가가 비정상적으로 낮아짐
  → 마진율 85~100%로 과대 계산

**재고원장(9.2)의 Unit Cost**로 전환해 근본 해소했습니다(품목당 단가 유일·정제):

| 원가 소스 | 전체 마진 | 마진>50% 행 |
|---|---|---|
| 9.1 AP Invoice · 최신1건 | 18.7% | 147 |
| 9.1 AP Invoice · 수량가중평균 | 11.2% | 15 |
| **9.2 재고원장 Unit Cost (적용)** | 8.6~14.8% | **월 0~3건** |

> 이상치가 사라졌을 뿐 아니라 원가 매칭 커버율도 87~99.6%로 향상되었습니다.

---

## 8. 트러블슈팅

| 증상 | 원인 / 해결 |
|---|---|
| `FileNotFoundError: 'ASCENDO_Blue.png'` | 로고 파일이 폴더에 없음 → 프로젝트 폴더에 함께 두기 |
| `FPDFUnicodeEncodingException ... "times"` | 맑은 고딕 폰트 등록 누락 → `pdf.add_font('Malgun', ...)` 확인 |
| `PermissionError: ... .pdf/.xlsx` | 출력 파일이 뷰어에서 열려 있음 → 닫고 재실행 |
| `IndexError ... iloc[0]` | 날짜 필터 결과 0건 → `start_date`/`end_date`를 데이터가 있는 범위로 조정 |
| PDF 한글 폰트 오류 (macOS/Linux) | `C:\Windows\Fonts\malgun.ttf` 경로 없음 → OS에 맞는 한글 폰트 경로로 수정 |

---

**Made by**: Jonghwan SEO

# 마진 분석 (margin_26)

월별 **판매 자료**와 **매입(AP Invoice) 자료**를 결합해 업체·상품·아이템·브랜드별 마진을 집계하고,
`Excel` + `PDF` 보고서를 자동 생성하는 도구입니다.

분석 코드는 [margin.ipynb](margin.ipynb) 노트북에 있습니다.

---

## 1. 입력 파일

| 구분 | 파일명 | 시트/헤더 | 사용 컬럼 |
|---|---|---|---|
| **매출자료** | `6.2.1 Delivery Status(List Only).xlsx` | `skiprows=6` | `Delivery Date.1`, `Buyer`, `Brand`, `Item Code`, `Description`, `Type 2`, `Type 3`, `Q'ty`, `Discounted Amount (IDR)` |
| **구매자료** | `9.1 AP Invoice Status.xlsx` | `skiprows=7` | `Item Code`, `A/P Inv Date`, `Unit Price (IDR)` → `P_Price(IDR)` |
| (로고) | `ASCENDO_Blue.png` | — | PDF 헤더 이미지 |

> **구매단가 처리**: AP Invoice는 `Item Code` 하나에 매입 이력이 여러 건 있으므로,
> `A/P Inv Date` 기준으로 정렬 후 **품목별 가장 최근 매입단가**만 남겨(`drop_duplicates(keep="last")`)
> 판매자료에 `Item Code`로 left merge 합니다. 매칭 안 되는 행은 `dropna()`로 제거됩니다.

---

## 2. 출력 파일

월(月)마다 아래 2개 파일이 생성됩니다 (예: 2026년 1~6월 → 12개):

| 파일 | 내용 |
|---|---|
| `sales analysis report YYYY-MM.xlsx` | 5개 시트 (Row_Data / 업체 / 상품 / 아이템 / 브랜드) |
| `sales analysis report YYYY-MM.pdf`  | A4 세로, 표 4종 (브랜드 → 상품 → 업체 → 아이템) |

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

### 일괄 스크립트: `.env` 수정

```dotenv
REPORT_YEAR=2026
START_MONTH=1
END_MONTH=6
```

> `.env`는 git에 커밋되지 않습니다. 최초에는 `.env.example`을 복사해 만드세요
> (`Copy-Item .env.example .env`).

### 노트북: 두 번째 셀 수정

```python
start_date = '2026-06-01'
end_date   = '2026-06-30'
```

현재 데이터는 **2026-01 ~ 2026-06** 범위가 존재합니다.

---

## 6. 트러블슈팅

| 증상 | 원인 / 해결 |
|---|---|
| `FileNotFoundError: 'ASCENDO_Blue.png'` | 로고 파일이 폴더에 없음 → 프로젝트 폴더에 함께 두기 |
| `FPDFUnicodeEncodingException ... "times"` | 맑은 고딕 폰트 등록 누락 → `pdf.add_font('Malgun', ...)` 확인 |
| `PermissionError: ... .pdf/.xlsx` | 출력 파일이 뷰어에서 열려 있음 → 닫고 재실행 |
| `IndexError ... iloc[0]` | 날짜 필터 결과 0건 → `start_date`/`end_date`를 데이터가 있는 범위로 조정 |
| PDF 한글 폰트 오류 (macOS/Linux) | `C:\Windows\Fonts\malgun.ttf` 경로 없음 → OS에 맞는 한글 폰트 경로로 수정 |

---

**Made by**: Jonghwan SEO

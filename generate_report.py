"""
월별 마진 보고서 일괄 생성기
- 매출자료(6.2.1 Delivery Status) + 구매자료(9.1 AP Invoice Status)를 결합
- 지정한 연도의 여러 개월에 대해 각각 Excel(.xlsx) + PDF 보고서를 생성

[원가 산정 방식]  ※ 마진율 이상치(>50%) 재검증 반영
  - 구매단가 = Item Code별 '수량가중평균' 단가 = Σ Amount(IDR) / Σ Q'ty
  - 취소 전표('A/P Invoice Cancelled')와 단가 0원 라인은 제외
  - '최신 1건' 방식은 부대비용(≈15%)·0원·취소 라인을 잘못 집어 마진율을
    과대(80~100%) 계산하는 문제가 있어 가중평균으로 대체함

[매출 정합성]
  - 매입단가가 없는(미매칭) 품목도 매출에는 포함 → 월 매출합계가 원본 피벗과 일치
  - 원가 없는 행은 마진 계산에서 제외(N/A)하고, 표에 'Coverage(%)'(마진산출 커버율) 표기

출력 파일명: 'sales analysis report YYYY-MM.xlsx' / '... .pdf'  (월은 숫자 2자리)

실행:  uv run python generate_report.py
"""
import calendar
import datetime
import os
from pathlib import Path

import pandas as pd
from fpdf import FPDF


# ─────────────────────────────────────────────────────────────
# .env 로더 (외부 의존성 없이 KEY=VALUE 를 환경변수로 로드)
def load_dotenv(path=".env"):
    p = Path(path)
    if not p.exists():
        return
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


load_dotenv()

# ★ 보고 대상 설정 (.env 로 덮어쓰기 가능, 없으면 아래 기본값 사용)
YEAR = int(os.getenv("REPORT_YEAR", "2026"))
START_MONTH = int(os.getenv("START_MONTH", "1"))
END_MONTH = int(os.getenv("END_MONTH", "6"))
MONTHS = range(START_MONTH, END_MONTH + 1)   # 기본 1월 ~ 6월

SALES_FILE = os.getenv("SALES_FILE", "6.2.1 Delivery Status(List Only).xlsx")
PURCHASE_FILE = os.getenv("PURCHASE_FILE", "9.1 AP Invoice Status.xlsx")
LOGO = os.getenv("LOGO_FILE", "ASCENDO_Blue.png")

# Unicode 폰트 (Windows 기본 - 맑은 고딕)
FONT_REG = os.getenv("FONT_REG", r"C:\Windows\Fonts\malgun.ttf")
FONT_BOLD = os.getenv("FONT_BOLD", r"C:\Windows\Fonts\malgunbd.ttf")
# ─────────────────────────────────────────────────────────────

SALES_COL = "Discounted Amount (IDR)"


# ===== 1. 원본 자료 로딩 (전체 1회) ============================
def load_sales():
    sheets = pd.ExcelFile(SALES_FILE).sheet_names
    sheet = next((s for s in sheets if s.startswith("6.2.1")), sheets[0])
    df = pd.read_excel(SALES_FILE, sheet_name=sheet, skiprows=6)
    df = df.drop("Unnamed: 0", axis=1, errors="ignore")
    df["Delivery Date.1"] = pd.to_datetime(df["Delivery Date.1"])
    return df


def load_cost_table():
    """Item Code별 수량가중평균 구매단가. 취소전표·0원 라인 제외."""
    sheets = pd.ExcelFile(PURCHASE_FILE).sheet_names
    sheet = next((s for s in sheets if s.startswith("9.1")), sheets[0])
    p = pd.read_excel(PURCHASE_FILE, sheet_name=sheet, skiprows=7)
    p = p.drop("Unnamed: 0", axis=1, errors="ignore")

    if "3.AP Invoice Type" in p.columns:
        p = p[p["3.AP Invoice Type"].astype(str).str.strip() == "A/P Invoice"]
    p = p[p["Unit Price (IDR)"].fillna(0) > 0]   # 0원(무상/조정) 라인 제외
    p = p[p["Q'ty"].fillna(0) > 0]               # 수량 없는/음수 라인 제외
    p = p.dropna(subset=["Item Code"])

    g = p.groupby("Item Code").agg(A=("Amount (IDR)", "sum"), Q=("Q'ty", "sum"))
    g["P_Price(IDR)"] = g["A"] / g["Q"]          # 수량가중평균
    return g.reset_index()[["Item Code", "P_Price(IDR)"]]


# ===== 서식 함수 ==============================================
def int_num(x):
    if pd.isna(x):
        return ""
    return "{:,}".format(int(round(x)))


def pct(x):
    if pd.isna(x):
        return "N/A"
    return "{:.1f}%".format(x * 100)


# ===== 집계 헬퍼 =============================================
def summarize(df, keys, total_key, with_qty=False):
    """keys 기준 집계 + TOTAL행 + 비율/커버율. (매출=전체, 마진=원가있는 행만)"""
    aggmap = {
        "Sales": (SALES_COL, "sum"),              # 전체 매출 (정합성)
        "CostedSales": ("CostedSales", "sum"),    # 원가 매칭된 매출 (마진 분모)
        "Margin": ("Margin(IDR)", "sum"),         # 원가 매칭 행의 마진 합 (NaN 자동 skip)
    }
    if with_qty:
        aggmap["Qty"] = ("Q'ty", "sum")
    g = df.groupby(keys).agg(**aggmap)

    g.loc[total_key, :] = g.sum()                 # 합계행
    g["Margin(%)"] = g["Margin"] / g["CostedSales"]
    g["Coverage(%)"] = g["CostedSales"] / g["Sales"]
    g["Sales Ratio(%)"] = g["Sales"] / g.loc[total_key, "Sales"]
    g["Margin Ratio(%)"] = g["Margin"] / g.loc[total_key, "Margin"]
    g = g.sort_values("Sales", ascending=False).reset_index()

    g = g.rename(columns={"Sales": "Sales(IDR)", "Margin": "Margin(IDR)"})
    order = list(keys) if isinstance(keys, list) else [keys]
    if with_qty:
        g = g.rename(columns={"Qty": "Q'ty"})
        order += ["Q'ty"]
    order += ["Sales(IDR)", "Margin(IDR)", "Margin(%)", "Coverage(%)",
              "Sales Ratio(%)", "Margin Ratio(%)"]
    return g[order]


def fmt(g):
    """숫자/퍼센트 서식 적용."""
    if "Q'ty" in g.columns:
        g["Q'ty"] = g["Q'ty"].apply(lambda x: "" if pd.isna(x) else "{:,}".format(int(round(x))))
    g["Sales(IDR)"] = g["Sales(IDR)"].apply(int_num)
    g["Margin(IDR)"] = g["Margin(IDR)"].apply(int_num)
    for c in ["Margin(%)", "Coverage(%)", "Sales Ratio(%)", "Margin Ratio(%)"]:
        g[c] = g[c].apply(pct)
    return g


# ===== 2. 한 달치 집계 =========================================
def build_month(df_s_all, cost, start_date, end_date):
    mask = (df_s_all["Delivery Date.1"] >= start_date) & (df_s_all["Delivery Date.1"] <= end_date)
    df_s = df_s_all.loc[mask]
    if df_s.empty:
        return None

    # 매출 전체 유지(미매칭 포함) → 매출 정합성. 원가 없는 행은 마진 N/A
    df = df_s.merge(cost, how="left", on="Item Code")
    df["Brand"] = df["Brand"].fillna("ETC")
    df["Type 3"] = df["Type 3"].replace("Others", "ETC")
    df["P_Amount(IDR)"] = df["P_Price(IDR)"] * df["Q'ty"]
    df["Margin(IDR)"] = df[SALES_COL] - df["P_Amount(IDR)"]              # 원가 없으면 NaN
    df["CostedSales"] = df[SALES_COL].where(df["P_Price(IDR)"].notna())  # 원가 있는 매출만
    df["Margin(%)"] = df["Margin(IDR)"] / df["CostedSales"]

    df_cust = fmt(summarize(df, "Buyer", "TOTAL"))
    df_prod = fmt(summarize(df, ["Type 3", "Type 2"], ("TOTAL", "ALL"), with_qty=True)
                  .rename(columns={"Type 3": "Product", "Type 2": "Type"}))
    df_item = fmt(summarize(df, "Description", "TOTAL", with_qty=True))
    df_bran = fmt(summarize(df, "Brand", "TOTAL"))
    return df, df_cust, df_prod, df_item, df_bran


# ===== 3. Excel 저장 ==========================================
def save_excel(fname, df, df_cust, df_prod, df_item, df_bran):
    writer = pd.ExcelWriter(fname, engine="openpyxl")
    df.to_excel(writer, sheet_name="1.Row_Data")
    df_cust.to_excel(writer, sheet_name="2.Margin by Customer")
    df_prod.to_excel(writer, sheet_name="3.Margin by Product")
    df_item.to_excel(writer, sheet_name="4.Margin by Item")
    df_bran.to_excel(writer, sheet_name="5.Margin by Brand")
    writer._save()


# ===== 4. PDF 저장 ============================================
def save_pdf(fname, month_name, year, current_date, df_cust, df_prod, df_item, df_bran):
    class PDF(FPDF):
        def header(self):
            self.set_font("Malgun", "B", 20)
            self.set_text_color(32, 32, 32)
            self.set_fill_color(240, 255, 255)
            self.set_line_width(0.4)
            self.cell(0, 20, border=1, align="C", fill=1)
            self.image(LOGO, 15, 14, 50)
            self.cell(-150, 20, f"SALES ANALYSIS for {month_name} {year}",
                      new_x="LMARGIN", new_y="NEXT", align="C")
            self.ln(10)

        def footer(self):
            self.set_y(-15)
            self.set_font("Malgun", "", 10)
            self.set_text_color(169, 169, 169)
            self.cell(0, 10, f"page {self.page_no()} / {{nb}}", align="C")

    pdf = PDF("P", "mm", "A4")
    pdf.add_font("Malgun", "", FONT_REG)
    pdf.add_font("Malgun", "B", FONT_BOLD)
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Malgun", size=7.5)
    pdf.add_page()
    line_height = pdf.font_size * 2

    pdf.cell(0, line_height, f"Date : {current_date}", new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.cell(0, line_height, "Made : Jonghwan SEO", new_x="LMARGIN", new_y="NEXT", align="R")

    list_df = [df_bran, df_prod, df_cust, df_item]
    list_df_name = ["1. MARGIN by BRANDs", "2. MARGIN by PRODUCTs", "3. MARGIN by CUSTOMERs", "4. MARGIN by ITEMs"]

    for lst, lst_name in zip(list_df, list_df_name):
        pdf.set_x(8)
        pdf.set_font("Malgun", style="B")
        pdf.cell(150, line_height * 2, lst_name)
        pdf.ln(line_height)

        lst = lst.astype("str")
        TABLE_COL_NAMES = tuple(lst.columns)
        TABLE_DATA = tuple(tuple(r) for r in lst.values.tolist())

        pdf.ln(line_height)
        col_width = pdf.epw / (len(TABLE_COL_NAMES) + 1)

        def render_table_header():
            pdf.set_font("Malgun", style="B")
            pdf.set_fill_color(220, 220, 220)
            for col_name in TABLE_COL_NAMES:
                w = col_width * 2 if col_name == TABLE_COL_NAMES[0] else col_width
                pdf.cell(w, line_height, col_name, border=1, align="C", fill=1)
            pdf.ln(line_height)
            pdf.set_font("Malgun", style="")

        render_table_header()
        for row in TABLE_DATA:
            if pdf.will_page_break(line_height):
                render_table_header()
            for datum in row:
                if datum == row[0]:
                    pdf.cell(col_width * 2, line_height, datum, border=1)
                else:
                    pdf.cell(col_width, line_height, datum, border=1, align="C")
            pdf.ln(line_height)

    pdf.output(fname)


# ===== 5. 메인 루프 ===========================================
def main():
    df_s_all = load_sales()
    cost = load_cost_table()
    current_date = datetime.datetime.now().strftime("%b-%d, %Y")

    for m in MONTHS:
        last_day = calendar.monthrange(YEAR, m)[1]
        start_date = f"{YEAR}-{m:02d}-01"
        end_date = f"{YEAR}-{m:02d}-{last_day:02d}"
        month_name = datetime.date(YEAR, m, 1).strftime("%b")
        stem = f"sales analysis report {YEAR}-{m:02d}"

        result = build_month(df_s_all, cost, start_date, end_date)
        if result is None:
            print(f"[skip] {YEAR}-{m:02d} : 데이터 없음")
            continue

        df, df_cust, df_prod, df_item, df_bran = result
        save_excel(f"{stem}.xlsx", df, df_cust, df_prod, df_item, df_bran)
        save_pdf(f"{stem}.pdf", month_name, str(YEAR), current_date, df_cust, df_prod, df_item, df_bran)

        tot = df_bran[df_bran["Brand"] == "TOTAL"].iloc[0]
        print(f"[done] {stem}  rows={len(df)}  Sales={tot['Sales(IDR)']}  "
              f"Margin={tot['Margin(%)']}  Cover={tot['Coverage(%)']}")

    print("완료: 모든 월 보고서 생성")


if __name__ == "__main__":
    main()

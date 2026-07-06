"""
월별 마진 보고서 일괄 생성기
- 매출자료(6.2.1 Delivery Status) + 구매자료(9.1 AP Invoice Status)를 결합
- 지정한 연도의 여러 개월에 대해 각각 Excel(.xlsx) + PDF 보고서를 생성

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


# ===== 1. 원본 자료 로딩 (전체 1회) ============================
def load_sales():
    sheets = pd.ExcelFile(SALES_FILE).sheet_names
    sheet = next((s for s in sheets if s.startswith("6.2.1")), sheets[0])
    df = pd.read_excel(SALES_FILE, sheet_name=sheet, skiprows=6)
    df = df.drop("Unnamed: 0", axis=1, errors="ignore")
    df["Delivery Date.1"] = pd.to_datetime(df["Delivery Date.1"])
    return df


def load_purchase():
    # AP Invoice = 매입 → Item Code별 '최신 구매단가'만 사용
    sheets = pd.ExcelFile(PURCHASE_FILE).sheet_names
    sheet = next((s for s in sheets if s.startswith("9.1")), sheets[0])
    df = pd.read_excel(PURCHASE_FILE, sheet_name=sheet, skiprows=7)
    df = df.rename(columns={"Unit Price (IDR)": "P_Price(IDR)"})
    df = df[["Item Code", "A/P Inv Date", "P_Price(IDR)"]].dropna(subset=["Item Code"])
    df["A/P Inv Date"] = pd.to_datetime(df["A/P Inv Date"])
    df = df.sort_values("A/P Inv Date").drop_duplicates("Item Code", keep="last")
    return df[["Item Code", "P_Price(IDR)"]]


# ===== 서식 함수 ==============================================
def int_num(x):
    return "{:,}".format(int(x))


def pct(x):
    return "{:.1f}".format(x * 100) + "%"


# ===== 2. 한 달치 집계 =========================================
def build_month(df_s_all, df_p, start_date, end_date):
    """해당 기간의 원본 df와 4개 집계표(cust/prod/item/bran)를 반환. 데이터 없으면 None."""
    mask = (df_s_all["Delivery Date.1"] >= start_date) & (df_s_all["Delivery Date.1"] <= end_date)
    df_s = df_s_all.loc[mask]
    if df_s.empty:
        return None

    # Row_Data : 구매단가 결합
    df = df_s.merge(df_p, how="left", on="Item Code").dropna(subset=["P_Price(IDR)"])
    if df.empty:
        return None
    df["P_Amount(IDR)"] = df["P_Price(IDR)"] * df["Q'ty"]
    df["Margin(IDR)"] = df["Discounted Amount (IDR)"] - df["P_Amount(IDR)"]
    df["Margin(%)"] = df["Margin(IDR)"] / df["Discounted Amount (IDR)"]
    df["Brand"] = df["Brand"].fillna("ETC")
    df["Type 3"] = df["Type 3"].replace("Others", "ETC")

    # 2. 업체별
    df_cust = df.groupby("Buyer")[["Discounted Amount (IDR)", "Margin(IDR)"]].agg(sum)
    df_cust["Margin(%)"] = df_cust["Margin(IDR)"] / df_cust["Discounted Amount (IDR)"]
    df_cust.loc["TOTAL", :] = df_cust.sum()
    df_cust.loc["TOTAL", "Margin(%)"] = df_cust.loc["TOTAL", "Margin(IDR)"] / df_cust.loc["TOTAL", "Discounted Amount (IDR)"]
    df_cust["Sales Ratio(%)"] = df_cust["Discounted Amount (IDR)"] / df_cust.loc["TOTAL", "Discounted Amount (IDR)"]
    df_cust["Margin Ratio(%)"] = df_cust["Margin(IDR)"] / df_cust.loc["TOTAL", "Margin(IDR)"]
    df_cust = df_cust.sort_values(["Margin(IDR)"], ascending=False).reset_index()

    # 3. 상품별 (Type 3 + Type 2)
    df_prod = df.groupby(["Type 3", "Type 2"])[["Q'ty", "Discounted Amount (IDR)", "Margin(IDR)"]].agg(sum)
    df_prod["Margin(%)"] = df_prod["Margin(IDR)"] / df_prod["Discounted Amount (IDR)"]
    df_prod.loc[("TOTAL", "ALL"), :] = df_prod.sum()
    df_prod.loc[("TOTAL", "ALL"), "Margin(%)"] = df_prod.loc[("TOTAL", "ALL"), "Margin(IDR)"] / df_prod.loc[("TOTAL", "ALL"), "Discounted Amount (IDR)"]
    df_prod["Sales Ratio(%)"] = df_prod["Discounted Amount (IDR)"] / df_prod.loc[("TOTAL", "ALL"), "Discounted Amount (IDR)"]
    df_prod["Margin Ratio(%)"] = df_prod["Margin(IDR)"] / df_prod.loc[("TOTAL", "ALL"), "Margin(IDR)"]
    df_prod = df_prod.sort_values(["Margin(IDR)"], ascending=False).reset_index()

    # 4. 아이템별 (Description)
    df_item = df.groupby(["Description"])[["Q'ty", "Discounted Amount (IDR)", "Margin(IDR)"]].agg(sum)
    df_item["Margin(%)"] = df_item["Margin(IDR)"] / df_item["Discounted Amount (IDR)"]
    df_item.loc["TOTAL", :] = df_item.sum()
    df_item.loc["TOTAL", "Margin(%)"] = df_item.loc["TOTAL", "Margin(IDR)"] / df_item.loc["TOTAL", "Discounted Amount (IDR)"]
    df_item["Sales Ratio(%)"] = df_item["Discounted Amount (IDR)"] / df_item.loc["TOTAL", "Discounted Amount (IDR)"]
    df_item["Margin Ratio(%)"] = df_item["Margin(IDR)"] / df_item.loc["TOTAL", "Margin(IDR)"]
    df_item = df_item.sort_values(["Margin(IDR)"], ascending=False).reset_index()

    # 5. 브랜드별
    df_bran = df.groupby(["Brand"])[["Discounted Amount (IDR)", "Margin(IDR)"]].agg(sum)
    df_bran["Margin(%)"] = df_bran["Margin(IDR)"] / df_bran["Discounted Amount (IDR)"]
    df_bran.loc["TOTAL", :] = df_bran.sum()
    df_bran.loc["TOTAL", "Margin(%)"] = df_bran.loc["TOTAL", "Margin(IDR)"] / df_bran.loc["TOTAL", "Discounted Amount (IDR)"]
    df_bran["Sales Ratio(%)"] = df_bran["Discounted Amount (IDR)"] / df_bran.loc["TOTAL", "Discounted Amount (IDR)"]
    df_bran["Margin Ratio(%)"] = df_bran["Margin(IDR)"] / df_bran.loc["TOTAL", "Margin(IDR)"]
    df_bran = df_bran.sort_values(["Margin(IDR)"], ascending=False).reset_index()

    # 열이름 변경
    for lst in [df_cust, df_prod, df_item, df_bran]:
        lst.rename(columns={"Discounted Amount (IDR)": "Sales(IDR)"}, inplace=True)
    df_prod = df_prod.rename(columns={"Type 3": "Product", "Type 2": "Type"})

    # 서식(콤마/퍼센트) 적용
    df_prod["Q'ty"] = df_prod["Q'ty"].astype("int").apply(lambda x: "{:,}".format(x))
    df_item["Q'ty"] = df_item["Q'ty"].astype("int").apply(lambda x: "{:,}".format(x))
    for lst in [df_cust, df_prod, df_item, df_bran]:
        lst["Sales(IDR)"] = lst["Sales(IDR)"].apply(int_num)
        lst["Margin(IDR)"] = lst["Margin(IDR)"].apply(int_num)
        lst["Margin(%)"] = lst["Margin(%)"].apply(pct)
        lst["Sales Ratio(%)"] = lst["Sales Ratio(%)"].apply(pct)
        lst["Margin Ratio(%)"] = lst["Margin Ratio(%)"].apply(pct)

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
    df_p = load_purchase()
    current_date = datetime.datetime.now().strftime("%b-%d, %Y")

    for m in MONTHS:
        last_day = calendar.monthrange(YEAR, m)[1]
        start_date = f"{YEAR}-{m:02d}-01"
        end_date = f"{YEAR}-{m:02d}-{last_day:02d}"
        month_name = datetime.date(YEAR, m, 1).strftime("%b")   # Jan, Feb ...
        stem = f"sales analysis report {YEAR}-{m:02d}"

        result = build_month(df_s_all, df_p, start_date, end_date)
        if result is None:
            print(f"[skip] {YEAR}-{m:02d} : 데이터 없음")
            continue

        df, df_cust, df_prod, df_item, df_bran = result
        save_excel(f"{stem}.xlsx", df, df_cust, df_prod, df_item, df_bran)
        save_pdf(f"{stem}.pdf", month_name, str(YEAR), current_date, df_cust, df_prod, df_item, df_bran)
        print(f"[done] {stem}.xlsx / .pdf  (rows={len(df)})")

    print("완료: 모든 월 보고서 생성")


if __name__ == "__main__":
    main()

"""
월별 마진 보고서 일괄 생성기 (IDR + USD)
- 매출자료(6.2.1 Delivery Status) + 원가자료(9.2 Inventory Ledger)를 결합
- 통화별(IDR / USD)로 각각 Excel(.xlsx) + PDF 보고서를 생성

[원가 산정 방식]
  - 구매단가 = 재고원장(9.2 Inventory Ledger)의 'Unit Cost' 컬럼 (Item Code별 1건)
    · IDR: 9.2 Inventory_Ledger_For_Costing.xlsx
    · USD: 9.2 Inventory_Ledger_For_Costing_USD.xlsx
  - Unit Cost가 0/공란이면 원장 내 대체 원가열로 재계산(Ending→Beginning→GR→Landed)
  - 그래도 없는 품목은 cost_overrides.csv 로 '빈 원가 채움'(fill-only).
    실제 원가가 원장에 생기면 자동 무시됨(수동보정은 금회 한정)

[매출 정합성]
  - 매입단가가 없는(미매칭) 품목도 매출에는 포함 → 월 매출합계가 원본 피벗과 일치
  - 원가 없는 행은 마진 계산에서 제외(N/A)하고, 표에 'Coverage(%)'(마진산출 커버율) 표기

출력 파일명: 'sales analysis report YYYY-MM.pdf/.xlsx' (IDR),
             'sales analysis report YYYY-MM USD.pdf/.xlsx' (USD)

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
START_MONTH = int(os.getenv("START_MONTH", "6"))
END_MONTH = int(os.getenv("END_MONTH", "6"))
MONTHS = range(START_MONTH, END_MONTH + 1)

SALES_FILE = os.getenv("SALES_FILE", "6.2.1 Delivery Status(List Only).xlsx")
PURCHASE_FILE = os.getenv("PURCHASE_FILE", "9.2 Inventory_Ledger_For_Costing.xlsx")
PURCHASE_FILE_USD = os.getenv("PURCHASE_FILE_USD", "9.2 Inventory_Ledger_For_Costing_USD.xlsx")
LOGO = os.getenv("LOGO_FILE", "ASCENDO_Blue.png")

# Unicode 폰트 (Windows 기본 - 맑은 고딕)
FONT_REG = os.getenv("FONT_REG", r"C:\Windows\Fonts\malgun.ttf")
FONT_BOLD = os.getenv("FONT_BOLD", r"C:\Windows\Fonts\malgunbd.ttf")

# 통화별 설정: (통화코드, 원장파일, 파일명 접미)
CURRENCIES = [
    ("IDR", PURCHASE_FILE, ""),
    ("USD", PURCHASE_FILE_USD, " USD"),
]
# ─────────────────────────────────────────────────────────────


# ===== 1. 원본 자료 로딩 =======================================
def load_sales():
    sheets = pd.ExcelFile(SALES_FILE).sheet_names
    sheet = next((s for s in sheets if s.startswith("6.2.1")), sheets[0])
    df = pd.read_excel(SALES_FILE, sheet_name=sheet, skiprows=6)
    df = df.drop("Unnamed: 0", axis=1, errors="ignore")
    df["Delivery Date.1"] = pd.to_datetime(df["Delivery Date.1"])
    return df


def load_cost_table(ledger=None, cur="IDR"):
    """Item Code별 구매단가 = 재고원장의 'Unit Cost'(Unnamed:10).
    0/공란이면 대체 원가열로 재계산(Ending→Beginning→GoodReceipt→Landed).
    그래도 없으면 cost_overrides.csv 로 '빈 원가만' 채움(fill-only, 금회 한정).
    """
    ledger = ledger or PURCHASE_FILE
    sheets = pd.ExcelFile(ledger).sheet_names
    sheet = next((s for s in sheets if s.startswith("9.2")), sheets[0])
    raw = pd.read_excel(ledger, sheet_name=sheet, skiprows=4)

    def num(col):
        return pd.to_numeric(raw[col], errors="coerce")

    def ratio(amount, qty):
        r = amount / qty
        return r.where((qty > 0) & (amount > 0))

    unit = num("Unnamed: 10")                    # Unit Cost
    cost = unit.where(unit > 0)
    cost = cost.fillna(ratio(num("Unnamed: 20"), num("Unnamed: 19")))   # Ending
    cost = cost.fillna(ratio(num("Unnamed: 4"),  num("Unnamed: 3")))    # Beginning
    cost = cost.fillna(ratio(num("Unnamed: 6"),  num("Good Receipt")))  # Good Receipt
    cost = cost.fillna(ratio(num("Landed Cost"), num("Good Receipt")))  # Landed

    p = pd.DataFrame({"Item Code": raw["Item Code"], "P_Price": cost})
    p = p.dropna(subset=["Item Code"])
    p = p[p["P_Price"] > 0]
    p = p.drop_duplicates("Item Code", keep="last")

    # 수동 원가 보정 (원장에 원가가 아예 없는 품목만 채움)
    ovf = Path("cost_overrides.csv")
    if ovf.exists():
        ov = pd.read_csv(ovf)
        pcol = f"P_Price({cur})"
        if pcol in ov.columns:
            ov = ov.rename(columns={pcol: "P_Price"})
            ov["Item Code"] = ov["Item Code"].astype(str).str.strip()
            ov["P_Price"] = pd.to_numeric(ov["P_Price"], errors="coerce")
            ov = ov[["Item Code", "P_Price"]].dropna()
            have = set(p["Item Code"].astype(str).str.strip())
            ov = ov[~ov["Item Code"].isin(have)]      # fill-only: 실제 원가 있으면 무시
            p = pd.concat([p, ov], ignore_index=True)
    return p[["Item Code", "P_Price"]]


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
def summarize(df, keys, total_key, cur, with_qty=False, sort_col="Sales"):
    """keys 기준 집계 + TOTAL행(맨 위) + 비율/커버율. 정렬은 sort_col 내림차순.
    (매출=전체, 마진=원가있는 행만).  df 는 _Sales/_CostedSales/_Margin 컬럼 보유."""
    aggmap = {
        "Sales": ("_Sales", "sum"),
        "CostedSales": ("_CostedSales", "sum"),
        "Margin": ("_Margin", "sum"),
    }
    if with_qty:
        aggmap["Qty"] = ("Q'ty", "sum")
    g = df.groupby(keys).agg(**aggmap)

    g.loc[total_key, :] = g.sum()
    g["Margin(%)"] = g["Margin"] / g["CostedSales"]
    g["Coverage(%)"] = g["CostedSales"] / g["Sales"]
    g["Sales Ratio(%)"] = g["Sales"] / g.loc[total_key, "Sales"]
    g["Margin Ratio(%)"] = g["Margin"] / g.loc[total_key, "Margin"]
    g = g.reset_index()

    # TOTAL 행을 맨 위 고정, 나머지는 sort_col 내림차순
    first_key = keys[0] if isinstance(keys, list) else keys
    is_total = g[first_key].astype(str) == "TOTAL"
    g = pd.concat([g[is_total], g[~is_total].sort_values(sort_col, ascending=False)],
                  ignore_index=True)

    g = g.rename(columns={"Sales": f"Sales({cur})", "Margin": f"Margin({cur})"})
    order = list(keys) if isinstance(keys, list) else [keys]
    if with_qty:
        g = g.rename(columns={"Qty": "Q'ty"})
        order += ["Q'ty"]
    order += [f"Sales({cur})", f"Margin({cur})", "Margin(%)", "Coverage(%)",
              "Sales Ratio(%)", "Margin Ratio(%)"]
    return g[order]


def fmt(g, cur):
    if "Q'ty" in g.columns:
        g["Q'ty"] = g["Q'ty"].apply(lambda x: "" if pd.isna(x) else "{:,}".format(int(round(x))))
    g[f"Sales({cur})"] = g[f"Sales({cur})"].apply(int_num)
    g[f"Margin({cur})"] = g[f"Margin({cur})"].apply(int_num)
    for c in ["Margin(%)", "Coverage(%)", "Sales Ratio(%)", "Margin Ratio(%)"]:
        g[c] = g[c].apply(pct)
    return g


# ===== 2. 한 달치 집계 =========================================
def build_month(df_s_all, cost, start_date, end_date, cur="IDR"):
    sales_col = f"Discounted Amount ({cur})"
    mask = (df_s_all["Delivery Date.1"] >= start_date) & (df_s_all["Delivery Date.1"] <= end_date)
    df_s = df_s_all.loc[mask]
    if df_s.empty:
        return None

    df = df_s.merge(cost, how="left", on="Item Code")
    df["Brand"] = df["Brand"].fillna("ETC")
    df["Type 3"] = df["Type 3"].replace("Others", "ETC")
    df["_Sales"] = df[sales_col]
    df["_PAmt"] = df["P_Price"] * df["Q'ty"]
    df["_Margin"] = df["_Sales"] - df["_PAmt"]
    df["_CostedSales"] = df["_Sales"].where(df["P_Price"].notna())

    df_cust = fmt(summarize(df, "Buyer", "TOTAL", cur, sort_col="Margin"), cur)
    df_prod = fmt(summarize(df, ["Type 3", "Type 2"], ("TOTAL", "ALL"), cur, with_qty=True, sort_col="Margin")
                  .rename(columns={"Type 3": "Product", "Type 2": "Type"}), cur)
    df_item = fmt(summarize(df, "Description", "TOTAL", cur, with_qty=True, sort_col="Margin"), cur)
    df_bran = fmt(summarize(df, "Brand", "TOTAL", cur), cur)   # 브랜드: 매출순 유지

    # Row_Data 정리
    row = (df.drop(columns=["_CostedSales", "_Sales"])
             .rename(columns={"P_Price": f"P_Price({cur})", "_PAmt": f"P_Amount({cur})",
                              "_Margin": f"Margin({cur})"}))
    row["Margin(%)"] = pd.to_numeric(row[f"Margin({cur})"]) / row[sales_col]
    return row, df_cust, df_prod, df_item, df_bran


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
def save_pdf(fname, month_name, year, current_date, df_cust, df_prod, df_item, df_bran, cur="IDR"):
    class PDF(FPDF):
        def header(self):
            self.set_font("Malgun", "B", 20)
            self.set_text_color(32, 32, 32)
            self.set_fill_color(240, 255, 255)
            self.set_line_width(0.4)
            self.cell(0, 20, border=1, align="C", fill=1)
            self.image(LOGO, 15, 14, 50)
            self.cell(-150, 20, f"SALES ANALYSIS for {month_name} {year} ({cur})",
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


# ===== 4b. 미커버(원가없는) 품목 목록 Excel ===================
def export_uncovered(df_s_all, cost, fname="uncovered_items.xlsx"):
    """원가 매칭이 안 된 매출을 월별·코드별로 정리해 Excel 저장 (IDR 기준)."""
    SALES_COL = "Discounted Amount (IDR)"
    codes = set(cost["Item Code"].astype(str).str.strip())
    start = f"{YEAR}-{START_MONTH:02d}-01"
    end = f"{YEAR}-{END_MONTH:02d}-{calendar.monthrange(YEAR, END_MONTH)[1]:02d}"
    s = df_s_all[(df_s_all["Delivery Date.1"] >= start) & (df_s_all["Delivery Date.1"] <= end)].copy()
    s["Month"] = s["Delivery Date.1"].dt.strftime("%Y-%m")
    s["code"] = s["Item Code"].astype(str).str.strip()
    unc = s[~s["code"].isin(codes)].copy()

    summ = s.groupby("Month").agg(Sales_IDR=(SALES_COL, "sum")).reset_index()
    us = unc.groupby("Month").agg(Uncovered_IDR=(SALES_COL, "sum"),
                                  Uncovered_rows=(SALES_COL, "size")).reset_index()
    summ = summ.merge(us, how="left", on="Month").fillna({"Uncovered_IDR": 0, "Uncovered_rows": 0})
    summ["Coverage(%)"] = (1 - summ["Uncovered_IDR"] / summ["Sales_IDR"]) * 100
    tot = pd.DataFrame({"Month": ["TOTAL"], "Sales_IDR": [summ["Sales_IDR"].sum()],
                        "Uncovered_IDR": [summ["Uncovered_IDR"].sum()],
                        "Uncovered_rows": [summ["Uncovered_rows"].sum()]})
    tot["Coverage(%)"] = (1 - tot["Uncovered_IDR"] / tot["Sales_IDR"]) * 100
    summ = pd.concat([summ, tot], ignore_index=True)

    detail = (unc.groupby(["Month", "Item Code", "Description", "Brand", "Type 3", "Type 2"])
              .agg(Qty=("Q'ty", "sum"), Sales_IDR=(SALES_COL, "sum"))
              .reset_index().sort_values(["Month", "Sales_IDR"], ascending=[True, False]))
    bycode = (unc.groupby(["Item Code", "Description", "Brand", "Type 3", "Type 2"])
              .agg(Qty=("Q'ty", "sum"), Sales_IDR=(SALES_COL, "sum"), Rows=(SALES_COL, "size"))
              .reset_index().sort_values("Sales_IDR", ascending=False))

    writer = pd.ExcelWriter(fname, engine="openpyxl")
    summ.to_excel(writer, sheet_name="1.Summary(월별)", index=False)
    bycode.to_excel(writer, sheet_name="2.코드별 합계", index=False)
    detail.to_excel(writer, sheet_name="3.월별·코드별 상세", index=False)
    writer._save()
    print(f"[uncovered] {fname}  미커버 {unc[SALES_COL].sum():,.0f} IDR / {len(unc)}행 / 코드 {unc['code'].nunique()}종")


# ===== 5. 메인 루프 ===========================================
def main():
    df_s_all = load_sales()
    current_date = datetime.datetime.now().strftime("%b-%d, %Y")

    for cur, ledger, suffix in CURRENCIES:
        cost = load_cost_table(ledger, cur)
        for m in MONTHS:
            last_day = calendar.monthrange(YEAR, m)[1]
            start_date = f"{YEAR}-{m:02d}-01"
            end_date = f"{YEAR}-{m:02d}-{last_day:02d}"
            month_name = datetime.date(YEAR, m, 1).strftime("%b")
            stem = f"sales analysis report {YEAR}-{m:02d}{suffix}"

            result = build_month(df_s_all, cost, start_date, end_date, cur)
            if result is None:
                print(f"[skip] {stem} : 데이터 없음")
                continue

            df, df_cust, df_prod, df_item, df_bran = result
            save_excel(f"{stem}.xlsx", df, df_cust, df_prod, df_item, df_bran)
            save_pdf(f"{stem}.pdf", month_name, str(YEAR), current_date,
                     df_cust, df_prod, df_item, df_bran, cur)

            tot = df_bran[df_bran["Brand"] == "TOTAL"].iloc[0]
            print(f"[done] {stem}  rows={len(df)}  Sales({cur})={tot[f'Sales({cur})']}  "
                  f"Margin={tot['Margin(%)']}  Cover={tot['Coverage(%)']}")

        if cur == "IDR":
            export_uncovered(df_s_all, cost)   # 미커버 목록(통화무관) 1회

    print("완료: IDR + USD 보고서 생성")


if __name__ == "__main__":
    main()

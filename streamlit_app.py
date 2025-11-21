import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import os

# ---------------------------
# 主处理函数
# ---------------------------
def generate_iif(income_file, gl_file, mapping_path, date_str):
    DATE = date_str
    mapping = pd.read_csv(mapping_path)

    je = pd.DataFrame(columns=["DATE","ACCOUNT","DEBIT","CREDIT","MEMO","NAME","DOCNUM"])
    doc_counter = 1

    def apply_mapping(row, mapping, file_type):
        for _, r in mapping.iterrows():
            if file_type == "Income" and r["Type"] == "Name":
                if str(row.get("Name","")).strip() == str(r["MatchValue"]).strip():
                    return r
            elif file_type == "GL" and r["Type"] == "Keyword":
                if str(r["MatchValue"]) in str(row.get("GL Account","")):
                    return r
        return None

    def process_file(df, file_type):
        nonlocal doc_counter
        for _, row in df.iterrows():
            rule = apply_mapping(row, mapping, file_type)
            if rule is None:
                continue

            je_date = DATE if file_type == "Income" else row["Date"]

            vendor = ""
            if file_type == "GL" and str(rule.get("UseVendor","No")).lower() == "yes":
                vendor = row.get("Payee / Payer","")

            if file_type == "Income":
                memo_str = str(rule.get("MemoTemplate","")).format(
                    current_month=datetime.strptime(DATE,"%m/%d/%Y").strftime("%B %Y")
                )
            else:
                memo_str = str(row.get("Remarks",""))

            docnum = f"JE{doc_counter:03d}"

            if file_type == "Income":
                amt = float(str(row["Amount"]).replace(",",""))
                direction = rule["Direction"]
                if amt < 0:
                    amt = -amt
                    direction = "CR" if direction=="DR" else "DR"
                debit_val = amt if direction=="DR" else 0
                credit_val = amt if direction=="CR" else 0
                debit_acc = rule["DebitAcc"]
                credit_acc = rule["CreditAcc"]

            if file_type == "GL":
                debit_val = float(str(row.get("Debit",0)).replace(",","")) if rule["Direction"]=="DR" else 0
                credit_val = float(str(row.get("Debit",0)).replace(",","")) if rule["Direction"]=="CR" else 0
                debit_acc = rule["DebitAcc"]
                credit_acc = rule["CreditAcc"]

            if debit_val > 0:
                je.loc[len(je)] = [je_date, debit_acc, debit_val, 0, memo_str, vendor, docnum]
                je.loc[len(je)] = [je_date, credit_acc, 0, debit_val, memo_str, vendor, docnum]
            elif credit_val > 0:
                je.loc[len(je)] = [je_date, debit_acc, 0, credit_val, memo_str, vendor, docnum]
                je.loc[len(je)] = [je_date, credit_acc, credit_val, 0, memo_str, vendor, docnum]

            doc_counter += 1

    # 处理 Income
    df_income = pd.read_csv(income_file)
    df_income.columns = ['Name','Amount','Col2']
    if 'Col2' in df_income.columns:
        df_income = df_income.drop(columns=['Col2'])
    df_income['Name'] = df_income['Name'].str.strip()
    process_file(df_income, "Income")

    # 处理 GL
    df_gl = pd.read_csv(gl_file)
    process_file(df_gl, "GL")

    # 生成 IIF
    output = []
    output.append("!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\tNAME\tDOCNUM")
    output.append("!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\tNAME\tDOCNUM")
    output.append("!ENDTRNS")

    for docnum, group in je.groupby("DOCNUM"):
        debit_row = group[group["DEBIT"] > 0].iloc[0]
        output.append(f"TRNS\tGENERAL JOURNAL\t{debit_row['DATE']}\t{debit_row['ACCOUNT']}\t{debit_row['DEBIT']:.2f}\t{debit_row['MEMO']}\t{debit_row['NAME']}\t{docnum}")
        for _, row in group.iterrows():
            if row["ACCOUNT"] == debit_row["ACCOUNT"] and row["DEBIT"] > 0:
                continue
            amount = row["DEBIT"] if row["DEBIT"] > 0 else -row["CREDIT"]
            output.append(f"SPL\tGENERAL JOURNAL\t{row['DATE']}\t{row['ACCOUNT']}\t{amount:.2f}\t{row['MEMO']}\t{row['NAME']}\t{docnum}")
        output.append("ENDTRNS")

    return "\n".join(output)


# ---------------------------
# Streamlit 界面
# ---------------------------
st.set_page_config(page_title="QuickBooks IIF Generator", layout="centered")
st.title("💼 QuickBooks IIF Generator")
st.markdown("上传 Income Statement & General Ledger，选择 Property 和日期，自动生成 `.iif` 文件。")

# 扫描当前目录下所有 Mapping CSV
property_options = []
for file in os.listdir("."):
    if file.endswith("Mapping.csv") and file != "Mapping.csv":  # 排除通用 Mapping.csv
        property_options.append(file.replace(" Mapping.csv",""))

# 添加 Other 选项
property_options.append("Other")

# 选择 Property
property_selected = st.selectbox("🏠 选择 Property", property_options)

# 根据选择确定 mapping 文件路径
if property_selected == "Other":
    mapping_path = "Mapping.csv"  # 通用 mapping
else:
    mapping_path = f"{property_selected} Mapping.csv"  # property 专用 mapping

# 输入日期
date_input = st.date_input("🗓️ 选择日期", value=datetime(2025,9,30))
date_str = date_input.strftime("%m/%d/%Y")

# 上传文件
income_file = st.file_uploader("📂 上传 Income Statement CSV", type=["csv"])
gl_file = st.file_uploader("📂 上传 General Ledger CSV", type=["csv"])

# 生成 IIF
if st.button("🚀 生成 IIF 文件"):
    if not income_file or not gl_file:
        st.error("⚠️ 请上传 Income Statement 和 General Ledger 文件。")
    else:
        with st.spinner(f"正在为 {property_selected} 生成 IIF 文件..."):
            mapping_path_a = f"{property_selected} Mapping.csv"
            iif_text_a = generate_iif(income_file, gl_file, mapping_path_a, date_str)
            
            buffer_a = BytesIO()
            buffer_a.write(iif_text_a.encode("utf-8"))
            buffer_a.seek(0)
            st.download_button(
                label=f"⬇️ 下载 {property_selected} IIF",
                data=buffer_a,
                file_name=f"{property_selected}_JE_{date_str}.iif",
                mime="text/plain"
            )

            # 2️⃣ 另一个公司 IIF
            mapping_path_b = "{property_selected} Moo Housing Mapping.csv"  # 你提供新的 Mapping
            iif_text_b = generate_iif(income_file, gl_file, mapping_path_b, date_str)
            
            buffer_b = BytesIO()
            buffer_b.write(iif_text_b.encode("utf-8"))
            buffer_b.seek(0)
            st.download_button(
                label=f"⬇️ 下载 OtherCompany IIF",
                data=buffer_b,
                file_name=f"Moo Housing_JE_{date_str}.iif",
                mime="text/plain"
            )

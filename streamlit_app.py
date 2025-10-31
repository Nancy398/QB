import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import os

# ---------------------------
# 封装主处理逻辑
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

    df_income = pd.read_csv(income_file)
    df_income.columns = ['Name','Amount','Col2']
    if 'Col2' in df_income.columns:
        df_income = df_income.drop(columns=['Col2'])
    df_income['Name'] = df_income['Name'].str.strip()
    process_file(df_income, "Income")

    df_gl = pd.read_csv(gl_file)
    process_file(df_gl, "GL")

    # 生成 IIF 内容
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

# 自动读取 mapping 文件夹下所有 Mapping 文件
mapping_dir = "mapping"
property_options = []
for file in os.listdir(mapping_dir):
    if file.endswith("Mapping.csv"):
        property_options.append(file.replace(" Mapping.csv",""))

if not property_options:
    st.error("⚠️ 未找到 mapping 文件，请确保 mapping 文件夹中有 *Mapping.csv 文件。")
    st.stop()

# 选择 Property
property_selected = st.selectbox("🏠 选择 Property", property_options)
mapping_path = os.path.join(mapping_dir, f"{property_selected} Mapping.csv")

# 输入日期
date_input = st.date_input("🗓️ 选择日期", value=datetime(2025,9,30))
date_str = date_input.strftime("%m/%d/%Y")

# 上传文件
income_file = st.file_uploader("📂 上传 Income Statement CSV", type=["csv"])
gl_file = st.file_uploader("📂 上传 General Ledger CSV", type=["csv"])

if st.button("🚀 生成 IIF 文件"):
    if not income_file or not gl_file:
        st.error("⚠️ 请上传 Income Statement 和 General Ledger 文件。")
    else:
        with st.spinner(f"正在为 {property_selected} 生成 IIF 文件..."):
            iif_text = generate_iif(income_file, gl_file, mapping_path, date_str)
            buffer = BytesIO()
            buffer.write(iif_text.encode("utf-8"))
            buffer.seek(0)
            file_name = f"{property_selected}_JE_{date_str.replace('/','-')}.iif"

            st.success(f"✅ {property_selected} 的 IIF 文件生成成功！")
            st.download_button(
                label="⬇️ 下载 IIF 文件",
                data=buffer,
                file_name=file_name,
                mime="text/plain"
            )

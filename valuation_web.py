#!/usr/bin/env python
# coding: utf-8

# In[2]:


import streamlit as st
import pandas as pd

st.set_page_config(page_title="財務指標自動計算工具", layout="wide")

st.title("財務指標自動計算工具 (Web 版)")

fields = [
    ("股價", "price"),
    ("流通股數", "shares"),
    ("每股帳面價值", "bvps"),
    ("每股營收", "sales_per_share"),
    ("每股盈餘(EPS)", "eps"),
    ("淨利總額", "net_income"),
    ("營收總額", "sales_total"),
    ("淨值總額", "equity_total"),
    ("現金與約當現金", "cash"),
    ("有息負債", "debt"),
    ("EBITDA", "ebitda"),
    ("FCF", "fcf"),
    ("資產總額", "assets"),
    ("每股現金股利", "div_per_share"),
    ("股利總額", "div_total")
]

# 側邊輸入
st.sidebar.header("請輸入財務數值")
inputs = {}
for name, key in fields:
    inputs[key] = st.sidebar.text_input(f"{name}", value="")

def safe_float(val):
    try:
        return float(val.replace(',', '').replace(' ', ''))
    except:
        return None

v = {k: safe_float(inputs[k]) for _, k in fields}

# 計算財務指標
results = {}
results["市值(Market Cap)"] = (
    v["price"] * v["shares"] if v["price"] is not None and v["shares"] is not None else None
)
results["PE"] = (
    results["市值(Market Cap)"] / v["net_income"] if results["市值(Market Cap)"] is not None and v["net_income"] not in (None, 0) else None
)
results["PB"] = (
    results["市值(Market Cap)"] / v["equity_total"] if results["市值(Market Cap)"] is not None and v["equity_total"] not in (None, 0) else None
)
results["PS"] = (
    results["市值(Market Cap)"] / v["sales_total"] if results["市值(Market Cap)"] is not None and v["sales_total"] not in (None, 0) else None
)
results["EV"] = (
    results["市值(Market Cap)"] + v["debt"] - v["cash"]
    if results["市值(Market Cap)"] is not None and v["debt"] is not None and v["cash"] is not None else None
)
results["EV/EBITDA"] = (
    results["EV"] / v["ebitda"] if results["EV"] is not None and v["ebitda"] not in (None, 0) else None
)
results["EV/FCF"] = (
    results["EV"] / v["fcf"] if results["EV"] is not None and v["fcf"] not in (None, 0) else None
)
results["EV/Sales"] = (
    results["EV"] / v["sales_total"] if results["EV"] is not None and v["sales_total"] not in (None, 0) else None
)
results["ROE"] = (
    v["net_income"] / v["equity_total"] if v["net_income"] is not None and v["equity_total"] not in (None, 0) else None
)
results["ROA"] = (
    v["net_income"] / v["assets"] if v["net_income"] is not None and v["assets"] not in (None, 0) else None
)
results["殖利率(Yield)"] = (
    v["div_per_share"] / v["price"] if v["div_per_share"] is not None and v["price"] not in (None, 0) else None
)

# 顯示結果
st.header("自動計算指標")
df = pd.DataFrame([
    {"指標": k, "計算結果": (f"{val:,.4f}" if isinstance(val, float) else "")}
    for k, val in results.items()
])
st.table(df)

# 匯出功能
st.header("匯出 Excel")
if st.button("匯出Excel"):
    df_input = pd.DataFrame(list(v.items()), columns=["項目", "輸入值"])
    df_out = pd.DataFrame([
        (k, (f"{val:,.4f}" if isinstance(val, float) else "")) for k, val in results.items()
    ], columns=["指標", "計算結果"])
    with pd.ExcelWriter("財務指標計算結果.xlsx", engine="openpyxl") as writer:
        df_input.to_excel(writer, sheet_name="輸入數據", index=False)
        df_out.to_excel(writer, sheet_name="財務指標", index=False)
    with open("財務指標計算結果.xlsx", "rb") as file:
        st.download_button("下載Excel", file, file_name="財務指標計算結果.xlsx")

if st.button("一鍵清除"):
    st.experimental_rerun()


# In[ ]:





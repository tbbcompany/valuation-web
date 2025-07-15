import streamlit as st
import pandas as pd
import json
import io

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

# 可自訂哪些欄位必填
required_keys = ["price", "shares", "net_income", "equity_total", "sales_total"]

# 公式→中文說明對應表
formula_hints = {
    "price": "股價",
    "shares": "流通股數",
    "bvps": "每股帳面價值",
    "sales_per_share": "每股營收",
    "eps": "每股盈餘(EPS)",
    "net_income": "淨利總額",
    "sales_total": "營收總額",
    "equity_total": "淨值總額",
    "cash": "現金與約當現金",
    "debt": "有息負債",
    "ebitda": "EBITDA",
    "fcf": "FCF",
    "assets": "資產總額",
    "div_per_share": "每股現金股利",
    "div_total": "股利總額",
    "market_cap": "市值",
    "ev": "企業價值(EV)",
    "None": "空值"
}

# 預設公式
default_formulas = {
    "市值(Market Cap)": "price * shares",
    "PE": "market_cap / net_income if net_income else None",
    "PB": "market_cap / equity_total if equity_total else None",
    "PS": "market_cap / sales_total if sales_total else None",
    "EV": "market_cap + debt - cash",
    "EV/EBITDA": "ev / ebitda if ebitda else None",
    "EV/FCF": "ev / fcf if fcf else None",
    "EV/Sales": "ev / sales_total if sales_total else None",
    "ROE": "net_income / equity_total if equity_total else None",
    "ROA": "net_income / assets if assets else None",
    "殖利率(Yield)": "div_per_share / price if price else None"
}

# ====== Session State 初始化 ======
defaults = {k: "" for _, k in fields}
if "inputs" not in st.session_state:
    st.session_state.inputs = defaults.copy()
if "formulas" not in st.session_state:
    st.session_state.formulas = default_formulas.copy()
if "formula_backup" not in st.session_state:
    st.session_state.formula_backup = False  # 是否已備份過
if "admin_pwd_fail" not in st.session_state:
    st.session_state.admin_pwd_fail = False  # 密碼失敗提示 flag
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False      # 管理員登入狀態

# ====== 資料輸入區 ======
st.sidebar.header("請輸入財務數值")
inputs = {}
for name, key in fields:
    val = st.sidebar.text_input(
        f"{name}",
        value=st.session_state.inputs.get(key, ""),
        key=key
    )
    st.session_state.inputs[key] = val
    inputs[key] = val

def safe_float(val):
    try:
        return float(val.replace(',', '').replace(' ', ''))
    except:
        return None

v = {k: safe_float(inputs[k]) for _, k in fields}

# ====== 必填欄位檢查 ======
missing = [k for k in required_keys if (inputs[k].strip() == "" or v[k] is None)]
if missing:
    missnames = "、".join([n for n, k in fields if k in missing])
    st.warning(f"⚠️ 請填寫以下必要欄位再進行計算：{missnames}")
    can_calculate = False
else:
    can_calculate = True

# ====== 公式計算主體 ======
results = {}
if can_calculate:
    local_vars = v.copy()
    try:
        market_cap = eval(st.session_state.formulas["市值(Market Cap)"], {}, local_vars)
        local_vars["market_cap"] = market_cap
        results["市值(Market Cap)"] = market_cap
        for k in st.session_state.formulas:
            if k == "市值(Market Cap)":
                continue
            val = eval(st.session_state.formulas[k], {}, local_vars)
            results[k] = val
            local_vars[k.lower().replace("/", "_").replace("(", "").replace(")", "").replace(" ", "_")] = val
    except Exception as e:
        st.error(f"計算發生錯誤: {e}")

# ====== 結果顯示 ======
st.header("自動計算指標")
if can_calculate:
    df = pd.DataFrame([
        {"指標": k, "計算結果": (f"{val:,.4f}" if isinstance(val, float) else "")}
        for k, val in results.items()
    ])
    st.table(df)
else:
    st.info("請先填完必要欄位，才能自動計算財務指標。")

# ====== 匯出 Excel ======
st.header("匯出 Excel")
if can_calculate and st.button("匯出Excel"):
    df_input = pd.DataFrame(list(v.items()), columns=["項目", "輸入值"])
    df_out = pd.DataFrame([
        (k, (f"{val:,.4f}" if isinstance(val, float) else "")) for k, val in results.items()
    ], columns=["指標", "計算結果"])
    with pd.ExcelWriter("財務指標計算結果.xlsx", engine="openpyxl") as writer:
        df_input.to_excel(writer, sheet_name="輸入數據", index=False)
        df_out.to_excel(writer, sheet_name="財務指標", index=False)
    with open("財務指標計算結果.xlsx", "rb") as file:
        st.download_button("下載Excel", file, file_name="財務指標計算結果.xlsx")

# ====== 一鍵清除功能 ======
if st.button("一鍵清除"):
    st.session_state.inputs = defaults.copy()
    st.rerun()

# ====== 管理員公式設定區（強制備份/匯出/還原/目前公式/密碼提示/公式轉換/取消登出） ======
with st.expander("管理員功能（公式設定/備份/還原）"):
    st.markdown("**目前所有公式如下：**")
    # ====== 顯示「一般人看得懂」的公式說明 ======
    for k, expr in st.session_state.formulas.items():
        show_expr = expr
        for en, zh in formula_hints.items():
            show_expr = show_expr.replace(en, zh)
        st.write(f"**{k}：** {show_expr}")

    st.code(json.dumps(st.session_state.formulas, ensure_ascii=False, indent=2), language="json")

    # ====== 管理員密碼區或管理員編輯模式 ======
    admin_password = "tbb1840"   # 請換成你自己的密碼
    # 如果已登入 admin_mode
    if st.session_state.admin_mode:
        st.success("管理員已登入。為保險請先按下『匯出公式』完成備份，才可進行修改或還原！")
        # 備份下載
        if not st.session_state.formula_backup:
            st.info("請先下載一次公式備份才能進行編輯或還原。")
            if st.button("匯出目前公式（下載json備份）"):
                backup = json.dumps(st.session_state.formulas, ensure_ascii=False, indent=2)
                st.download_button("下載公式.json", io.BytesIO(backup.encode("utf-8")), file_name="公式備份.json")
                st.session_state.formula_backup = True
        else:
            # 還原公式上傳
            uploaded_file = st.file_uploader("上傳備份公式（.json）進行還原", type=["json"])
            if uploaded_file:
                try:
                    data = json.load(uploaded_file)
                    if isinstance(data, dict):
                        st.session_state.formulas = data
                        st.success("已成功還原所有公式，立即生效！")
                        st.session_state.formula_backup = False
                        st.rerun()
                    else:
                        st.error("檔案格式錯誤，請上傳正確的公式json。")
                except Exception as e:
                    st.error(f"讀取公式檔錯誤：{e}")

            st.markdown("---")
            # 可編輯欄位
            for k in st.session_state.formulas:
                new_formula = st.text_input(f"{k} 公式", value=st.session_state.formulas[k], key=f"formula_{k}")
                st.session_state.formulas[k] = new_formula
            btn1, btn2 = st.columns(2)
            with btn1:
                if st.button("儲存公式（即時生效）"):
                    st.success("已更新公式，立即套用！")
                    st.session_state.formula_backup = False  # 儲存後再強制下次編輯前備份
                    st.session_state.admin_mode = False       # 儲存後自動登出
                    st.rerun()
            with btn2:
                if st.button("取消/登出"):
                    st.info("已登出管理員模式！")
                    st.session_state.admin_mode = False
                    st.session_state.admin_pwd_fail = False
                    st.rerun()
    else:
        pwd = st.text_input("請輸入管理密碼", type="password")
        if pwd:
            if pwd == admin_password:
                st.session_state.admin_pwd_fail = False
                st.session_state.admin_mode = True
                st.experimental_rerun()
            else:
                st.session_state.admin_pwd_fail = True
        if st.session_state.admin_pwd_fail:
            st.error("密碼錯誤，請重新輸入！")
        elif not pwd:
            st.info("僅管理員可修改、還原、編輯公式，請輸入正確密碼。")

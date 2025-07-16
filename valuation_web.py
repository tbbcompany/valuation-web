mport streamlit as st
import pandas as pd
import json
import io

st.set_page_config(page_title="財務指標自動計算工具", layout="wide")
st.title("財務指標自動計算工具 (Web 版)")

# --- 預設欄位與欄位管理 ---
default_fields = [
    {"name": "股價", "key": "price", "required": True},
    {"name": "流通股數", "key": "shares", "required": True},
    {"name": "每股帳面價值", "key": "bvps", "required": False},
    {"name": "每股營收", "key": "sales_per_share", "required": False},
    {"name": "每股盈餘(EPS)", "key": "eps", "required": False},
    {"name": "淨利總額", "key": "net_income", "required": True},
    {"name": "營收總額", "key": "sales_total", "required": True},
    {"name": "淨值總額", "key": "equity_total", "required": True},
    {"name": "現金與約當現金", "key": "cash", "required": False},
    {"name": "有息負債", "key": "debt", "required": False},
    {"name": "EBITDA", "key": "ebitda", "required": False},
    {"name": "FCF", "key": "fcf", "required": False},
    {"name": "資產總額", "key": "assets", "required": False},
    {"name": "每股現金股利", "key": "div_per_share", "required": False},
    {"name": "股利總額", "key": "div_total", "required": False},
]

# 公式→中文說明對應表（欄位與公式管理都用）
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

# --- 預設公式 ---
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
if "fields" not in st.session_state:
    st.session_state.fields = default_fields.copy()
if "formulas" not in st.session_state:
    st.session_state.formulas = default_formulas.copy()
if "inputs" not in st.session_state:
    st.session_state.inputs = {f['key']: "" for f in st.session_state.fields}
if "formula_backup" not in st.session_state:
    st.session_state.formula_backup = False  # 是否已備份過
if "admin_pwd_fail" not in st.session_state:
    st.session_state.admin_pwd_fail = False  # 密碼失敗提示 flag
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False      # 管理員登入狀態

# 取得必填欄位
required_keys = [f['key'] for f in st.session_state.fields if f.get("required")]

# ====== 資料輸入區 ======
st.sidebar.header("請輸入財務數值")
inputs = {}
for f in st.session_state.fields:
    val = st.sidebar.text_input(
        f"{f['name']}",
        value=st.session_state.inputs.get(f['key'], ""),
        key=f['key']
    )
    st.session_state.inputs[f['key']] = val
    inputs[f['key']] = val

def safe_float(val):
    try:
        return float(val.replace(',', '').replace(' ', ''))
    except:
        return None

v = {f['key']: safe_float(inputs[f['key']]) for f in st.session_state.fields}

# ====== 必填欄位檢查 ======
missing = [k for k in required_keys if (inputs[k].strip() == "" or v[k] is None)]
if missing:
    missnames = "、".join([f['name'] for f in st.session_state.fields if f['key'] in missing])
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
    st.session_state.inputs = {f['key']: "" for f in st.session_state.fields}
    st.rerun()

# ====== 管理員欄位/公式/匯出還原管理 ======
with st.expander("管理員功能（欄位/公式/匯出/還原）"):
    st.markdown("**目前所有公式如下：**")
    for k, expr in st.session_state.formulas.items():
        show_expr = expr
        for en, zh in formula_hints.items():
            show_expr = show_expr.replace(en, zh)
        st.write(f"**{k}：** {show_expr}")
    st.code(json.dumps(st.session_state.formulas, ensure_ascii=False, indent=2), language="json")

    admin_password = "tbb1840"   # 改成你自己的密碼

    # ====== 欄位動態管理（需登入） ======
    if st.session_state.admin_mode:
        st.success("管理員已登入。")
        with st.expander("欄位管理", expanded=True):
            st.write("目前所有欄位：")
            st.table(pd.DataFrame(st.session_state.fields))
            # 新增欄位
            st.subheader("新增欄位")
            col1, col2, col3 = st.columns([2,2,1])
            with col1:
                new_name = st.text_input("欄位中文名稱", key="addfield_name")
            with col2:
                new_key = st.text_input("欄位英文key", key="addfield_key")
            with col3:
                new_required = st.checkbox("必填", value=False, key="addfield_required")
            if st.button("新增欄位"):
                if new_name and new_key and not any(f['key'] == new_key for f in st.session_state.fields):
                    st.session_state.fields.append({"name": new_name, "key": new_key, "required": new_required})
                    st.session_state.inputs[new_key] = ""
                    st.success(f"已新增欄位：{new_name} ({new_key})")
                    st.rerun()
                elif any(f['key'] == new_key for f in st.session_state.fields):
                    st.error("此英文key已存在，請換一個。")
                else:
                    st.error("欄位名稱與key皆需填寫。")
            # 刪除欄位
            st.subheader("刪除欄位")
            del_options = [f"{f['name']} ({f['key']})" for f in st.session_state.fields]
            del_choice = st.selectbox("選擇要刪除的欄位", del_options, key="del_field_choice")
            if st.button("刪除選定欄位"):
                del_key = st.session_state.fields[del_options.index(del_choice)]['key']
                st.session_state.fields = [f for f in st.session_state.fields if f['key'] != del_key]
                st.session_state.inputs.pop(del_key, None)
                st.success("已刪除欄位，立即生效")
                st.rerun()
            # 欄位匯出
            if st.button("匯出欄位清單"):
                field_json = json.dumps(st.session_state.fields, ensure_ascii=False, indent=2)
                st.download_button("下載欄位清單.json", io.BytesIO(field_json.encode("utf-8")), file_name="欄位清單.json")
            # 欄位還原
            up_field_file = st.file_uploader("上傳欄位清單(.json)進行還原", type=["json"], key="fields_restore")
            if up_field_file:
                try:
                    data = json.load(up_field_file)
                    if isinstance(data, list) and all("key" in d and "name" in d for d in data):
                        st.session_state.fields = data
                        # 移除不存在欄位的 inputs
                        for k in list(st.session_state.inputs.keys()):
                            if k not in [f["key"] for f in data]:
                                st.session_state.inputs.pop(k)
                        st.success("欄位清單已還原，立即生效")
                        st.rerun()
                    else:
                        st.error("檔案格式錯誤，請確認上傳正確的欄位清單json。")
                except Exception as e:
                    st.error(f"讀取欄位檔錯誤：{e}")

        # ====== 公式備份/還原、修改區 ======
        st.markdown("---")
        st.subheader("公式管理（同上）")
        if not st.session_state.formula_backup:
            st.info("請先下載一次公式備份才能進行編輯或還原。")
            if st.button("匯出目前公式（下載json備份）"):
                backup = json.dumps(st.session_state.formulas, ensure_ascii=False, indent=2)
                st.download_button("下載公式.json", io.BytesIO(backup.encode("utf-8")), file_name="公式備份.json")
                st.session_state.formula_backup = True
        else:
            # 還原公式上傳
            uploaded_file = st.file_uploader("上傳備份公式（.json）進行還原", type=["json"], key="formulas_restore")
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
                    st.session_state.formula_backup = False
                    st.session_state.admin_mode = False
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
            st.info("僅管理員可修改、還原、編輯公式/欄位，請輸入正確密碼。")

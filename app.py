import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ================= 1. 全局页面配置与 CSS =================
st.set_page_config(page_title="地磅管家 Pro", page_icon="🚛", layout="wide", initial_sidebar_state="auto")

st.markdown("""
<style>
.block-container { padding-top: 3.5rem !important; padding-bottom: 1rem !important; max-width: 100% !important; }
.stSelectbox > div > div > div, .stNumberInput input, .stButton > button, .stPopover > button {
    height: 36px !important; min-height: 36px !important; font-size: 14px !important;
}
.stTextInput>div>div, .stNumberInput>div>div, .stSelectbox>div>div {
    background-color: transparent !important; border: 1px solid #ddd !important; border-radius: 4px !important;
}
div[data-testid="stPopoverBody"] { max-width: 92vw !important; }
button:has(p:contains("✅")) {
    background-color: #f0fdf4 !important; border-color: #22c55e !important; color: #166534 !important;
}
@media (max-width: 768px) {
    [data-testid="stHorizontalBlock"] { flex-direction: row !important; flex-wrap: wrap !important; }
    [data-testid="column"] {
        width: calc(50% - 0.5rem) !important; flex: 0 0 calc(50% - 0.5rem) !important;
        min-width: calc(50% - 0.5rem) !important; margin-bottom: 0.5rem !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ================= 2. 左侧边栏 =================
st.sidebar.markdown("### 📅 报表时间")
report_date = st.sidebar.date_input("报表日期", value=datetime.now(), label_visibility="collapsed")
report_time = st.sidebar.text_input("时间段", value="07:00-18:00", label_visibility="collapsed")
st.sidebar.divider()

st.sidebar.markdown("### ⚙️ 数值取整设置")
st.sidebar.caption("勾选后取整数，不勾选保留2位小数")
round_retail = st.sidebar.checkbox("零售(微信/现金)取整", value=True)
round_sign = st.sidebar.checkbox("签单客户取整", value=False)
round_fee = st.sidebar.checkbox("加工费取整", value=False)
round_freight = st.sidebar.checkbox("运费取整", value=False)
st.sidebar.divider()

# ================= 3. 核心算法 =================
def do_round(val, category=""):
    if pd.isna(val): return 0.0
    val = float(val)
    should_round = False
    if category == "retail" and round_retail: should_round = True
    elif category == "sign" and round_sign: should_round = True
    elif category == "fee" and round_fee: should_round = True
    elif category == "freight" and round_freight: should_round = True
    
    if should_round: return float(int(val + 0.5) if val >= 0 else int(val - 0.5))
    else: return round(val + 1e-9, 2)

# ⭐️ 核心升级：动态抹除无效小数位
def fmt_val(v, category=""):
    if pd.isna(v): return "0"
    v = float(v)
    should_round = False
    if category == "retail" and round_retail: should_round = True
    elif category == "sign" and round_sign: should_round = True
    elif category == "fee" and round_fee: should_round = True
    elif category == "freight" and round_freight: should_round = True
    
    if should_round: return f"{int(v + 0.5) if v >= 0 else int(v - 0.5)}"
    else: 
        res = f"{v:.2f}"
        if '.' in res: res = res.rstrip('0').rstrip('.')
        return res

def fmt_weight(v):
    if pd.isna(v): return "0"
    v = float(v)
    res = f"{v:.2f}"
    if '.' in res: res = res.rstrip('0').rstrip('.')
    return res

def safe_concat(dfs):
    valid_dfs = [df for df in dfs if not df.empty]
    if not valid_dfs: return dfs[0] if dfs else pd.DataFrame()
    return pd.concat(valid_dfs, ignore_index=True)

def create_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=['客户名称', '余额']).to_excel(writer, sheet_name='客户余额', index=False)
        pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价']).to_excel(writer, sheet_name='加工费规则', index=False)
        pd.DataFrame(columns=['原始名称', '标准名称']).to_excel(writer, sheet_name='物料归类映射', index=False)
        pd.DataFrame(columns=['单号', '重车时间', '车号', '货物名称', '标准归类', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费单价', '加工费', '备注2']).to_excel(writer, sheet_name='过磅明细', index=False)
        pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额']).to_excel(writer, sheet_name='公司配送-运费', index=False)
        pd.DataFrame(columns=['日期', '客户名称', '收入类型', '金额', '备注']).to_excel(writer, sheet_name='财务收入明细', index=False)
    return output.getvalue()

st.sidebar.markdown("### 📋 空白模板")
st.sidebar.download_button(label="下载更新架构的总账本模板", data=create_template(), file_name="地磅总账本_空白_V2.xlsx")

def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() == '': return ''
    try:
        f_val = float(val)
        if f_val > 30000: return pd.to_datetime(f_val, unit='D', origin='1899-12-30').strftime('%Y-%m-%d %H:%M')
    except: pass
    return str(val)

# ================= 4. 主界面：文件上传 =================
st.subheader("🚛 地磅管家 Pro")
st.divider()
st.markdown("#### 第一步：上传业务文件")

c1, c2 = st.columns(2)
with c1: db_file = st.file_uploader("📂 1. 上传【总账本.xlsx】", type=['xlsx'])
with c2: daily_file = st.file_uploader("📝 2. 上传【今日过磅单】", type=['xls', 'xlsx', 'csv'])

@st.dialog("➕ 添加新司机")
def add_driver_modal(t):
    new_d = st.text_input(f"请输入【{t}】的新司机姓名：", placeholder="输入姓名")
    if st.button("确认添加", type="primary", use_container_width=True):
        if new_d.strip(): st.session_state[f"custom_drv_{t}"] = new_d.strip()
        st.rerun()

# ================= 5. 核心拦截流水线 =================
if db_file is not None and daily_file is not None:
    try:
        xls = pd.ExcelFile(db_file)
        df_bal = pd.read_excel(xls, sheet_name='客户余额') if '客户余额' in xls.sheet_names else pd.DataFrame(columns=['客户名称', '余额'])
        
        if '加工费规则' in xls.sheet_names:
            df_rules = pd.read_excel(xls, sheet_name='加工费规则')
            if '物料名称' not in df_rules.columns: df_rules = pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价'])
        else: df_rules = pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价'])
            
        if '物料归类映射' in xls.sheet_names:
            df_mapping = pd.read_excel(xls, sheet_name='物料归类映射')
            if '原始名称' not in df_mapping.columns: df_mapping = pd.DataFrame(columns=['原始名称', '标准名称'])
        else: df_mapping = pd.DataFrame(columns=['原始名称', '标准名称'])
            
        mapping_dict = dict(zip(df_mapping['原始名称'].astype(str).str.strip(), df_mapping['标准名称'].astype(str).str.strip()))
        df_hist = pd.read_excel(xls, sheet_name='过磅明细') if '过磅明细' in xls.sheet_names else pd.DataFrame()
        df_freight = pd.read_excel(xls, sheet_name='公司配送-运费') if '公司配送-运费' in xls.sheet_names else pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额'])
        df_income = pd.read_excel(xls, sheet_name='财务收入明细') if '财务收入明细' in xls.sheet_names else pd.DataFrame(columns=['日期', '客户名称', '收入类型', '金额', '备注'])
            
        if daily_file.name.endswith('.csv'): df_daily_raw = pd.read_csv(daily_file)
        else: df_daily_raw = pd.read_excel(daily_file)
            
        if '状态' in df_daily_raw.columns:
            df_daily_raw = df_daily_raw[~df_daily_raw['状态'].astype(str).str.contains('作废|生产', na=False)]
            
        df_daily_raw['单号'] = df_daily_raw.get('单号', '').astype(str).str.replace('.0', '', regex=False)
        df_daily_raw['重车时间'] = df_daily_raw.get('重车时间', '').apply(parse_excel_date)
        df_daily_raw['货物名称'] = df_daily_raw.get('货物名称', '未知物料').astype(str).str.strip()
        df_daily_raw['净重'] = pd.to_numeric(df_daily_raw.get('净重', 0), errors='coerce').fillna(0)
        df_daily_raw['单价'] = pd.to_numeric(df_daily_raw.get('单价', 0), errors='coerce').fillna(0)
        df_daily_raw['金额'] = pd.to_numeric(df_daily_raw.get('金额', 0), errors='coerce').fillna(0)
        df_daily_raw['过磅类型'] = df_daily_raw.get('过磅类型', '').astype(str)
        df_daily_raw['备注'] = df_daily_raw.get('备注', '').fillna('')
        df_daily_raw['备注2'] = ''
        
        df_daily_raw['_orig_idx'] = df_daily_raw.index
        df_daily_raw['汇报专用名称'] = ''

        # ---------------- 🚀 第一关：混合料拆分 ----------------
        mixed_mask = df_daily_raw['货物名称'].str.contains(r'\+', na=False)
        if mixed_mask.any():
            st.warning("⚠️ **第一步：检测到混合料，请配置拆分比例**")
            mixed_mats = df_daily_raw.loc[mixed_mask, '货物名称'].unique()
            split_ratios = {}
            for mat in mixed_mats:
                parts = mat.split('+')
                p1, p2 = parts[0].strip(), parts[1].strip() if len(parts)>1 else "其他料"
                col_s1, col_s2 = st.columns([1, 1])
                with col_s1: pct1 = st.number_input(f"【{mat}】[{p1}] 占比(%)", min_value=1.0, max_value=99.0, value=50.0, step=1.0, key=f"s1_{mat}")
                with col_s2: st.number_input(f"[{p2}] 占比(%)", value=100.0 - pct1, disabled=True, key=f"s2_{mat}")
                split_ratios[mat] = pct1
            
            split_confirmed = st.checkbox("✅ 比例已确认，立刻拆分并进入下一步", key="chk_split")
            
            new_rows = []
            for idx, row in df_daily_raw.iterrows():
                mat = str(row['货物名称'])
                if mat in split_ratios:
                    pct1 = split_ratios[mat]
                    pct2 = 100.0 - pct1
                    parts = mat.split('+')
                    mat1, mat2 = parts[0].strip(), parts[1].strip() if len(parts)>1 else "其他"
                    
                    r1, r2 = row.copy(), row.copy()
                    w_orig, p_orig, a_orig = row['净重'], row['单价'], row['金额']
                    t, dh = str(row['过磅类型']), str(row['单号'])
                    category = "none"
                    if '微信' in t or '现金' in t: category = "retail"
                    elif '签单' in t: category = "sign"
                    
                    if p_orig == 0 and a_orig == 0:
                        w1 = round(w_orig * (pct1 / 100.0), 3)
                        w2 = w_orig - w1
                        r1['净重'], r1['单价'], r1['金额'] = w1, 0.0, 0.0
                        r2['净重'], r2['单价'], r2['金额'] = w2, 0.0, 0.0
                        r1['_pre_rounded'], r2['_pre_rounded'] = False, False
                    else:
                        true_total_a = (w_orig * p_orig) if p_orig > 0 else a_orig
                        rounded_total = do_round(true_total_a, category)
                        w1 = round(w_orig * (pct1 / 100.0), 3)
                        w2 = w_orig - w1
                        exact_a1 = w1 * p_orig if p_orig > 0 else true_total_a * (pct1 / 100.0)
                        a1 = do_round(exact_a1, category)
                        a2 = round(rounded_total - a1, 2)
                        r1['净重'], r1['单价'], r1['金额'] = w1, p_orig, a1
                        r2['净重'], r2['单价'], r2['金额'] = w2, p_orig, a2
                        r1['_pre_rounded'], r2['_pre_rounded'] = True, True
                    
                    r1['货物名称'], r2['货物名称'] = mat1, mat2
                    r1['汇报专用名称'], r2['汇报专用名称'] = mat, mat 
                    r1['备注2'] = f"{mat1} {pct1:g}%+{mat2} {pct2:g}% 已拆分单号 {dh}"
                    r2['备注2'] = f"{mat1} {pct1:g}%+{mat2} {pct2:g}% 已拆分单号 {dh}"
                    new_rows.extend([r1, r2])
                else:
                    row['_pre_rounded'] = True if row['金额'] > 0 else False
                    new_rows.append(row)
            df_daily_raw = pd.DataFrame(new_rows)
            st.divider()
            if not split_confirmed: st.stop()

        # ---------------- 🚀 第二关：字典映射与新物料归类 ----------------
        df_daily_raw['标准归类'] = df_daily_raw['货物名称'].apply(lambda x: mapping_dict.get(x, x))
        df_daily_raw['汇报专用名称'] = df_daily_raw.apply(lambda r: r['标准归类'] if r['汇报专用名称'] == '' else r['汇报专用名称'], axis=1)
        
        known_originals = set(df_mapping['原始名称'].dropna().astype(str).str.strip().unique())
        known_standards = set(df_rules['物料名称'].dropna().astype(str).str.strip().unique())
        
        unknown_mats = [mat for mat in df_daily_raw['货物名称'].astype(str).str.strip().unique() 
                        if mat != '' and mat not in known_originals and mat not in known_standards]
        
        if unknown_mats:
            st.error("🛑 **第二步：发现未建立映射的未知物料！请指定父类：**")
            new_mapping_records = []
            cols = st.columns(3)
            mapping_inputs = {}
            map_options = ["(独立物料)", "(非销售/不计价)"] + sorted(list(known_standards))
            
            for i, u_mat in enumerate(unknown_mats):
                with cols[i % 3]:
                    mapping_inputs[u_mat] = st.selectbox(f"【{u_mat}】归类为：", ["(请选择...)"] + map_options, key=f"map_{u_mat}")
            
            map_confirmed = st.checkbox("✅ 归类映射已确认", key="chk_map")
            st.divider()
            if not map_confirmed: st.stop()
            
            for u_mat, sel_val in mapping_inputs.items():
                if sel_val != "(请选择...)":
                    std_name = u_mat if sel_val == "(独立物料)" else sel_val
                    mapping_dict[u_mat] = std_name
                    new_mapping_records.append({'原始名称': u_mat, '标准名称': std_name})
            if new_mapping_records:
                df_mapping = safe_concat([df_mapping, pd.DataFrame(new_mapping_records)])
                df_daily_raw['标准归类'] = df_daily_raw['货物名称'].apply(lambda x: mapping_dict.get(x, x))
                df_daily_raw['汇报专用名称'] = df_daily_raw.apply(lambda r: r['标准归类'] if r['汇报专用名称'] == r['货物名称'] else r['汇报专用名称'], axis=1)

        # ---------------- 🚀 第三关：拦截补填缺失单价 ----------------
        if '收货单位' not in df_daily_raw.columns: df_daily_raw['收货单位'] = ''
        def fix_shdw(row):
            shdw = str(row.get('收货单位', '')).strip()
            gb_type = str(row.get('过磅类型', '')).strip()
            if shdw == '' or shdw == 'nan':
                if gb_type == '': return '内部单'
                elif '微信' in gb_type or '现金' in gb_type: return '零售客户'
                else: return '未知客户'
            return shdw
        df_daily_raw['收货单位'] = df_daily_raw.apply(fix_shdw, axis=1)
        
        missing_mask = (df_daily_raw['单价'] == 0) & (df_daily_raw['金额'] == 0) & (df_daily_raw['过磅类型'].str.strip() != '') & (df_daily_raw['标准归类'] != '(非销售/不计价)')
        
        if missing_mask.any():
            st.warning("⚠️ **第三步：检测到以下有效销售记录缺失单价，请补充：**")
            missing_groups = df_daily_raw[missing_mask].groupby(['收货单位', '货物名称'])
            cols = st.columns(4)
            price_inputs = {}
            for i, ((cust, mat), _) in enumerate(missing_groups):
                with cols[i % 4]:
                    price_inputs[(cust, mat)] = st.number_input(f"【{cust}】\n{mat} (元)", min_value=0.0, format="%.2f", key=f"miss_p_{cust}_{mat}")
            
            price_confirmed = st.checkbox("✅ 单价已全部补齐", key="chk_price")
            st.divider()
            if not price_confirmed: st.stop()
            
            for idx, row in df_daily_raw[missing_mask].iterrows():
                cust, mat = row['收货单位'], row['货物名称']
                df_daily_raw.at[idx, '单价'] = price_inputs.get((cust, mat), 0.0)

        # ---------------- 🚀 第四关：网页原地补齐加工费 ----------------
        missing_rules = []
        valid_sales = df_daily_raw[(df_daily_raw['过磅类型'].astype(str).str.strip() != '') & (df_daily_raw['标准归类'] != '(非销售/不计价)')]
        unique_combinations = valid_sales[['标准归类', '单价']].drop_duplicates()
        
        for _, r in unique_combinations.iterrows():
            mat = str(r['标准归类']).strip()
            price = float(r['单价'])
            is_exist = False
            if not df_rules.empty:
                match = df_rules[(df_rules['物料名称'].astype(str).str.strip() == mat) & (df_rules['销售单价'].astype(float) == price)]
                if not match.empty:
                    val = match.iloc[0]['加工费单价']
                    if pd.notna(val) and str(val).strip() != '': is_exist = True
            
            if not is_exist: missing_rules.append((mat, price))
                
        if missing_rules:
            st.error("🛑 **第四步：发现规则库中未记录的加工费单价！请在下方直接补齐（系统将为您自动建档）：**")
            cols = st.columns(4)
            fee_inputs = {}
            for i, (mat, price) in enumerate(missing_rules):
                with cols[i % 4]:
                    fee_inputs[(mat, price)] = st.number_input(f"【{mat}】\n单价:{price}元 的加工费", min_value=0.0, format="%.2f", key=f"miss_f_{mat}_{price}")
            
            fee_confirmed = st.checkbox("✅ 加工费已补齐，完成最终核算", key="chk_fee")
            st.divider()
            if not fee_confirmed: st.stop()
            
            new_rules = []
            for (mat, price), fee in fee_inputs.items():
                new_rules.append({'物料名称': mat, '销售单价': price, '加工费单价': fee})
            df_rules = safe_concat([df_rules, pd.DataFrame(new_rules)])

        # ================= 🚀 最终核算阶段 =================
        new_amts = []
        for idx, row in df_daily_raw.iterrows():
            if row['标准归类'] == '(非销售/不计价)': new_amts.append(0.0)
            elif row.get('_pre_rounded', False): new_amts.append(row['金额'])
            else:
                w, p, orig_a, t = row['净重'], row['单价'], row['金额'], str(row['过磅类型'])
                exact_amt = (w * p) if p > 0 else orig_a 
                if '微信' in t or '现金' in t: new_amts.append(do_round(exact_amt, "retail"))
                elif '签单' in t: new_amts.append(do_round(exact_amt, "sign"))
                else: new_amts.append(do_round(exact_amt, "none"))
        df_daily_raw['金额'] = new_amts
        
        def calc_fee_price(row):
            gb_type = str(row.get('过磅类型', '')).strip()
            std_name = str(row.get('标准归类', '')).strip()
            if gb_type == '' or std_name == '(非销售/不计价)': return 0.0 
            price = float(row['单价'])
            match = df_rules[(df_rules['物料名称'].astype(str).str.strip() == std_name) & (df_rules['销售单价'].astype(float) == price)]
            return float(match.iloc[0]['加工费单价']) if not match.empty else 0.0
            
        df_daily_raw['加工费单价'] = df_daily_raw.apply(calc_fee_price, axis=1)
        df_daily_raw['加工费'] = df_daily_raw.apply(lambda r: do_round(r['净重'] * r['加工费单价'], "fee") if r['加工费单价'] > 0 else 0.0, axis=1)
        
        df_report_base = df_daily_raw.groupby('_orig_idx').agg({
            '单号': 'first', '重车时间': 'first', '车号': 'first', '收货单位': 'first',
            '汇报专用名称': 'first', '标准归类': 'first', '过磅类型': 'first', '备注': 'first',
            '净重': 'sum', '金额': 'sum'
        }).reset_index(drop=True)
        
        df_sales_report = df_report_base[(df_report_base['过磅类型'].astype(str).str.strip() != '') & (df_report_base['标准归类'] != '(非销售/不计价)')].copy()
        
        all_known_drivers = sorted([str(d) for d in df_freight['司机姓名'].dropna().unique() if str(d).strip() and str(d)!='nan'])
        driver_options = ["(未选择)"] + all_known_drivers + ["➕ 手动输入新司机..."]
        all_known_custs = sorted(list(set(df_bal['客户名称'].dropna().astype(str)) | set(df_hist['收货单位'].dropna().astype(str))))
        all_known_custs = [c for c in all_known_custs if c.strip() and c != 'nan' and c != '内部单']
        cust_options = ["(不录入)"] + all_known_custs + ["➕ 手动输入新客户..."]

        # ================== 公司配送 ==================
        has_delivery = False
        freight_total = 0.0
        new_freight_records = []
        new_freight_df = pd.DataFrame()

        delivery_mask = df_sales_report['备注'].astype(str).str.contains('公司配送', na=False)
        has_delivery = delivery_mask.any()
        expander_title = f"🔴 🚚 检测到 {len(df_sales_report[delivery_mask])} 车公司配送，点击展开分配" if has_delivery else "🚚 公司配送与额外运费分配"

        with st.expander(expander_title, expanded=False):
            if has_delivery:
                delivery_df = df_sales_report[delivery_mask].copy()
                truck_counts = delivery_df['车号'].value_counts()
                unique_trucks = truck_counts.index.tolist()
                unique_delivery_custs = delivery_df['收货单位'].dropna().unique()
                
                mem_driver, mem_price = {}, {}
                if not df_freight.empty:
                    for _, r in df_freight.iterrows():
                        if pd.notna(r.get('车号')) and pd.notna(r.get('司机姓名')): mem_driver[str(r['车号'])] = str(r['司机姓名'])
                        if pd.notna(r.get('收货单位')) and pd.notna(r.get('运费单价')): mem_price[str(r['收货单位'])] = float(r['运费单价'])
                
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    if st.button("🔄 全选/反选", use_container_width=True):
                        curr = st.session_state.get("batch_sel", False)
                        st.session_state["batch_sel"] = not curr
                        for t in unique_trucks: st.session_state[f"chk_{t}"] = not curr
                        st.rerun()
                with col_b2:
                    with st.popover("⚙️ 批量设置运价/司机", use_container_width=True):
                        st.markdown("**1. 批量分配司机**")
                        b_drv = st.selectbox("统一分配司机", driver_options[:-1])
                        if st.button("应用司机(仅对勾选)", use_container_width=True):
                            for t in unique_trucks:
                                if st.session_state.get(f"chk_{t}", False):
                                    if b_drv != "(未选择)": st.session_state[f"custom_drv_{t}"] = b_drv
                            st.rerun()
                        
                        st.divider()
                        st.markdown("**2. 按客户设置运价**")
                        b_cust = st.selectbox("目标车辆", ["(对所有已勾选车辆)"] + list(unique_delivery_custs))
                        b_prc = st.number_input("统一设置运价", min_value=0.0, step=1.0, format="%.2f")
                        if st.button("应用运价", use_container_width=True):
                            for t in unique_trucks:
                                t_cust = delivery_df[delivery_df['车号'] == t].iloc[0]['收货单位']
                                if b_cust == "(对所有已勾选车辆)":
                                    if st.session_state.get(f"chk_{t}", False): st.session_state[f"p_{t}"] = b_prc
                                elif t_cust == b_cust:
                                    st.session_state[f"p_{t}"] = b_prc
                            st.rerun()
                
                driver_map, price_map = {}, {}

                def render_truck_row(t):
                    count = truck_counts[t]
                    this_df = delivery_df[delivery_df['车号'] == t]
                    
                    curr_drv = st.session_state.get(f"custom_drv_{t}", mem_driver.get(str(t), ""))
                    curr_prc = st.session_state.get(f"p_{t}", mem_price.get(str(this_df.iloc[0]['收货单位']), 0.0))
                    prefix = "✅ " if bool(curr_drv and curr_prc > 0) else "🚛 "
                    
                    c_chk, c_info, c_drv, c_prc = st.columns([2.5, 2, 2.5, 3])
                    with c_chk: st.checkbox(f"{prefix}{t} ({count}趟)", key=f"chk_{t}")
                    with c_info:
                        with st.popover("📋 单号明细", use_container_width=True):
                            for _, r in this_df.iterrows():
                                st.markdown(f"📄 **单号**: `{r['单号']}`\n\n🏢 **客户**: {r['收货单位']}\n\n📦 **货物**: {r['汇报专用名称']} | **净重**: `{r['净重']} 吨`\n\n🕒 **时间**: {r['重车时间']}")
                                st.markdown("---")
                    with c_drv:
                        opts = driver_options.copy()
                        if curr_drv and curr_drv not in opts: opts.insert(1, curr_drv)
                        idx = opts.index(curr_drv) if curr_drv in opts else 0
                        d_sel = st.selectbox(f"司机_{t}", opts, index=idx, key=f"d_sel_{t}", label_visibility="collapsed")
                        if d_sel == "➕ 手动输入新司机...": add_driver_modal(t)
                        else: driver_map[t] = d_sel if d_sel != "(未选择)" else ""
                            
                    with c_prc:
                        p_val = st.number_input(f"运价_{t}", value=curr_prc, step=1.0, format="%.2f", key=f"p_{t}", label_visibility="collapsed", placeholder="¥ 运价")
                        price_map[t] = p_val
                        
                    st.markdown("<hr style='margin: 0.3em 0; border-style: dashed; border-color: #eee;'/>", unsafe_allow_html=True)

                for t in unique_trucks[:4]: render_truck_row(t)
                if len(unique_trucks) > 4:
                    with st.expander(f"↓ 展开剩余 {len(unique_trucks)-4} 辆车", expanded=False):
                        for t in unique_trucks[4:]: render_truck_row(t)
                
                for idx, row in delivery_df.iterrows():
                    t = row['车号']
                    d_name = driver_map.get(t, "")
                    p_val = price_map.get(t, 0.0)
                    new_freight_records.append({
                        '单号': row['单号'], '重车时间': row['重车时间'], '车号': row['车号'],
                        '收货单位': row['收货单位'], '货物名称': row['汇报专用名称'], '净重': row['净重'],
                        '司机姓名': d_name, '运费单价': p_val, '运费金额': do_round(row['净重'] * p_val, "freight")
                    })
                new_freight_df = pd.DataFrame(new_freight_records)
                freight_total = new_freight_df['运费金额'].sum()
            else:
                f_val = st.number_input("今日无配送，若有额外运费(元)请补充:", value=0.0, step=1.0, format="%.2f", label_visibility="visible")
                freight_total = do_round(f_val, "freight")

        # ================== 财务资金登记 ==================
        with st.expander("💰 财务资金登记 (点击展开录入收入与预存)", expanded=False):
            tab_income, tab_deposit = st.tabs(["📥 收入登记", "💳 预存登记"])
            today_income_records = []

            with tab_income:
                if "income_rows" not in st.session_state: st.session_state.income_rows = 1
                for i in range(st.session_state.income_rows):
                    c_c, c_t, c_a, c_n = st.columns([2.5, 2, 2.5, 3])
                    label_vis = "visible" if i == 0 else "collapsed"
                    
                    with c_c:
                        c_sel = st.selectbox(f"客户名称", cust_options, key=f"inc_c_sel_{i}", label_visibility=label_vis)
                        if c_sel == "➕ 手动输入新客户...":
                            c_name = st.text_input(f"新客户", key=f"inc_c_new_{i}", label_visibility="collapsed", placeholder="客户名称")
                        else: c_name = c_sel if c_sel != "(不录入)" else ""
                    with c_t: 
                        i_type = st.selectbox(f"收入类型", ["微信", "现金", "银行卡", "其他"], key=f"inc_t_{i}", label_visibility=label_vis)
                    with c_a: 
                        i_amt = st.number_input(f"金额(元)", min_value=0.0, step=100.0, format="%.2f", key=f"inc_a_{i}", label_visibility=label_vis)
                    with c_n:
                        i_note = st.text_input(f"备注", key=f"inc_n_{i}", label_visibility=label_vis, placeholder="选填")
                    
                    st.markdown("<hr style='margin: 0.5em 0; border-style: dashed; border-color: #eee;'/>", unsafe_allow_html=True)
                    if c_name and i_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': c_name, '收入类型': i_type, '金额': i_amt, '备注': i_note})
                
                col_add, _ = st.columns([1,5])
                with col_add:
                    if st.button("➕ 新增收入行", key="add_income_btn"):
                        st.session_state.income_rows += 1
                        st.rerun()

            with tab_deposit:
                if "deposit_rows" not in st.session_state: st.session_state.deposit_rows = 1
                for i in range(st.session_state.deposit_rows):
                    c_c, c_t, c_a, c_n = st.columns([2.5, 2, 2.5, 3])
                    label_vis = "visible" if i == 0 else "collapsed"
                    
                    with c_c:
                        c_sel = st.selectbox(f"客户名称 ", cust_options, key=f"dep_c_sel_{i}", label_visibility=label_vis)
                        if c_sel == "➕ 手动输入新客户...":
                            c_name = st.text_input(f"新客户", key=f"dep_c_new_{i}", label_visibility="collapsed", placeholder="客户名称")
                        else: c_name = c_sel if c_sel != "(不录入)" else ""
                    with c_t: 
                        i_type = st.selectbox(f"预存类型", ["预存微信", "预存银行卡", "预存现金", "预存备用金", "预存其他"], key=f"dep_t_{i}", label_visibility=label_vis)
                    with c_a: 
                        i_amt = st.number_input(f"金额(元) ", min_value=0.0, step=100.0, format="%.2f", key=f"dep_a_{i}", label_visibility=label_vis)
                    with c_n:
                        i_note = st.text_input(f"备注 ", key=f"dep_n_{i}", label_visibility=label_vis, placeholder="选填")
                    
                    st.markdown("<hr style='margin: 0.5em 0; border-style: dashed; border-color: #eee;'/>", unsafe_allow_html=True)
                    if c_name and i_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': c_name, '收入类型': i_type, '金额': i_amt, '备注': i_note})
                
                col_add, _ = st.columns([1,5])
                with col_add:
                    if st.button("➕ 新增预存行", key="add_deposit_btn"):
                        st.session_state.deposit_rows += 1
                        st.rerun()

        # ---- 计算资金余额逻辑 ----
        df_cash = df_sales_report[df_sales_report['过磅类型'].astype(str).str.contains('现金', na=False)]
        df_wx = df_sales_report[df_sales_report['过磅类型'].astype(str).str.contains('微信', na=False)]
        df_sign = df_sales_report[df_sales_report['过磅类型'].astype(str).str.contains('签单', na=False)]

        orig_bal_dict = dict(zip(df_bal['客户名称'], df_bal['余额']))
        deposit_dict = {}

        retail_wx_amt = do_round(df_wx['金额'].sum(), "retail")
        retail_cash_amt = do_round(df_cash['金额'].sum(), "retail")
        if retail_wx_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售微信', '金额': retail_wx_amt, '备注': '自动汇总'})
        if retail_cash_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售现金', '金额': retail_cash_amt, '备注': '自动汇总'})

        for r in today_income_records:
            if "预存" in r['收入类型'] or r['收入类型'] == "银行卡":
                if r['客户名称'] != '零售客户':
                    deposit_dict[r['客户名称']] = deposit_dict.get(r['客户名称'], 0.0) + r['金额']

        bal_dict = {}
        sign_custs = df_sign.groupby('收货单位')
        all_custs = set(list(orig_bal_dict.keys()) + list(deposit_dict.keys()) + list(sign_custs.groups.keys()))
        
        for c in all_custs:
            spent = do_round(df_sign[df_sign['收货单位'] == c]['金额'].sum(), "sign") if c in sign_custs.groups else 0.0
            bal_dict[c] = orig_bal_dict.get(c, 0.0) + deposit_dict.get(c, 0.0) - spent

        daily_fee = df_daily_raw[(df_daily_raw['过磅类型'].astype(str).str.strip() != '') & (df_daily_raw['标准归类'] != '(非销售/不计价)')]['加工费'].sum()
        current_month = report_date.strftime("%Y-%m")
        monthly_fee = daily_fee
        if not df_hist.empty and '重车时间' in df_hist.columns and '加工费' in df_hist.columns:
            df_hist['加工费'] = pd.to_numeric(df_hist['加工费'], errors='coerce').fillna(0)
            monthly_fee += do_round(df_hist[df_hist['重车时间'].astype(str).str.startswith(current_month, na=False)]['加工费'].sum(), "fee")

        # ================= ⭐️ 汇报文本生成 (空车打“无”) =================
        report = f"{report_date.strftime('%y年%m月%d日')} {report_time}\n"
        
        if len(df_cash) == 0:
            report += "\n现金:无\n"
        else:
            report += f"\n现金:{len(df_cash)}车{fmt_weight(df_cash['净重'].sum())}吨{fmt_val(retail_cash_amt, 'retail')}元\n"
            for prod, grp in df_cash.groupby('汇报专用名称'):
                report += f"{prod}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨{fmt_val(grp['金额'].sum(), 'retail')}元\n"
        
        if len(df_wx) == 0:
            report += "\n微信:无\n"
        else:
            report += f"\n微信:{len(df_wx)}车{fmt_weight(df_wx['净重'].sum())}吨{fmt_val(retail_wx_amt, 'retail')}元\n"
            for prod, grp in df_wx.groupby('汇报专用名称'):
                report += f"{prod}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨{fmt_val(grp['金额'].sum(), 'retail')}元\n"
            
        if len(df_sign) == 0:
            report += "\n签单:无\n\n"
        else:
            report += f"\n签单:{len(df_sign)}车{fmt_weight(df_sign['净重'].sum())}吨\n"
            for cust, grp in sign_custs:
                report += f"{cust}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨\n"
                for prod, p_grp in grp.groupby('汇报专用名称'):
                    report += f"{prod}:{len(p_grp)}车{fmt_weight(p_grp['净重'].sum())}吨{fmt_val(p_grp['金额'].sum(), 'sign')} 元\n"
                cust_money = do_round(grp['金额'].sum(), "sign")
                report += f"共金额:{fmt_val(cust_money, 'sign')} 元\n"
                report += f"上日余额:{fmt_val(orig_bal_dict.get(cust, 0.0), 'sign')} 元\n"
                if deposit_dict.get(cust, 0.0) > 0: report += f"今日充值:{fmt_val(deposit_dict.get(cust, 0.0), 'sign')} 元\n"
                report += f"当日余额:{fmt_val(bal_dict.get(cust, 0.0), 'sign')} 元\n\n"
            
        pure_depositors = [c for c in deposit_dict.keys() if c not in sign_custs.groups]
        if pure_depositors:
            report += "【纯充值客户余额刷新】\n"
            for c in pure_depositors:
                report += f"{c} 今日充值:{fmt_val(deposit_dict.get(c, 0.0), 'sign')}元 | 最新余额:{fmt_val(bal_dict.get(c, 0.0), 'sign')}元\n\n"
            
        total_money = df_sales_report['金额'].sum()
        unsold_count = len(df_daily_raw) - len(df_daily_raw[df_daily_raw['过磅类型'].astype(str).str.strip() != ''])
        unsold_str = f" (内含未销售单据/废料等 {unsold_count} 车，已全量留底)" if unsold_count > 0 else ""
        
        report += f"1,当日销售共计:{len(df_sales_report)} 车{fmt_weight(df_sales_report['净重'].sum())} 吨 {fmt_val(total_money, 'sign')} 元,公司配送运费:{fmt_val(freight_total, 'freight')}元 ,合计:{fmt_val(total_money + freight_total, 'sign')} 元。{unsold_str}\n"
        report += f"2,当日加工费:{fmt_val(daily_fee, 'fee')} 元,{report_date.month}月1日-{report_date.day}日加工费合计:{fmt_val(monthly_fee, 'fee')} 元。\n"
        
        collection_parts = []
        if retail_wx_amt > 0: collection_parts.append(f"微信零售:{fmt_val(retail_wx_amt, 'retail')}元")
        if retail_cash_amt > 0: collection_parts.append(f"现金零售:{fmt_val(retail_cash_amt, 'retail')}元")
        
        custom_income_total = 0.0
        for r in today_income_records:
            if r['客户名称'] != '零售客户':
                c_name = r['客户名称']
                i_type = r['收入类型']
                amt_str = fmt_val(r['金额'], 'none')
                custom_income_total += float(r['金额'])
                if c_name: collection_parts.append(f"{c_name}{i_type}:{amt_str}元")
                else: collection_parts.append(f"{i_type}:{amt_str}元")
        
        total_collection = retail_wx_amt + retail_cash_amt + custom_income_total
        collection_str = ",".join(collection_parts) if collection_parts else "0元"
        
        report += f"3,当日合计收款:{collection_str},共计:{fmt_val(total_collection, 'none')} 元\n"

        st.divider()
        st.markdown("#### 第二步：复制每日汇报")
        st.code(report, language="text", line_numbers=False)

        # ---------------- 组装新账本 ----------------
        cols_to_drop = ['_pre_rounded', '_orig_idx', '汇报专用名称', '标准归类']
        for c in cols_to_drop:
            if c in df_daily_raw.columns: df_daily_raw = df_daily_raw.drop(columns=[c])
            
        new_df_bal = pd.DataFrame(list(bal_dict.items()), columns=['客户名称', '余额'])
        cols_to_keep = ['单号', '重车时间', '车号', '货物名称', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费单价', '加工费', '备注2']
        available_cols = [c for c in cols_to_keep if c in df_daily_raw.columns]
        
        new_df_hist = safe_concat([df_hist, df_daily_raw[available_cols]])
        new_df_freight = safe_concat([df_freight, new_freight_df])
        
        today_income_df = pd.DataFrame(today_income_records) if today_income_records else pd.DataFrame()
        new_df_income = safe_concat([df_income, today_income_df])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            new_df_bal.to_excel(writer, sheet_name='客户余额', index=False)
            df_rules.to_excel(writer, sheet_name='加工费规则', index=False)
            df_mapping.to_excel(writer, sheet_name='物料归类映射', index=False)
            new_df_hist.to_excel(writer, sheet_name='过磅明细', index=False)
            new_df_freight.to_excel(writer, sheet_name='公司配送-运费', index=False)
            new_df_income.to_excel(writer, sheet_name='财务收入明细', index=False)
            
        st.success("✅ 核算完成，可下载更新后的总账本", icon="✅")
        
        col_btn, _ = st.columns([1,2])
        with col_btn:
            st.download_button(
                label="💾 下载更新后总账本",
                data=output.getvalue(),
                file_name=f"{report_date.strftime('%Y%m%d')}_DiBang总账本.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"❌ 文件处理失败，请检查：\n1. 上传的两个文件格式是否正确；\n2. 总账本是否符合模板要求。\n\n错误详情：{str(e)}", icon="⚠️")

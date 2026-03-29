import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ================= 1. 全局页面配置与 CSS =================

st.set_page_config(page_title=“地磅管家 Pro”, page_icon=“🚛”, layout=“wide”, initial_sidebar_state=“auto”)

st.markdown(”””

<style>
/* ── 页面基础 ── */
*, *::before, *::after { box-sizing: border-box; }
.block-container { padding: 0.8rem 0.5rem !important; max-width: 100% !important; }

/* ── 关键修复：让 Streamlit 的列容器在手机上不溢出 ──
   Streamlit 用 data-testid="stHorizontalBlock" 包裹所有 st.columns()
   默认是 flex-wrap:nowrap，强制改成 wrap 后各列会在空间不足时自动折行
   但我们这里不想折行，而是让每列等比缩小适应屏幕宽度 */
[data-testid="stHorizontalBlock"] {
    flex-wrap: nowrap !important;
    align-items: flex-end !important;
    gap: 4px !important;
    width: 100% !important;
    overflow: hidden !important;     /* 彻底禁止溢出滚动 */
}

/* 每个列子容器：不允许超出父容器，允许收缩 */
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
    min-width: 0 !important;         /* ← 这是关键！默认min-width会撑开列 */
    flex-shrink: 1 !important;
    overflow: hidden !important;
    padding: 0 2px !important;
}

/* ── 所有输入组件：统一小尺寸，适配手机 ── */
.stSelectbox > div > div,
.stNumberInput > div > div,
.stTextInput > div > div {
    background: #fff !important;
    border: 1px solid #ddd !important;
    border-radius: 5px !important;
    min-width: 0 !important;
}

/* 输入框内文字和高度 */
.stSelectbox label,
.stNumberInput label,
.stTextInput label,
.stCheckbox label {
    font-size: 11px !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
}

input[type="number"],
input[type="text"] {
    font-size: 13px !important;
    padding: 4px 6px !important;
    min-width: 0 !important;
    width: 100% !important;
}

/* 下拉框内部文字截断 */
[data-testid="stSelectbox"] [data-baseweb="select"] > div {
    font-size: 12px !important;
    padding: 2px 4px !important;
    min-width: 0 !important;
}

/* 数字输入框加减按钮紧凑 */
button[aria-label="Step down"],
button[aria-label="Step up"] {
    width: 20px !important;
    padding: 0 !important;
    flex-shrink: 0 !important;
}

/* 复选框垂直居中 */
[data-testid="stCheckbox"] {
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    padding-top: 22px !important;   /* 对齐label高度 */
}

/* popover 弹窗不超屏 */
div[data-testid="stPopoverBody"] { max-width: 90vw !important; }

/* 按钮 */
.stButton > button, .stPopover > button {
    font-size: 12px !important;
    padding: 4px 6px !important;
    white-space: nowrap !important;
}
</style>

“””, unsafe_allow_html=True)

# ================= 2. 左侧边栏 =================

st.sidebar.markdown(”### 📅 报表时间”)
report_date = st.sidebar.date_input(“报表日期”, value=datetime.now(), label_visibility=“collapsed”)
report_time = st.sidebar.text_input(“时间段”, value=“07:00-18:00”, label_visibility=“collapsed”)
st.sidebar.divider()

st.sidebar.markdown(”### ⚙️ 数值取整设置”)
st.sidebar.caption(“勾选后取整数，不勾选保留2位小数”)
round_retail = st.sidebar.checkbox(“零售(微信/现金)取整”, value=True)
round_sign   = st.sidebar.checkbox(“签单客户取整”,       value=False)
round_fee    = st.sidebar.checkbox(“加工费取整”,         value=False)
round_freight= st.sidebar.checkbox(“运费取整”,           value=False)
st.sidebar.divider()

# ================= 3. 核心算法 =================

def do_round(val, category=””):
if pd.isna(val): return 0.0
val = float(val)
sr = (category==“retail” and round_retail) or (category==“sign” and round_sign)   
or (category==“fee” and round_fee) or (category==“freight” and round_freight)
if sr: return float(int(val+0.5) if val>=0 else int(val-0.5))
return round(val+1e-9, 2)

def fmt_val(v, category=””):
if pd.isna(v): return “0.00”
v = float(v)
sr = (category==“retail” and round_retail) or (category==“sign” and round_sign)   
or (category==“fee” and round_fee) or (category==“freight” and round_freight)
if sr: return f”{int(v+0.5) if v>=0 else int(v-0.5)}”
return f”{v:.2f}”

def fmt_weight(v):
if pd.isna(v): return “0”
v = float(v)
return f”{int(round(v))}” if abs(v-round(v))<1e-5 else f”{v:.2f}”

def safe_concat(dfs):
valid = [d for d in dfs if not d.empty]
if not valid: return dfs[0] if dfs else pd.DataFrame()
return pd.concat(valid, ignore_index=True)

def create_template():
output = io.BytesIO()
with pd.ExcelWriter(output, engine=‘openpyxl’) as writer:
pd.DataFrame(columns=[‘客户名称’,‘余额’]).to_excel(writer, sheet_name=‘客户余额’, index=False)
pd.DataFrame(columns=[‘物料名称’,‘销售单价’,‘加工费单价’]).to_excel(writer, sheet_name=‘加工费规则’, index=False)
pd.DataFrame(columns=[‘单号’,‘重车时间’,‘车号’,‘货物名称’,‘净重’,‘单价’,‘金额’,‘收货单位’,‘过磅类型’,‘备注’,‘加工费’]).to_excel(writer, sheet_name=‘过磅明细’, index=False)
pd.DataFrame(columns=[‘单号’,‘重车时间’,‘车号’,‘收货单位’,‘货物名称’,‘净重’,‘司机姓名’,‘运费单价’,‘运费金额’]).to_excel(writer, sheet_name=‘公司配送-运费’, index=False)
pd.DataFrame(columns=[‘日期’,‘客户名称’,‘收入类型’,‘金额’,‘备注’]).to_excel(writer, sheet_name=‘财务收入明细’, index=False)
return output.getvalue()

st.sidebar.markdown(”### 📋 空白模板”)
st.sidebar.download_button(“下载总账本模板”, data=create_template(), file_name=“地磅总账本_空白.xlsx”)

def parse_excel_date(val):
if pd.isna(val) or str(val).strip()==’’: return ‘’
try:
f=float(val)
if f>30000: return pd.to_datetime(f,unit=‘D’,origin=‘1899-12-30’).strftime(’%Y-%m-%d %H:%M’)
except: pass
return str(val)

# ================= 4. 文件上传 =================

st.subheader(“🚛 地磅管家”)
st.divider()
st.markdown(”#### 第一步：上传业务文件”)
db_file    = st.file_uploader(“📂 1. 上传【总账本.xlsx】”,  type=[‘xlsx’])
daily_file = st.file_uploader(“📝 2. 上传【今日过磅单】”, type=[‘xls’,‘xlsx’,‘csv’])

@st.dialog(“➕ 添加新司机”)
def add_driver_modal(t):
new_d = st.text_input(f”请输入【{t}】的新司机姓名：”, placeholder=“输入姓名”)
if st.button(“确认添加”, type=“primary”, use_container_width=True):
if new_d.strip(): st.session_state[f”custom_drv_{t}”] = new_d.strip()
st.rerun()

# ================= 5. 核心计算 =================

if db_file is not None and daily_file is not None:
try:
df_bal = pd.read_excel(db_file, sheet_name=‘客户余额’)
try:    df_rules = pd.read_excel(db_file, sheet_name=‘加工费规则’)
except: df_rules = pd.DataFrame(columns=[‘物料名称’,‘销售单价’,‘加工费单价’])
try:    df_hist = pd.read_excel(db_file, sheet_name=‘过磅明细’)
except: df_hist = pd.DataFrame()
try:    df_freight = pd.read_excel(db_file, sheet_name=‘公司配送-运费’)
except: df_freight = pd.DataFrame(columns=[‘单号’,‘重车时间’,‘车号’,‘收货单位’,‘货物名称’,‘净重’,‘司机姓名’,‘运费单价’,‘运费金额’])
try:    df_income = pd.read_excel(db_file, sheet_name=‘财务收入明细’)
except: df_income = pd.DataFrame(columns=[‘日期’,‘客户名称’,‘收入类型’,‘金额’,‘备注’])

```
    if daily_file.name.endswith('.csv'): df_daily_raw = pd.read_csv(daily_file)
    else:                                df_daily_raw = pd.read_excel(daily_file)

    if '状态' in df_daily_raw.columns:
        df_daily_raw = df_daily_raw[~df_daily_raw['状态'].astype(str).str.contains('作废|生产', na=False)]

    df_daily_raw['单号']    = df_daily_raw.get('单号','').astype(str).str.replace('.0','',regex=False)
    df_daily_raw['重车时间'] = df_daily_raw.get('重车时间','').apply(parse_excel_date)
    df_daily_raw['净重']    = pd.to_numeric(df_daily_raw.get('净重',0), errors='coerce').fillna(0)
    df_daily_raw['单价']    = pd.to_numeric(df_daily_raw.get('单价',0), errors='coerce').fillna(0)

    orig_amts = pd.to_numeric(df_daily_raw.get('金额',0), errors='coerce').fillna(0)
    types     = df_daily_raw.get('过磅类型','').astype(str)
    new_amts  = []
    for w,p,orig_a,t in zip(df_daily_raw['净重'], df_daily_raw['单价'], orig_amts, types):
        exact = (w*p) if p>0 else orig_a
        if '微信' in t or '现金' in t: new_amts.append(do_round(exact,"retail"))
        elif '签单' in t:              new_amts.append(do_round(exact,"sign"))
        else:                          new_amts.append(do_round(exact,"none"))

    df_daily_raw['金额']    = new_amts
    df_daily_raw['收货单位'] = df_daily_raw.get('收货单位','内部单').fillna('内部单')
    df_daily_raw['货物名称'] = df_daily_raw.get('货物名称','未知物料').fillna('未知物料')
    df_daily_raw['过磅类型'] = types
    df_daily_raw['备注']    = df_daily_raw.get('备注','').fillna('')

    def calc_fee(row):
        if str(row['过磅类型']).strip()=='': return 0.0
        if df_rules.empty: return 0.0
        m = df_rules[(df_rules['物料名称']==row['货物名称']) & (df_rules['销售单价']==row['单价'])]
        if not m.empty: return do_round(row['净重']*m.iloc[0]['加工费单价'],"fee")
        return 0.0

    df_daily_raw['加工费'] = df_daily_raw.apply(calc_fee, axis=1)
    df_sales = df_daily_raw[df_daily_raw['过磅类型'].astype(str).str.strip()!=''].copy()

    all_known_drivers = sorted([str(d) for d in df_freight['司机姓名'].dropna().unique() if str(d).strip() and str(d)!='nan'])
    driver_options = ["(未选择)"] + all_known_drivers + ["➕ 手动输入新司机..."]

    all_known_custs = sorted(list(
        set(df_bal['客户名称'].dropna().astype(str)) |
        (set(df_hist['收货单位'].dropna().astype(str)) if not df_hist.empty else set())
    ))
    all_known_custs = [c for c in all_known_custs if c.strip() and c!='nan' and c!='内部单']
    cust_options = ["(不录入)"] + all_known_custs + ["➕ 手动输入新客户..."]

    # ================== 公司配送 ==================
    has_delivery    = False
    freight_total   = 0.0
    new_freight_records = []
    new_freight_df  = pd.DataFrame()

    st.divider()
    if '备注' in df_sales.columns:
        delivery_mask = df_sales['备注'].astype(str).str.contains('公司配送', na=False)
        if delivery_mask.any():
            has_delivery  = True
            delivery_df   = df_sales[delivery_mask].copy()
            truck_counts  = delivery_df['车号'].value_counts()
            unique_trucks = truck_counts.index.tolist()

            mem_driver, mem_price = {}, {}
            if not df_freight.empty:
                for _, r in df_freight.iterrows():
                    if pd.notna(r.get('车号')) and pd.notna(r.get('司机姓名')):
                        mem_driver[str(r['车号'])] = str(r['司机姓名'])
                    if pd.notna(r.get('收货单位')) and pd.notna(r.get('运费单价')):
                        mem_price[str(r['收货单位'])] = float(r['运费单价'])

            st.markdown(f"#### 🚚 今日有 {len(delivery_df)} 车公司配送，请分配司机与运价")

            # 操作栏：全选 + 批量设置，2列
            cb1, cb2 = st.columns(2)
            with cb1:
                if st.button("🔄 全选/反选", use_container_width=True, key="sel_all"):
                    curr = st.session_state.get("batch_sel", False)
                    st.session_state["batch_sel"] = not curr
                    for t in unique_trucks: st.session_state[f"chk_{t}"] = not curr
                    st.rerun()
            with cb2:
                with st.popover("⚙️ 批量设置运价", use_container_width=True):
                    b_drv = st.selectbox("统一司机", driver_options[:-1], key="batch_drv")
                    b_prc = st.number_input("统一运价", min_value=0.0, format="%.2f", key="batch_prc")
                    if st.button("✅ 应用到勾选车辆", use_container_width=True, key="batch_apply"):
                        for t in unique_trucks:
                            if st.session_state.get(f"chk_{t}", False):
                                if b_drv != "(未选择)": st.session_state[f"custom_drv_{t}"] = b_drv
                                st.session_state[f"p_{t}"] = b_prc
                        st.rerun()

            driver_map, price_map = {}, {}

            def render_truck_row(t):
                count   = truck_counts[t]
                this_df = delivery_df[delivery_df['车号']==t]
                curr_drv = st.session_state.get(f"custom_drv_{t}", mem_driver.get(str(t),""))
                curr_prc = float(st.session_state.get(f"p_{t}", mem_price.get(str(this_df.iloc[0]['收货单位']),0.0)))
                is_done  = bool(curr_drv and curr_prc>0)
                prefix   = "✅" if is_done else ""

                # 4列同行：[勾选] [车牌弹窗] [司机下拉] [运价]
                # 比例 0.6 : 2.8 : 3.3 : 2.3 — 勾选极窄，车牌略宽，司机最宽，运价次之
                c_chk, c_info, c_drv, c_prc = st.columns([0.6, 2.8, 3.3, 2.3])

                with c_chk:
                    st.checkbox("", key=f"chk_{t}", label_visibility="collapsed")

                with c_info:
                    with st.popover(f"{prefix}{t}({count})", use_container_width=True):
                        for _, r in this_df.iterrows():
                            st.markdown(
                                f"📄 **单号**: `{r['单号']}`\n\n"
                                f"🏢 **客户**: {r['收货单位']}\n\n"
                                f"📦 **货物**: {r['货物名称']} | **净重**: `{r['净重']} 吨`\n\n"
                                f"🕒 **时间**: {r['重车时间']}"
                            )
                            st.markdown("---")

                with c_drv:
                    opts = driver_options.copy()
                    if curr_drv and curr_drv not in opts: opts.insert(1, curr_drv)
                    idx = opts.index(curr_drv) if curr_drv in opts else 0
                    d_sel = st.selectbox(f"司机", opts, index=idx, key=f"d_sel_{t}", label_visibility="collapsed")
                    if d_sel == "➕ 手动输入新司机...":
                        add_driver_modal(t)
                    else:
                        driver_map[t] = d_sel if d_sel!="(未选择)" else ""

                with c_prc:
                    p_val = st.number_input("运价", value=curr_prc, min_value=0.0, format="%.1f",
                                            key=f"p_{t}", label_visibility="collapsed")
                    price_map[t] = p_val

            for t in unique_trucks[:4]:
                render_truck_row(t)
            if len(unique_trucks) > 4:
                with st.expander(f"↓ 展开剩余 {len(unique_trucks)-4} 辆车", expanded=False):
                    for t in unique_trucks[4:]:
                        render_truck_row(t)

            for _, row in delivery_df.iterrows():
                t      = row['车号']
                d_name = driver_map.get(t,"")
                p_val  = price_map.get(t, 0.0)
                new_freight_records.append({
                    '单号': row['单号'], '重车时间': row['重车时间'], '车号': row['车号'],
                    '收货单位': row['收货单位'], '货物名称': row['货物名称'], '净重': row['净重'],
                    '司机姓名': d_name, '运费单价': p_val,
                    '运费金额': do_round(row['净重']*p_val, "freight")
                })
            new_freight_df = pd.DataFrame(new_freight_records)
            freight_total  = new_freight_df['运费金额'].sum()

    if not has_delivery:
        with st.expander("📦 额外运费补充（无配送可忽略）", expanded=False):
            freight_total = do_round(st.number_input("额外运费(元):", value=0.0, format="%.2f",
                                                      label_visibility="collapsed"), "freight")

    # ================== 财务与预存 ==================
    st.divider()
    st.markdown("#### 💰 财务资金登记")
    tab_income, tab_deposit = st.tabs(["📥 收入登记", "💳 预存登记"])
    today_income_records = []

    # ── 收入登记：4列同行 [客户] [类型] [金额] [备注] ──
    with tab_income:
        if "income_rows" not in st.session_state: st.session_state.income_rows = 1
        for i in range(st.session_state.income_rows):
            lv = "visible" if i==0 else "collapsed"
            # 比例：客户最宽，备注次之，类型和金额中等
            fc1, fc2, fc3, fc4 = st.columns([2.8, 2.0, 2.0, 2.2])
            with fc1:
                c_sel = st.selectbox("客户", cust_options, key=f"inc_c_sel_{i}", label_visibility=lv)
                if c_sel=="➕ 手动输入新客户...":
                    c_name = st.text_input("新客户", key=f"inc_c_new_{i}", label_visibility="collapsed", placeholder="客户名")
                else:
                    c_name = c_sel if c_sel!="(不录入)" else ""
            with fc2:
                i_type = st.selectbox("类型", ["微信","现金","银行卡","其他"], key=f"inc_t_{i}", label_visibility=lv)
            with fc3:
                i_amt  = st.number_input("金额", min_value=0.0, format="%.0f", key=f"inc_a_{i}", label_visibility=lv)
            with fc4:
                i_note = st.text_input("备注", key=f"inc_n_{i}", label_visibility=lv, placeholder="选填")
            if c_name and i_amt>0:
                today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"),
                    '客户名称': c_name, '收入类型': i_type, '金额': i_amt, '备注': i_note})
        if st.button("➕ 新增收入行", key="add_income_btn", use_container_width=True):
            st.session_state.income_rows += 1
            st.rerun()

    # ── 预存登记：4列同行 [客户] [类型] [金额] [备注] ──
    with tab_deposit:
        if "deposit_rows" not in st.session_state: st.session_state.deposit_rows = 1
        for i in range(st.session_state.deposit_rows):
            lv = "visible" if i==0 else "collapsed"
            fd1, fd2, fd3, fd4 = st.columns([2.8, 2.0, 2.0, 2.2])
            with fd1:
                c_sel = st.selectbox("客户 ", cust_options, key=f"dep_c_sel_{i}", label_visibility=lv)
                if c_sel=="➕ 手动输入新客户...":
                    c_name = st.text_input("新客户", key=f"dep_c_new_{i}", label_visibility="collapsed", placeholder="客户名")
                else:
                    c_name = c_sel if c_sel!="(不录入)" else ""
            with fd2:
                i_type = st.selectbox("类型 ", ["预存微信","银行卡","预存现金","其他"], key=f"dep_t_{i}", label_visibility=lv)
            with fd3:
                i_amt  = st.number_input("金额 ", min_value=0.0, format="%.0f", key=f"dep_a_{i}", label_visibility=lv)
            with fd4:
                i_note = st.text_input("备注 ", key=f"dep_n_{i}", label_visibility=lv, placeholder="选填")
            if c_name and i_amt>0:
                today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"),
                    '客户名称': c_name, '收入类型': i_type, '金额': i_amt, '备注': i_note})
        if st.button("➕ 新增预存行", key="add_deposit_btn", use_container_width=True):
            st.session_state.deposit_rows += 1
            st.rerun()

    # ---- 计算逻辑 ----
    df_cash  = df_sales[df_sales['过磅类型'].astype(str).str.contains('现金', na=False)]
    df_wx    = df_sales[df_sales['过磅类型'].astype(str).str.contains('微信', na=False)]
    df_sign  = df_sales[df_sales['过磅类型'].astype(str).str.contains('签单', na=False)]

    orig_bal_dict = dict(zip(df_bal['客户名称'], df_bal['余额']))
    deposit_dict  = {}

    retail_wx_amt   = do_round(df_wx['金额'].sum(),   "retail")
    retail_cash_amt = do_round(df_cash['金额'].sum(), "retail")
    if retail_wx_amt   > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售微信', '金额': retail_wx_amt,   '备注': '自动汇总'})
    if retail_cash_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售现金', '金额': retail_cash_amt, '备注': '自动汇总'})

    for r in today_income_records:
        if "预存" in r['收入类型'] or r['收入类型']=="银行卡":
            if r['客户名称']!='零售客户':
                deposit_dict[r['客户名称']] = deposit_dict.get(r['客户名称'],0.0) + r['金额']

    bal_dict   = {}
    sign_custs = df_sign.groupby('收货单位')
    all_custs  = set(list(orig_bal_dict.keys()) + list(deposit_dict.keys()) + list(sign_custs.groups.keys()))
    for c in all_custs:
        spent = do_round(df_sign[df_sign['收货单位']==c]['金额'].sum(),"sign") if c in sign_custs.groups else 0.0
        bal_dict[c] = orig_bal_dict.get(c,0.0) + deposit_dict.get(c,0.0) - spent

    daily_fee     = do_round(df_sales['加工费'].sum(),"fee")
    current_month = report_date.strftime("%Y-%m")
    monthly_fee   = daily_fee
    if not df_hist.empty and '重车时间' in df_hist.columns and '加工费' in df_hist.columns:
        df_hist['加工费'] = pd.to_numeric(df_hist['加工费'], errors='coerce').fillna(0)
        monthly_fee += do_round(df_hist[df_hist['重车时间'].astype(str).str.startswith(current_month,na=False)]['加工费'].sum(),"fee")

    # ================= 汇报文本 =================
    report = f"{report_date.strftime('%y年%m月%d日')} {report_time}\n"

    report += f"\n现金:{len(df_cash)}车{fmt_weight(df_cash['净重'].sum())}吨{fmt_val(retail_cash_amt,'retail')}元\n"
    for prod,grp in df_cash.groupby('货物名称'):
        report += f"{prod}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨{fmt_val(grp['金额'].sum(),'retail')}元\n"

    report += f"\n微信:{len(df_wx)}车{fmt_weight(df_wx['净重'].sum())}吨{fmt_val(retail_wx_amt,'retail')}元\n"
    for prod,grp in df_wx.groupby('货物名称'):
        report += f"{prod}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨{fmt_val(grp['金额'].sum(),'retail')}元\n"

    report += f"\n签单:{len(df_sign)}车{fmt_weight(df_sign['净重'].sum())}吨\n"
    for cust,grp in sign_custs:
        report += f"{cust}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨\n"
        for prod,p_grp in grp.groupby('货物名称'):
            report += f"{prod}:{len(p_grp)}车{fmt_weight(p_grp['净重'].sum())}吨{fmt_val(p_grp['金额'].sum(),'sign')} 元\n"
        cust_money = do_round(grp['金额'].sum(),"sign")
        report += f"共金额:{fmt_val(cust_money,'sign')} 元\n"
        report += f"上日余额:{fmt_val(orig_bal_dict.get(cust,0.0),'sign')} 元\n"
        if deposit_dict.get(cust,0.0)>0: report += f"今日充值:{fmt_val(deposit_dict.get(cust,0.0),'sign')} 元\n"
        report += f"当日余额:{fmt_val(bal_dict.get(cust,0.0),'sign')} 元\n"

    pure_depositors = [c for c in deposit_dict if c not in sign_custs.groups]
    if pure_depositors:
        report += "\n【纯充值客户余额刷新】\n"
        for c in pure_depositors:
            report += f"{c} 今日充值:{fmt_val(deposit_dict.get(c,0.0),'sign')}元 | 最新余额:{fmt_val(bal_dict.get(c,0.0),'sign')}元\n"

    total_money  = df_sales['金额'].sum()
    unsold_count = len(df_daily_raw) - len(df_sales)
    unsold_str   = f" (内含未销售单据 {unsold_count} 车，已全量留底)" if unsold_count>0 else ""

    report += f"\n1,当日销售共计:{len(df_sales)} 车{fmt_weight(df_sales['净重'].sum())} 吨 {fmt_val(total_money,'sign')} 元,公司配送运费:{fmt_val(freight_total,'freight')}元 ,合计:{fmt_val(total_money+freight_total,'sign')} 元。{unsold_str}\n"
    report += f"2,当日加工费:{fmt_val(daily_fee,'fee')} 元,{report_date.month}月1日-{report_date.day}日加工费合计:{fmt_val(monthly_fee,'fee')} 元。\n"

    collection_parts      = []
    custom_income_total   = 0.0
    if retail_wx_amt   > 0: collection_parts.append(f"微信零售:{fmt_val(retail_wx_amt,'retail')}元")
    if retail_cash_amt > 0: collection_parts.append(f"现金零售:{fmt_val(retail_cash_amt,'retail')}元")
    for r in today_income_records:
        if r['客户名称']!='零售客户':
            custom_income_total += float(r['金额'])
            label = r['客户名称'] or ''
            collection_parts.append(f"{label}{r['收入类型']}:{fmt_val(r['金额'],'none')}元")

    total_collection = retail_wx_amt + retail_cash_amt + custom_income_total
    report += f"3,当日合计收款:{','.join(collection_parts) if collection_parts else '0元'},共计:{fmt_val(total_collection,'none')} 元\n"

    st.divider()
    st.markdown("#### 第二步：复制每日汇报")
    st.code(report, language="text", line_numbers=False)

    # ---- 组装新账本 ----
    new_df_bal = pd.DataFrame(list(bal_dict.items()), columns=['客户名称','余额'])
    cols_keep  = ['单号','重车时间','车号','货物名称','净重','单价','金额','收货单位','过磅类型','备注','加工费']
    avail_cols = [c for c in cols_keep if c in df_daily_raw.columns]

    new_df_hist    = safe_concat([df_hist, df_daily_raw[avail_cols]])
    new_df_freight = safe_concat([df_freight, new_freight_df])
    today_income_df = pd.DataFrame(today_income_records) if today_income_records else pd.DataFrame()
    new_df_income  = safe_concat([df_income, today_income_df])

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        new_df_bal.to_excel(writer,     sheet_name='客户余额',    index=False)
        df_rules.to_excel(writer,       sheet_name='加工费规则',   index=False)
        new_df_hist.to_excel(writer,    sheet_name='过磅明细',    index=False)
        new_df_freight.to_excel(writer, sheet_name='公司配送-运费', index=False)
        new_df_income.to_excel(writer,  sheet_name='财务收入明细', index=False)

    st.success("✅ 核算完成，可下载更新后的总账本", icon="✅")
    st.download_button(
        label="💾 下载更新后总账本",
        data=out.getvalue(),
        file_name=f"{report_date.strftime('%Y%m%d')}_DiBang总账本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

except Exception as e:
    st.error(f"❌ 文件处理失败，请检查：\n1. 文件格式是否正确；\n2. 总账本是否符合模板要求。\n\n错误详情：{str(e)}", icon="⚠️")
```
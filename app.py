import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="地磅管家 Pro", page_icon="🚛", layout="centered")

# ================= 0. 左侧边栏：全局设置与排版引擎 =================
st.sidebar.header("📅 报表时间设置")
report_date = st.sidebar.date_input("报表日期", value=datetime.now())
report_time = st.sidebar.text_input("时间段", value="07:00-18:00")

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 四舍五入设置 (抹零取整)")
st.sidebar.write("勾选后取整数；不勾选则严格保留两位小数 (.00)")

round_retail = st.sidebar.checkbox("🛒 零售 (微信/现金)", value=True)
round_sign = st.sidebar.checkbox("🏢 签单客户", value=False)
round_fee = st.sidebar.checkbox("⚙️ 加工费", value=False)
round_freight = st.sidebar.checkbox("🚚 运费", value=False)

def do_round(val, category=""):
    """数值计算时的四舍五入引擎"""
    if pd.isna(val): return 0.0
    val = float(val)
    should_round = False
    if category == "retail" and round_retail: should_round = True
    elif category == "sign" and round_sign: should_round = True
    elif category == "fee" and round_fee: should_round = True
    elif category == "freight" and round_freight: should_round = True
    
    if should_round: return float(int(val + 0.5) if val >= 0 else int(val - 0.5))
    else: return round(val + 1e-9, 2)

def fmt_val(v, category=""):
    """金额排版引擎：严格遵循开关，关掉时必定保留两位小数"""
    if pd.isna(v): return "0.00"
    v = float(v)
    should_round = False
    if category == "retail" and round_retail: should_round = True
    elif category == "sign" and round_sign: should_round = True
    elif category == "fee" and round_fee: should_round = True
    elif category == "freight" and round_freight: should_round = True
    
    if should_round: return f"{int(v + 0.5) if v >= 0 else int(v - 0.5)}"
    else: return f"{v:.2f}"

def fmt_weight(v):
    """吨数排版引擎：智能去掉没用的小数"""
    if pd.isna(v): return "0"
    v = float(v)
    if abs(v - round(v)) < 1e-5: return f"{int(round(v))}"
    return f"{v:.2f}"

# ================= 1. 初始化空账本 =================
def create_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=['客户名称', '余额']).to_excel(writer, sheet_name='客户余额', index=False)
        pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价']).to_excel(writer, sheet_name='加工费规则', index=False)
        pd.DataFrame(columns=['单号', '重车时间', '车号', '货物名称', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费']).to_excel(writer, sheet_name='过磅明细', index=False)
        pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额']).to_excel(writer, sheet_name='公司配送-运费', index=False)
        pd.DataFrame(columns=['日期', '客户名称', '收入类型', '金额', '备注']).to_excel(writer, sheet_name='财务收入明细', index=False)
    return output.getvalue()

st.sidebar.markdown("---")
st.sidebar.download_button("⬇️ 下载【5表合一_空白总账本】", data=create_template(), file_name="地磅总账本_空白.xlsx")

def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() == '': return ''
    try:
        f_val = float(val)
        if f_val > 30000: return pd.to_datetime(f_val, unit='D', origin='1899-12-30').strftime('%Y-%m-%d %H:%M')
    except: pass
    return str(val)

# ================= 2. 主界面：文件上传 =================
st.title("🚛 地磅管家")
st.write("### 第一步：上传文件")

col1, col2 = st.columns(2)
with col1: db_file = st.file_uploader("📂 1. 上传【总账本.xlsx】", type=['xlsx'])
with col2: daily_file = st.file_uploader("📝 2. 上传【今日单】", type=['xls', 'xlsx', 'csv'])

# ================= 3. 核心计算 =================
if db_file is not None and daily_file is not None:
    try:
        df_bal = pd.read_excel(db_file, sheet_name='客户余额')
        try: df_rules = pd.read_excel(db_file, sheet_name='加工费规则')
        except: df_rules = pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价'])
        try: df_hist = pd.read_excel(db_file, sheet_name='过磅明细')
        except: df_hist = pd.DataFrame()
        try: df_freight = pd.read_excel(db_file, sheet_name='公司配送-运费')
        except: df_freight = pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额'])
        try: df_income = pd.read_excel(db_file, sheet_name='财务收入明细')
        except: df_income = pd.DataFrame(columns=['日期', '客户名称', '收入类型', '金额', '备注'])
            
        if daily_file.name.endswith('.csv'): df_daily_raw = pd.read_csv(daily_file)
        else: df_daily_raw = pd.read_excel(daily_file)
            
        if '状态' in df_daily_raw.columns:
            df_daily_raw = df_daily_raw[~df_daily_raw['状态'].astype(str).str.contains('作废|生产', na=False)]
            
        df_daily_raw['单号'] = df_daily_raw.get('单号', '').astype(str).str.replace('.0', '', regex=False)
        df_daily_raw['重车时间'] = df_daily_raw.get('重车时间', '').apply(parse_excel_date)
        df_daily_raw['净重'] = pd.to_numeric(df_daily_raw.get('净重', 0), errors='coerce').fillna(0)
        df_daily_raw['单价'] = pd.to_numeric(df_daily_raw.get('单价', 0), errors='coerce').fillna(0)
        
        raw_amts = pd.to_numeric(df_daily_raw.get('金额', 0), errors='coerce').fillna(0)
        types = df_daily_raw.get('过磅类型', '').astype(str)
        new_amts = []
        for amt, t in zip(raw_amts, types):
            if '微信' in t or '现金' in t: new_amts.append(do_round(amt, "retail"))
            elif '签单' in t: new_amts.append(do_round(amt, "sign"))
            else: new_amts.append(do_round(amt, "none"))
        df_daily_raw['金额'] = new_amts
        
        df_daily_raw['收货单位'] = df_daily_raw.get('收货单位', '内部单').fillna('内部单')
        df_daily_raw['货物名称'] = df_daily_raw.get('货物名称', '未知物料').fillna('未知物料')
        df_daily_raw['过磅类型'] = types
        df_daily_raw['备注'] = df_daily_raw.get('备注', '').fillna('')
        
        def calc_fee(row):
            if str(row['过磅类型']).strip() == '': return 0.0 
            if df_rules.empty: return 0.0
            match = df_rules[(df_rules['物料名称'] == row['货物名称']) & (df_rules['销售单价'] == row['单价'])]
            if not match.empty: return do_round(row['净重'] * match.iloc[0]['加工费单价'], "fee")
            return 0.0
            
        df_daily_raw['加工费'] = df_daily_raw.apply(calc_fee, axis=1)
        df_sales = df_daily_raw[df_daily_raw['过磅类型'].astype(str).str.strip() != ''].copy()
        
        # ---------------- 物流运费逻辑 (智能下拉框) ----------------
        has_delivery = False
        freight_total = 0.0
        new_freight_records = []
        
        all_known_drivers = sorted([str(d) for d in df_freight['司机姓名'].dropna().unique() if str(d).strip() and str(d)!='nan'])
        driver_options = ["(未选择)"] + all_known_drivers + ["➕ 手动输入新司机..."]
        
        if '备注' in df_sales.columns:
            delivery_mask = df_sales['备注'].astype(str).str.contains('公司配送', na=False)
            if delivery_mask.any():
                has_delivery = True
                delivery_df = df_sales[delivery_mask].copy()
                unique_trucks = delivery_df['车号'].dropna().unique()
                unique_custs = delivery_df['收货单位'].dropna().unique()
                
                mem_driver, mem_price = {}, {}
                if not df_freight.empty:
                    for _, r in df_freight.iterrows():
                        if pd.notna(r.get('车号')) and pd.notna(r.get('司机姓名')): mem_driver[str(r['车号'])] = str(r['司机姓名'])
                        if pd.notna(r.get('收货单位')) and pd.notna(r.get('运费单价')): mem_price[str(r['收货单位'])] = float(r['运费单价'])
                
                st.warning(f"🚚 今日有 {len(delivery_df)} 车【公司配送】。请分配司机与运价：")
                col_d, col_p = st.columns(2)
                driver_map, price_map = {}, {}
                with col_d:
                    st.success("👤 设定司机 (下拉选择或新增)")
                    for t in unique_trucks:
                        mem_d = mem_driver.get(str(t), "(未选择)")
                        idx = driver_options.index(mem_d) if mem_d in driver_options else 0
                        d_sel = st.selectbox(f"车号 [{t}]", driver_options, index=idx, key=f"d_sel_{t}")
                        if d_sel == "➕ 手动输入新司机...": driver_map[t] = st.text_input(f"手动输入 [{t}] 司机", key=f"d_new_{t}")
                        else: driver_map[t] = d_sel if d_sel != "(未选择)" else ""
                with col_p:
                    st.success("💰 设定运价")
                    for c in unique_custs: price_map[c] = st.number_input(f"客户 [{c}]", value=mem_price.get(str(c), 0.0), step=1.0, format="%.2f", key=f"p_{c}")
                
                for idx, row in delivery_df.iterrows():
                    d_name = driver_map.get(row['车号'], "")
                    p_val = price_map.get(row['收货单位'], 0.0)
                    new_freight_records.append({
                        '单号': row['单号'], '重车时间': row['重车时间'], '车号': row['车号'],
                        '收货单位': row['收货单位'], '货物名称': row['货物名称'], '净重': row['净重'],
                        '司机姓名': d_name, '运费单价': p_val, '运费金额': do_round(row['净重'] * p_val, "freight")
                    })
                
                new_freight_df = pd.DataFrame(new_freight_records)
                freight_total = new_freight_df['运费金额'].sum()

        if not has_delivery:
            f_val = st.number_input("🚚 今日未检测到配送。若有额外运费(元):", value=0.0, step=10.0)
            freight_total = do_round(f_val, "freight")
            new_freight_df = pd.DataFrame()

        # ================== 财务预存登记 (智能下拉框) ==================
        st.write("### 💰 财务收入与预存登记 (无需录入请留空)")
        all_known_custs = sorted(list(set(df_bal['客户名称'].dropna().astype(str)) | set(df_hist['收货单位'].dropna().astype(str))))
        all_known_custs = [c for c in all_known_custs if c.strip() and c != 'nan' and c != '内部单']
        cust_options = ["(不录入)"] + all_known_custs + ["➕ 手动输入新客户..."]
        
        today_income_records = []
        for i in range(3):
            c1, c2, c3, c4 = st.columns([3, 2, 2, 3])
            with c1:
                c_sel = st.selectbox(f"客户名称 {i+1}", cust_options, key=f"c_sel_{i}", label_visibility="collapsed" if i>0 else "visible")
                if c_sel == "➕ 手动输入新客户...": c_name = st.text_input(f"新客户 {i+1}", key=f"c_new_{i}", label_visibility="collapsed")
                else: c_name = c_sel if c_sel != "(不录入)" else ""
            with c2: i_type = st.selectbox(f"收入类型 {i+1}", ["预存微信", "银行卡", "预存现金", "备用金", "其他"], key=f"t_{i}", label_visibility="collapsed" if i>0 else "visible")
            with c3: i_amt = st.number_input(f"金额(元) {i+1}", min_value=0.0, step=100.0, format="%.2f", key=f"a_{i}", label_visibility="collapsed" if i>0 else "visible")
            with c4: i_note = st.text_input(f"备注 {i+1}", key=f"n_{i}", label_visibility="collapsed" if i>0 else "visible")
                
            if c_name and i_amt > 0:
                today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': c_name, '收入类型': i_type, '金额': i_amt, '备注': i_note})

        df_cash = df_sales[df_sales['过磅类型'].astype(str).str.contains('现金', na=False)]
        df_wx = df_sales[df_sales['过磅类型'].astype(str).str.contains('微信', na=False)]
        df_sign = df_sales[df_sales['过磅类型'].astype(str).str.contains('签单', na=False)]

        # -------- 计算资金与余额 --------
        orig_bal_dict = dict(zip(df_bal['客户名称'], df_bal['余额']))
        deposit_dict = {}

        retail_wx_amt = do_round(df_wx['金额'].sum(), "retail")
        retail_cash_amt = do_round(df_cash['金额'].sum(), "retail")
        if retail_wx_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售微信', '金额': retail_wx_amt, '备注': '自动汇总'})
        if retail_cash_amt > 0: today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售现金', '金额': retail_cash_amt, '备注': '自动汇总'})

        for r in today_income_records:
            if r['收入类型'] in ["预存微信", "银行卡", "预存现金"] and r['客户名称'] != '零售客户':
                deposit_dict[r['客户名称']] = deposit_dict.get(r['客户名称'], 0.0) + r['金额']

        bal_dict = {}
        sign_custs = df_sign.groupby('收货单位')
        all_custs = set(list(orig_bal_dict.keys()) + list(deposit_dict.keys()) + list(sign_custs.groups.keys()))
        
        for c in all_custs:
            spent = do_round(df_sign[df_sign['收货单位'] == c]['金额'].sum(), "sign") if c in sign_custs.groups else 0.0
            bal_dict[c] = orig_bal_dict.get(c, 0.0) + deposit_dict.get(c, 0.0) - spent

        daily_fee = do_round(df_sales['加工费'].sum(), "fee")
        current_month = report_date.strftime("%Y-%m")
        monthly_fee = daily_fee
        if not df_hist.empty and '重车时间' in df_hist.columns and '加工费' in df_hist.columns:
            df_hist['加工费'] = pd.to_numeric(df_hist['加工费'], errors='coerce').fillna(0)
            monthly_fee += do_round(df_hist[df_hist['重车时间'].astype(str).str.startswith(current_month, na=False)]['加工费'].sum(), "fee")
        
        # ---------------- 生成完美格式的汇报单 ----------------
        report = f"{report_date.strftime('%y年%m月%d日')} {report_time}\n\n"
        
        report += "【今日资金收入】\n"
        if retail_wx_amt > 0: report += f"零售微信: {fmt_val(retail_wx_amt, 'retail')} 元\n"
        if retail_cash_amt > 0: report += f"零售现金: {fmt_val(retail_cash_amt, 'retail')} 元\n"
        for r in today_income_records:
            if r['客户名称'] != '零售客户':
                prefix = f"{r['客户名称']} " if r['客户名称'] else ""
                report += f"{prefix}{r['收入类型']}: {fmt_val(r['金额'], 'none')} 元\n"
        report += "\n"
        
        report += f"现金:{len(df_cash)}车{fmt_weight(df_cash['净重'].sum())}吨{fmt_val(retail_cash_amt, 'retail')}元\n"
        for prod, grp in df_cash.groupby('货物名称'):
            report += f"{prod}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨{fmt_val(grp['金额'].sum(), 'retail')}元\n"
        report += "\n"
        
        report += f"微信:{len(df_wx)}车{fmt_weight(df_wx['净重'].sum())}吨{fmt_val(retail_wx_amt, 'retail')}元\n"
        for prod, grp in df_wx.groupby('货物名称'):
            report += f"{prod}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨{fmt_val(grp['金额'].sum(), 'retail')}元\n"
            
        report += f"\n签单:{len(df_sign)}车{fmt_weight(df_sign['净重'].sum())}吨\n\n"
        
        for cust, grp in sign_custs:
            report += f"{cust}:{len(grp)}车{fmt_weight(grp['净重'].sum())}吨\n"
            for prod, p_grp in grp.groupby('货物名称'):
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
                report += f"{c} 今日充值:{fmt_val(deposit_dict.get(c, 0.0), 'sign')}元 | 最新余额:{fmt_val(bal_dict.get(c, 0.0), 'sign')}元\n"
            report += "\n"
            
        total_money = df_sales['金额'].sum()
        unsold_count = len(df_daily_raw) - len(df_sales)
        unsold_str = f" (内含未销售单据 {unsold_count} 车，已全量留底)" if unsold_count > 0 else ""
        
        report += f"1,当日销售共计:{len(df_sales)} 车{fmt_weight(df_sales['净重'].sum())} 吨 {fmt_val(total_money, 'sign')} 元,公司配送运费:{fmt_val(freight_total, 'freight')}元 ,合计:{fmt_val(total_money + freight_total, 'sign')} 元。{unsold_str}\n"
        report += f"2,当日加工费:{fmt_val(daily_fee, 'fee')} 元,{report_date.month}月1日-{report_date.day}日加工费合计:{fmt_val(monthly_fee, 'fee')} 元。\n"
        report += f"3,当日合计收款:微信零售:{fmt_val(retail_wx_amt, 'retail')} 元,共计:{fmt_val(retail_wx_amt + retail_cash_amt, 'retail')} 元\n"

        st.write("### 第二步：复制汇报单")
        st.code(report, language="text")
        
        # ---------------- 组装新账本 ----------------
        st.write("### 第三步：下载更新后的总账本")
        new_df_bal = pd.DataFrame(list(bal_dict.items()), columns=['客户名称', '余额'])
        
        cols_to_keep = ['单号', '重车时间', '车号', '货物名称', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费']
        available_cols = [c for c in cols_to_keep if c in df_daily_raw.columns]
        new_df_hist = pd.concat([df_hist, df_daily_raw[available_cols]], ignore_index=True)
        
        new_df_freight = pd.concat([df_freight, new_freight_df], ignore_index=True) if not new_freight_df.empty else df_freight
        new_df_income = pd.concat([df_income, pd.DataFrame(today_income_records)], ignore_index=True) if today_income_records else df_income
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            new_df_bal.to_excel(writer, sheet_name='客户余额', index=False)
            df_rules.to_excel(writer, sheet_name='加工费规则', index=False)
            new_df_hist.to_excel(writer, sheet_name='过磅明细', index=False)
            new_df_freight.to_excel(writer, sheet_name='公司配送-运费', index=False)
            new_df_income.to_excel(writer, sheet_name='财务收入明细', index=False)
            
        st.success("✅ 核算完毕！页面更清爽，财务客户与物流司机支持智能下拉手写！")
        st.download_button(
            label=f"💾 下载【{report_date.strftime('%Y%m%d')}_DiBang总账本.xlsx】",
            data=output.getvalue(),
            file_name=f"{report_date.strftime('%Y%m%d')}_DiBang总账本.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
            
    except Exception as e:
        st.error(f"处理出错，请确保上传的文件正确。错误信息: {e}")

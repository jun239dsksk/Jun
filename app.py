import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="地磅管家 Pro", page_icon="🚛", layout="centered")

st.title("🚛 地磅管家 (业财一体联动版)")

# ================= 0. 全局设置与格式化引擎 =================
st.sidebar.header("⚙️ 四舍五入设置 (抹零取整)")
st.sidebar.write("勾选后，该项金额将自动四舍五入到元；不勾选则保留两位小数。")

round_retail = st.sidebar.checkbox("🛒 零售 (微信/现金)", value=True)
round_sign = st.sidebar.checkbox("🏢 签单客户", value=False)
round_fee = st.sidebar.checkbox("⚙️ 加工费", value=False)
round_freight = st.sidebar.checkbox("🚚 运费", value=False)

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

def fmt(v):
    if pd.isna(v): return "0"
    v = float(v)
    if abs(v - round(v)) < 1e-5: return f"{int(round(v))}"
    return f"{v:.2f}"

# ================= 1. 初始化空账本 (新增财务表) =================
def create_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=['客户名称', '余额']).to_excel(writer, sheet_name='客户余额', index=False)
        pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价']).to_excel(writer, sheet_name='加工费规则', index=False)
        pd.DataFrame(columns=['单号', '重车时间', '车号', '货物名称', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费']).to_excel(writer, sheet_name='过磅明细', index=False)
        pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额']).to_excel(writer, sheet_name='公司配送-运费', index=False)
        # 新增第五张表：财务收入明细
        pd.DataFrame(columns=['日期', '客户名称', '收入类型', '金额', '备注']).to_excel(writer, sheet_name='财务收入明细', index=False)
    return output.getvalue()

st.sidebar.markdown("---")
st.sidebar.header("🛠️ 首次使用向导")
st.sidebar.download_button(
    label="⬇️ 下载【5表合一_空白总账本.xlsx】",
    data=create_template(),
    file_name="地磅总账本_空白.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="包含：客户余额、加工费规则、历史明细、公司配送运费、财务收入明细 五个工作表"
)

def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() == '': return ''
    try:
        f_val = float(val)
        if f_val > 30000: return pd.to_datetime(f_val, unit='D', origin='1899-12-30').strftime('%Y-%m-%d %H:%M')
    except: pass
    return str(val)

# ================= 2. 文件上传与参数配置 =================
st.write("### 第一步：参数与文件上传")

col_date, col_time = st.columns(2)
with col_date: report_date = st.date_input("📅 报表日期", value=datetime.now())
with col_time: report_time = st.text_input("🕒 报表时间段", value="07:00-18:00")

col1, col2 = st.columns(2)
with col1: db_file = st.file_uploader("📂 1. 上传【总账本.xlsx】", type=['xlsx'])
with col2: daily_file = st.file_uploader("📝 2. 上传【今日单】", type=['xls', 'xlsx', 'csv'])

# ================= 3. 核心计算 =================
if db_file is not None and daily_file is not None:
    try:
        # 读取总账本 (兼容新老版本)
        df_bal = pd.read_excel(db_file, sheet_name='客户余额')
        try: df_rules = pd.read_excel(db_file, sheet_name='加工费规则')
        except: df_rules = pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价'])
        try: df_hist = pd.read_excel(db_file, sheet_name='过磅明细')
        except: df_hist = pd.DataFrame()
        try: df_freight = pd.read_excel(db_file, sheet_name='公司配送-运费')
        except: df_freight = pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额'])
        try: df_income = pd.read_excel(db_file, sheet_name='财务收入明细')
        except: df_income = pd.DataFrame(columns=['日期', '客户名称', '收入类型', '金额', '备注'])
            
        # 读取今日单
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

        # 核心分离：分离出纯销售单
        df_sales = df_daily_raw[df_daily_raw['过磅类型'].astype(str).str.strip() != ''].copy()
        
        # ---------------- 物流运费逻辑 ----------------
        has_delivery = False
        freight_total = 0.0
        new_freight_records = []
        
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
                
                st.warning(f"🚚 今日有 {len(delivery_df)} 车【公司配送】。系统已合并并读取历史习惯！")
                col_d, col_p = st.columns(2)
                driver_map, price_map = {}, {}
                with col_d:
                    st.success("👤 设定司机")
                    for t in unique_trucks: driver_map[t] = st.text_input(f"[{t}]", value=mem_driver.get(str(t), ""), key=f"d_{t}")
                with col_p:
                    st.success("💰 设定运价")
                    for c in unique_custs: price_map[c] = st.number_input(f"[{c}]", value=mem_price.get(str(c), 0.0), step=1.0, format="%.2f", key=f"p_{c}")
                
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

        # ================== 财务与充值中心 ==================
        st.write("### 💰 财务收入与预存登记")
        st.info("系统会自动将今日零售汇总计入收入表。如果您今天收到了客户预存或备用金，请在下方点击加号添加：")
        
        fin_init = pd.DataFrame({'客户名称': [''], '收入类型': ['预存微信'], '金额': [0.0], '备注': ['']})
        edited_fin = st.data_editor(
            fin_init,
            column_config={
                "客户名称": st.column_config.TextColumn("客户名称 (预存必填，需与过磅单一致)"),
                "收入类型": st.column_config.SelectboxColumn("收入类型", options=["预存微信", "银行卡", "预存现金", "备用金", "其他"], required=True),
                "金额": st.column_config.NumberColumn("金额 (元)", min_value=0.0, format="%.2f"),
                "备注": st.column_config.TextColumn("备注说明")
            },
            num_rows="dynamic",
            use_container_width=True
        )

        df_cash = df_sales[df_sales['过磅类型'].astype(str).str.contains('现金', na=False)]
        df_wx = df_sales[df_sales['过磅类型'].astype(str).str.contains('微信', na=False)]
        df_sign = df_sales[df_sales['过磅类型'].astype(str).str.contains('签单', na=False)]

        # -------- 计算资金与余额 --------
        orig_bal_dict = dict(zip(df_bal['客户名称'], df_bal['余额']))
        deposit_dict = {}
        today_income_records = []

        # 1. 自动记账：零售微信与现金
        retail_wx_amt = df_wx['金额'].sum()
        retail_cash_amt = df_cash['金额'].sum()
        if retail_wx_amt > 0:
            today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售微信', '金额': retail_wx_amt, '备注': '系统自动汇总'})
        if retail_cash_amt > 0:
            today_income_records.append({'日期': report_date.strftime("%Y-%m-%d"), '客户名称': '零售客户', '收入类型': '零售现金', '金额': retail_cash_amt, '备注': '系统自动汇总'})

        # 2. 手工记账：读取刚才表格填写的记录
        for idx, row in edited_fin.iterrows():
            amt = float(row.get('金额') or 0.0)
            if amt > 0:
                c_name = str(row.get('客户名称', '')).strip()
                if c_name.lower() == 'nan': c_name = ""
                i_type = str(row.get('收入类型', '')).strip()
                
                # 如果是客户预存，记录到存款字典中
                if i_type in ["预存微信", "银行卡", "预存现金"] and c_name:
                    deposit_dict[c_name] = deposit_dict.get(c_name, 0.0) + amt
                    
                today_income_records.append({
                    '日期': report_date.strftime("%Y-%m-%d"),
                    '客户名称': c_name, '收入类型': i_type, '金额': amt, '备注': str(row.get('备注', ''))
                })

        # 3. 终极余额结算 = 历史余额 + 今日预存 - 今日消费
        bal_dict = {}
        sign_custs = df_sign.groupby('收货单位')
        all_custs = set(list(orig_bal_dict.keys()) + list(deposit_dict.keys()) + list(sign_custs.groups.keys()))
        
        for c in all_custs:
            orig = orig_bal_dict.get(c, 0.0)
            dep = deposit_dict.get(c, 0.0)
            spent = df_sign[df_sign['收货单位'] == c]['金额'].sum() if c in sign_custs.groups else 0.0
            bal_dict[c] = orig + dep - spent

        # -------- 计算加工费 --------
        daily_fee = df_sales['加工费'].sum()
        current_month = report_date.strftime("%Y-%m")
        monthly_fee = daily_fee
        if not df_hist.empty and '重车时间' in df_hist.columns and '加工费' in df_hist.columns:
            df_hist['加工费'] = pd.to_numeric(df_hist['加工费'], errors='coerce').fillna(0)
            hist_this_month = df_hist[df_hist['重车时间'].astype(str).str.startswith(current_month, na=False)]
            monthly_fee += hist_this_month['加工费'].sum()
        
        # ---------------- 生成完美格式的汇报单 ----------------
        report = f"{report_date.strftime('%y年%m月%d日')} {report_time}\n\n"
        
        # 【新增】第一板块：今日资金流转明细
        report += "【今日资金收入】\n"
        if retail_wx_amt > 0: report += f"零售微信: {fmt(retail_wx_amt)} 元\n"
        if retail_cash_amt > 0: report += f"零售现金: {fmt(retail_cash_amt)} 元\n"
        for idx, row in edited_fin.iterrows():
            amt = float(row.get('金额', 0.0))
            if amt > 0:
                c_name = str(row.get('客户名称', '')).strip()
                if c_name.lower() == 'nan': c_name = ""
                i_type = str(row.get('收入类型', ''))
                prefix = f"{c_name} " if c_name else ""
                report += f"{prefix}{i_type}: {fmt(amt)} 元\n"
        report += "\n"
        
        report += f"现金:{len(df_cash)}车{fmt(df_cash['净重'].sum())}吨{fmt(df_cash['金额'].sum())}元\n"
        cash_grp = df_cash.groupby('货物名称')[['净重', '金额']].sum()
        cash_cnt = df_cash.groupby('货物名称').size()
        for prod in cash_grp.index:
            report += f"{prod}:{cash_cnt[prod]}车{fmt(cash_grp.loc[prod, '净重'])}吨{fmt(cash_grp.loc[prod, '金额'])}元\n"
        report += "\n"
        
        report += f"微信:{len(df_wx)}车{fmt(df_wx['净重'].sum())}吨{fmt(df_wx['金额'].sum())}元\n"
        wx_grp = df_wx.groupby('货物名称')[['净重', '金额']].sum()
        wx_cnt = df_wx.groupby('货物名称').size()
        for prod in wx_grp.index:
            report += f"{prod}:{wx_cnt[prod]}车{fmt(wx_grp.loc[prod, '净重'])}吨{fmt(wx_grp.loc[prod, '金额'])}元\n"
            
        report += f"\n签单:{len(df_sign)}车{fmt(df_sign['净重'].sum())}吨\n\n"
        
        for cust, grp in sign_custs:
            report += f"{cust}:{len(grp)}车{fmt(grp['净重'].sum())}吨\n"
            prod_grp = grp.groupby('货物名称')[['净重', '金额']].sum()
            prod_cnt = grp.groupby('货物名称').size()
            for prod in prod_grp.index:
                report += f"{prod}:{prod_cnt[prod]}车{fmt(prod_grp.loc[prod, '净重'])}吨{fmt(prod_grp.loc[prod, '金额'])} 元\n"
            
            cust_money = grp['金额'].sum()
            orig = orig_bal_dict.get(cust, 0.0)
            dep = deposit_dict.get(cust, 0.0)
            curr = bal_dict.get(cust, 0.0)
            
            report += f"共金额:{fmt(cust_money)} 元\n"
            report += f"上日余额:{fmt(orig)} 元\n"
            if dep > 0: report += f"今日充值:{fmt(dep)} 元\n"
            report += f"当日余额:{fmt(curr)} 元\n\n"
            
        # 【新增】检测是否有客户今天“只充值，没拉货”
        pure_depositors = [c for c in deposit_dict.keys() if c not in sign_custs.groups]
        if pure_depositors:
            report += "【纯充值客户余额刷新】\n"
            for c in pure_depositors:
                orig = orig_bal_dict.get(c, 0.0)
                dep = deposit_dict.get(c, 0.0)
                curr = bal_dict.get(c, 0.0)
                report += f"{c} 今日充值:{fmt(dep)}元 | 最新余额:{fmt(curr)}元\n"
            report += "\n"
            
        total_money = df_sales['金额'].sum()
        unsold_count = len(df_daily_raw) - len(df_sales)
        unsold_str = f" (内含未销售单据 {unsold_count} 车，已全量留底)" if unsold_count > 0 else ""
        
        report += f"1,当日销售共计:{len(df_sales)} 车{fmt(df_sales['净重'].sum())} 吨 {fmt(total_money)} 元,公司配送运费:{fmt(freight_total)}元 ,合计:{fmt(total_money + freight_total)} 元。{unsold_str}\n"
        report += f"2,当日加工费:{fmt(daily_fee)} 元,{report_date.month}月1日-{report_date.day}日加工费合计:{fmt(monthly_fee)} 元。\n"
        report += f"3,当日合计收款:微信零售:{fmt(df_wx['金额'].sum())} 元,共计:{fmt(df_wx['金额'].sum() + df_cash['金额'].sum())} 元\n"

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
            new_df_income.to_excel(writer, sheet_name='财务收入明细', index=False) # 写入第五张表
            
        st.success("✅ 核算完毕！各项预存已计入余额，资金明细已并入总账本！")
        st.download_button(
            label="💾 下载【最新5表合一_地磅总账本.xlsx】",
            data=output.getvalue(),
            file_name=f"地磅总账本_{report_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
            
    except Exception as e:
        st.error(f"处理出错，请确保上传的文件正确。错误信息: {e}")

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="地磅管家 Pro", page_icon="🚛", layout="centered")

st.title("🚛 地磅管家 (精准核算版)")

# ================= 1. 初始化空账本 (新增单号与时间) =================
def create_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=['客户名称', '余额']).to_excel(writer, sheet_name='客户余额', index=False)
        pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价']).to_excel(writer, sheet_name='加工费规则', index=False)
        # 增加：单号、重车时间、备注
        pd.DataFrame(columns=['单号', '重车时间', '车号', '货物名称', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费']).to_excel(writer, sheet_name='过磅明细', index=False)
        # 增加：单号、重车时间
        pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额']).to_excel(writer, sheet_name='公司配送-运费', index=False)
    return output.getvalue()

st.sidebar.header("🛠️ 首次使用向导")
st.sidebar.download_button(
    label="⬇️ 下载【全新4表版_空白总账本.xlsx】",
    data=create_template(),
    file_name="地磅总账本_带运费.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="包含：客户余额、加工费规则、历史明细、公司配送运费 四个工作表"
)

# ----------------- 智能时间解析器 -----------------
def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() == '':
        return ''
    try:
        # 尝试解析 Excel 的数字时间戳 (如 46032.31388)
        f_val = float(val)
        if f_val > 30000:
            return pd.to_datetime(f_val, unit='D', origin='1899-12-30').strftime('%Y-%m-%d %H:%M')
    except:
        pass
    return str(val)

# ================= 2. 文件上传 =================
st.write("### 第一步：上传账本与今日过磅单")
col1, col2 = st.columns(2)
with col1:
    db_file = st.file_uploader("📂 1. 上传【总账本.xlsx】", type=['xlsx'])
with col2:
    daily_file = st.file_uploader("📝 2. 上传【今日单】", type=['xls', 'xlsx', 'csv'])

# ================= 3. 核心计算 =================
if db_file is not None and daily_file is not None:
    try:
        # 读取总账本
        df_bal = pd.read_excel(db_file, sheet_name='客户余额')
        try: df_rules = pd.read_excel(db_file, sheet_name='加工费规则')
        except: df_rules = pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价'])
        try: df_hist = pd.read_excel(db_file, sheet_name='过磅明细')
        except: df_hist = pd.DataFrame()
        try: df_freight = pd.read_excel(db_file, sheet_name='公司配送-运费')
        except: df_freight = pd.DataFrame(columns=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额'])
            
        # 读取今日单
        if daily_file.name.endswith('.csv'):
            df_daily = pd.read_csv(daily_file)
        else:
            df_daily = pd.read_excel(daily_file)
            
        # ---------- 关键清洗：只保留真正的销售记录 ----------
        # 1. 过滤掉过磅类型为空的（如材料过磅、废铁、生产区）
        df_daily = df_daily.dropna(subset=['过磅类型'])
        df_daily = df_daily[df_daily['过磅类型'].astype(str).str.strip() != '']
        
        # 2. 规范化基础字段
        df_daily['单号'] = df_daily.get('单号', '').astype(str).str.replace('.0', '', regex=False)
        df_daily['重车时间'] = df_daily.get('重车时间', '').apply(parse_excel_date)
        df_daily['净重'] = pd.to_numeric(df_daily.get('净重', 0), errors='coerce').fillna(0)
        df_daily['单价'] = pd.to_numeric(df_daily.get('单价', 0), errors='coerce').fillna(0)
        df_daily['金额'] = pd.to_numeric(df_daily.get('金额', 0), errors='coerce').fillna(0)
        df_daily['收货单位'] = df_daily.get('收货单位', '未知客户').fillna('未知客户')
        df_daily['货物名称'] = df_daily.get('货物名称', '未知物料').fillna('未知物料')
        df_daily['备注'] = df_daily.get('备注', '').fillna('')
        
        now = datetime.now()
        
        # ---------------- 智能物流拦截引擎 ----------------
        has_delivery = False
        freight_total = 0.0
        new_freight_records = pd.DataFrame()
        
        if '备注' in df_daily.columns:
            delivery_mask = df_daily['备注'].astype(str).str.contains('公司配送', na=False)
            if delivery_mask.any():
                has_delivery = True
                st.warning("🚚 检测到今日有【公司配送】的车辆！请在下方填入司机和运费单价：")
                
                # 给弹窗加入单号和重车时间，方便核对
                edit_view = df_daily[delivery_mask][['单号', '重车时间', '车号', '收货单位', '货物名称', '净重']].copy()
                edit_view['司机姓名'] = ""  
                edit_view['运费单价'] = 0.0 
                
                edited_df = st.data_editor(
                    edit_view,
                    column_config={
                        "司机姓名": st.column_config.TextColumn("司机姓名", help="填入司机名字", required=True),
                        "运费单价": st.column_config.NumberColumn("运费单价 (元/吨)", min_value=0.0, format="%.2f")
                    },
                    disabled=['单号', '重车时间', '车号', '收货单位', '货物名称', '净重'], 
                    hide_index=True,
                    use_container_width=True
                )
                
                edited_df['运费金额'] = edited_df['净重'] * edited_df['运费单价']
                freight_total = edited_df['运费金额'].sum()
                
                new_freight_records = edited_df.copy()
                new_freight_records = new_freight_records[['单号', '重车时间', '车号', '收货单位', '货物名称', '净重', '司机姓名', '运费单价', '运费金额']]

        if not has_delivery:
            freight_total = st.number_input("🚚 今日报表中暂未检测到配送。若有额外运费(元)，可手动输入:", value=0.0, step=10.0)

        # ---------------- 加工费与分类计算 ----------------
        def calc_fee(row):
            if df_rules.empty: return 0.0
            match = df_rules[(df_rules['物料名称'] == row['货物名称']) & (df_rules['销售单价'] == row['单价'])]
            if not match.empty: return row['净重'] * match.iloc[0]['加工费单价']
            return 0.0
            
        df_daily['加工费'] = df_daily.apply(calc_fee, axis=1)
        daily_fee = df_daily['加工费'].sum()
        
        current_month = now.strftime("%Y-%m")
        monthly_fee = daily_fee
        if not df_hist.empty and '重车时间' in df_hist.columns and '加工费' in df_hist.columns:
            df_hist['加工费'] = pd.to_numeric(df_hist['加工费'], errors='coerce').fillna(0)
            hist_this_month = df_hist[df_hist['重车时间'].astype(str).str.startswith(current_month, na=False)]
            monthly_fee += hist_this_month['加工费'].sum()
        
        bal_dict = dict(zip(df_bal['客户名称'], df_bal['余额']))
        
        # 【关键修复】现在严格依据“过磅类型”列来区分现金、微信和签单！
        df_cash = df_daily[df_daily['过磅类型'].astype(str).str.contains('现金', na=False)]
        df_wx = df_daily[df_daily['过磅类型'].astype(str).str.contains('微信', na=False)]
        df_sign = df_daily[df_daily['过磅类型'].astype(str).str.contains('签单', na=False)]
        
        report = f"{now.strftime('%y年%m月%d日 07:00-18:00')}\n\n"
        
        report += f"现金:{len(df_cash)}车{df_cash['净重'].sum():.2f}吨{df_cash['金额'].sum():.2f}元\n\n"
        
        report += f"微信:{len(df_wx)}车{df_wx['净重'].sum():.2f}吨{df_wx['金额'].sum():.2f}元\n"
        wx_grp = df_wx.groupby('货物名称')[['净重', '金额']].sum()
        wx_cnt = df_wx.groupby('货物名称').size()
        for prod in wx_grp.index:
            report += f"{prod}:{wx_cnt[prod]}车{wx_grp.loc[prod, '净重']:.2f}吨{wx_grp.loc[prod, '金额']:.2f}元\n"
            
        report += f"\n签单:{len(df_sign)}车{df_sign['净重'].sum():.2f}吨\n\n"
        
        sign_custs = df_sign.groupby('收货单位')
        for cust, grp in sign_custs:
            report += f"{cust}:{len(grp)}车{grp['净重'].sum():.2f}吨\n"
            prod_grp = grp.groupby('货物名称')[['净重', '金额']].sum()
            prod_cnt = grp.groupby('货物名称').size()
            for prod in prod_grp.index:
                report += f"{prod}:{prod_cnt[prod]}车{prod_grp.loc[prod, '净重']:.2f}吨{prod_grp.loc[prod, '金额']:.2f} 元\n"
            
            cust_money = grp['金额'].sum()
            prev_bal = bal_dict.get(cust, 0.0)
            curr_bal = prev_bal - cust_money
            bal_dict[cust] = curr_bal
            
            report += f"共金额:{cust_money:.2f} 元\n"
            report += f"上日余额:{prev_bal:.2f} 元\n"
            report += f"当日余额:{curr_bal:.2f} 元\n\n"
            
        total_money = df_daily['金额'].sum()
        report += f"1,当日销售共计:{len(df_daily)} 车{df_daily['净重'].sum():.2f} 吨 {total_money:.2f} 元,公司配送运费:{freight_total:.2f}元 ,合计:{(total_money + freight_total):.2f} 元。\n"
        report += f"2,当日加工费:{daily_fee:.2f} 元,{now.month}月1日-{now.day}日加工费合计:{monthly_fee:.2f} 元。\n"
        report += f"3,当日合计收款:微信零售:{df_wx['金额'].sum():.2f} 元,共计{(df_wx['金额'].sum() + df_cash['金额'].sum()):.2f} 元\n"

        st.write("### 第二步：复制汇报单与司机填报")
        st.code(report, language="text")
        
        # ---------------- 组装新账本 ----------------
        st.write("### 第三步：下载更新后的总账本")
        new_df_bal = pd.DataFrame(list(bal_dict.items()), columns=['客户名称', '余额'])
        
        cols_to_keep = ['单号', '重车时间', '车号', '货物名称', '净重', '单价', '金额', '收货单位', '过磅类型', '备注', '加工费']
        available_cols = [c for c in cols_to_keep if c in df_daily.columns]
        new_df_hist = pd.concat([df_hist, df_daily[available_cols]], ignore_index=True)
        
        new_df_freight = pd.concat([df_freight, new_freight_records], ignore_index=True) if not new_freight_records.empty else df_freight
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            new_df_bal.to_excel(writer, sheet_name='客户余额', index=False)
            df_rules.to_excel(writer, sheet_name='加工费规则', index=False)
            new_df_hist.to_excel(writer, sheet_name='过磅明细', index=False)
            new_df_freight.to_excel(writer, sheet_name='公司配送-运费', index=False)
            
        st.success("✅ 核算完毕！单号与时间已入库，非销售记录已剔除！")
        st.download_button(
            label="💾 下载【最新四表合一_地磅总账本.xlsx】",
            data=output.getvalue(),
            file_name=f"地磅总账本_{now.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
            
    except Exception as e:
        st.error(f"处理出错，请确保上传的文件正确。错误信息: {e}")

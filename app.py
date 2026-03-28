import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="地磅管家 Pro", page_icon="🚛", layout="centered")

st.title("🚛 地磅管家 Pro (Streamlit版)")

# 侧边栏导航
st.sidebar.header("功能菜单")
page = st.sidebar.radio("请选择操作", ["📥 导入与每日汇报", "📋 历史数据预览"])

if page == "📥 导入与每日汇报":
    st.write("### 上传今日过磅记录 (Excel/CSV)")
    uploaded_file = st.file_uploader("点击选择或拖拽文件到此处", type=['xls', 'xlsx', 'csv'])

    if uploaded_file is not None:
        try:
            # 读取表格
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # 过滤作废记录 (容错处理)
            if '状态' in df.columns:
                df = df[~df['状态'].astype(str).str.contains('作废|生产', na=False)]
            
            st.success(f"✅ 成功读取数据，共 {len(df)} 条有效记录！")
            
            # --- 核心核算逻辑 ---
            # 确保列名存在，填充空值
            df['净重'] = pd.to_numeric(df.get('净重', 0), errors='coerce').fillna(0)
            df['金额'] = pd.to_numeric(df.get('金额', 0), errors='coerce').fillna(0)
            df['收货单位'] = df.get('收货单位', '未知客户').fillna('未知客户')
            df['货物名称'] = df.get('货物名称', '未知物料').fillna('未知物料')
            
            # 分类统计
            df_cash = df[df['收货单位'].str.contains('现金', na=False)]
            df_wx = df[df['收货单位'].str.contains('微信', na=False)]
            df_sign = df[~df['收货单位'].str.contains('现金|微信', na=False)]
            
            # 构建汇报文本
            today_str = datetime.now().strftime("%y年%m月%d日 07:00-18:00")
            report = f"{today_str}\n\n"
            
            # 现金
            if len(df_cash) > 0:
                report += f"现金:{len(df_cash)}车{df_cash['净重'].sum():.2f}吨{df_cash['金额'].sum():.2f}元\n\n"
            else:
                report += "现金:无\n\n"
            
            # 微信
            if len(df_wx) > 0:
                report += f"微信:{len(df_wx)}车{df_wx['净重'].sum():.2f}吨{df_wx['金额'].sum():.2f}元\n"
                wx_grp = df_wx.groupby('货物名称')[['净重', '金额']].sum()
                wx_cnt = df_wx.groupby('货物名称').size()
                for prod in wx_grp.index:
                    report += f"{prod}:{wx_cnt[prod]}车{wx_grp.loc[prod, '净重']:.2f}吨{wx_grp.loc[prod, '金额']:.2f}元\n"
            else:
                report += "微信:无\n"
            
            # 签单
            report += f"\n签单:{len(df_sign)}车{df_sign['净重'].sum():.2f}吨\n\n"
            sign_custs = df_sign.groupby('收货单位')
            for cust, grp in sign_custs:
                report += f"{cust}:{len(grp)}车{grp['净重'].sum():.2f}吨\n"
                prod_grp = grp.groupby('货物名称')[['净重', '金额']].sum()
                prod_cnt = grp.groupby('货物名称').size()
                for prod in prod_grp.index:
                    report += f"{prod}:{prod_cnt[prod]}车{prod_grp.loc[prod, '净重']:.2f}吨{prod_grp.loc[prod, '金额']:.2f} 元\n"
                report += f"共金额:{grp['金额'].sum():.2f} 元\n"
                report += f"当日余额: -- 元 (暂未连接数据库)\n\n"
            
            # 底部汇总
            total_money = df['金额'].sum()
            total_cars = len(df)
            total_weight = df['净重'].sum()
            
            report += f"1,当日销售共计:{total_cars} 车{total_weight:.2f} 吨 {total_money:.2f} 元,公司配送运费:0元 ,合计:{total_money:.2f} 元。\n"
            report += f"3,当日合计收款:微信零售:{df_wx['金额'].sum():.2f} 元,共计{(df_wx['金额'].sum() + df_cash['金额'].sum()):.2f} 元\n"

            st.write("### 📝 今日自动汇报单")
            st.code(report, language="text")
            
        except Exception as e:
            st.error(f"读取文件时出错，请确保表格格式正确。错误信息: {e}")

elif page == "📋 历史数据预览":
    st.info("💡 提示：纯云端部署无永久本地存储。如需永久保存客户余额与历史账单，下一步我们将教您连接腾讯文档/飞书多维表格。")

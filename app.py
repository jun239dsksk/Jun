import streamlit as st
import pandas as pd
from datetime import datetime
import io
import traceback

# =====================================================
# 0. 日志引擎
# =====================================================

if 'app_logs' not in st.session_state:
    st.session_state.app_logs = []

def add_log(level, msg):
    t = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    st.session_state.app_logs.append('[' + t + '] [' + level + '] ' + msg)

@st.dialog('📝 系统运行日志')
def show_logs_dialog():
    logs = st.session_state.app_logs
    if not logs:
        st.info('暂无日志记录')
        return
    info_c = sum(1 for l in logs if '[INFO]' in l)
    warn_c = sum(1 for l in logs if '[WARN]' in l)
    err_c  = sum(1 for l in logs if '[ERROR]' in l)
    m1, m2, m3 = st.columns(3)
    m1.metric('✅ INFO', info_c)
    m2.metric('⚠️ WARN', warn_c)
    m3.metric('❌ ERROR', err_c)
    st.markdown('---')
    filter_level = st.radio('筛选级别', ['全部', 'INFO', 'WARN', 'ERROR'], horizontal=True)
    filtered = logs if filter_level == '全部' else [l for l in logs if '[' + filter_level + ']' in l]
    if not filtered:
        st.info('暂无 ' + filter_level + ' 级别日志')
    else:
        lines_html = []
        for line in filtered:
            if '[ERROR]' in line:
                color = '#ff4b4b'
            elif '[WARN]' in line:
                color = '#ffa500'
            else:
                color = '#21c55e'
            lines_html.append(
                '<div style="font-family:monospace;font-size:12px;padding:2px 0;'
                'border-left:3px solid ' + color + ';padding-left:8px;margin-bottom:2px;'
                'background:#f8f9fa;border-radius:0 4px 4px 0;">' + line + '</div>'
            )
        st.markdown('\n'.join(lines_html), unsafe_allow_html=True)
    st.markdown('---')
    log_text = '\n'.join(logs)
    st.download_button(
        '📥 导出日志 (.txt)',
        data=log_text,
        file_name='地磅日志_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.txt',
        use_container_width=True,
    )

# =====================================================
# 1. 页面配置 & CSS
# =====================================================

st.set_page_config(
    page_title='地磅管家 Pro',
    page_icon='🚛',
    layout='wide',
    initial_sidebar_state='auto',
)

st.markdown('''
<style>
.block-container {
    padding-top: 2.5rem !important;
    padding-bottom: 1rem !important;
    max-width: 100% !important;
}
.stSelectbox > div > div > div,
.stNumberInput input,
.stButton > button,
.stPopover > button {
    height: 36px !important;
    min-height: 36px !important;
    font-size: 13px !important;
}
.stTextInput > div > div,
.stNumberInput > div > div,
.stSelectbox > div > div {
    background-color: transparent !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 6px !important;
}
div[data-testid="stPopoverBody"] { max-width: 92vw !important; }
button:has(p:contains("✅")) {
    background-color: #f0fdf4 !important;
    border-color: #22c55e !important;
    color: #166534 !important;
}
hr { border-color: #f1f5f9 !important; margin: 0.6rem 0 !important; }
div[data-testid="stAlert"] { padding: 0.5rem 0.8rem !important; }
details > summary { font-size: 13px !important; font-weight: 600 !important; }
.stCode { font-size: 13px !important; }
@media (max-width: 768px) {
    [data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 0.3rem !important;
    }
    [data-testid="column"] {
        min-width: 0 !important;
        flex-shrink: 1 !important;
        margin-bottom: 0.3rem !important;
    }
    .block-container { padding-left: 0.5rem !important; padding-right: 0.5rem !important; }
}
</style>
''', unsafe_allow_html=True)

# =====================================================
# 2. 侧边栏
# =====================================================

with st.sidebar:
    st.markdown('### 📅 报表时间')
    report_date = st.date_input('报表日期', value=datetime.now(), label_visibility='collapsed')
    report_time = st.text_input('时间段', value='07:00-18:00', label_visibility='collapsed')
    st.divider()
    st.markdown('### ⚙️ 取整设置')
    st.caption('勾选 = 取整，否则保留2位小数')
    round_retail  = st.checkbox('零售(微信/现金)', value=True)
    round_sign    = st.checkbox('签单客户',         value=False)
    round_fee     = st.checkbox('加工费',           value=False)
    round_freight = st.checkbox('运费',             value=False)
    st.divider()
    if st.button('📝 查看 / 导出系统日志', use_container_width=True):
        show_logs_dialog()
    st.divider()
    st.markdown('### 📋 空白模板')

# =====================================================
# 3. 工具函数
# =====================================================

def do_round(val, category=''):
    if pd.isna(val):
        return 0.0
    val = float(val)
    flags = {'retail': round_retail, 'sign': round_sign, 'fee': round_fee, 'freight': round_freight}
    if flags.get(category, False):
        return float(int(val + 0.5) if val >= 0 else int(val - 0.5))
    return round(val + 1e-9, 2)

def fmt_val(v, category=''):
    if pd.isna(v):
        return '0'
    v = float(v)
    flags = {'retail': round_retail, 'sign': round_sign, 'fee': round_fee, 'freight': round_freight}
    if flags.get(category, False):
        return str(int(v + 0.5) if v >= 0 else int(v - 0.5))
    res = f'{v:.2f}'.rstrip('0').rstrip('.')
    return res

def fmt_weight(v):
    if pd.isna(v):
        return '0'
    res = f'{float(v):.2f}'.rstrip('0').rstrip('.')
    return res

def safe_concat(dfs):
    valid = [df for df in dfs if not df.empty]
    if not valid:
        return dfs[0] if dfs else pd.DataFrame()
    return pd.concat(valid, ignore_index=True)

def parse_excel_date(val):
    if pd.isna(val) or str(val).strip() == '':
        return ''
    try:
        f_val = float(val)
        if f_val > 30000:
            return pd.to_datetime(f_val, unit='D', origin='1899-12-30').strftime('%Y-%m-%d %H:%M')
    except Exception:
        pass
    return str(val)

def create_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=['客户名称', '余额']).to_excel(writer, sheet_name='客户余额', index=False)
        pd.DataFrame(columns=['物料名称', '销售单价', '加工费单价']).to_excel(writer, sheet_name='加工费规则', index=False)
        pd.DataFrame(columns=['原始名称', '标准名称']).to_excel(writer, sheet_name='物料归类映射', index=False)
        pd.DataFrame(columns=['原始名称', '标准名称']).to_excel(writer, sheet_name='客户归类映射', index=False)
        pd.DataFrame(columns=['单号','重车时间','车号','货物名称','标准归类','净重','单价','金额',
                              '收货单位','过磅类型','备注','加工费单价','加工费','备注2']
                    ).to_excel(writer, sheet_name='过磅明细', index=False)
        pd.DataFrame(columns=['单号','重车时间','车号','收货单位','货物名称','净重','司机姓名','运费单价','运费金额']
                    ).to_excel(writer, sheet_name='公司配送-运费', index=False)
        pd.DataFrame(columns=['日期','客户名称','收入类型','金额','备注']
                    ).to_excel(writer, sheet_name='财务收入明细', index=False)
    return output.getvalue()

with st.sidebar:
    st.download_button(
        label='⬇ 下载空白模板(V3)',
        data=create_template(),
        file_name='地磅总账本_空白_V3.xlsx',
        use_container_width=True,
    )

# =====================================================
# 4. 主界面头部
# =====================================================

st.markdown('## 🚛 地磅管家 Pro')
st.divider()
st.markdown('##### 📂 第一步：上传业务文件')
c1, c2 = st.columns(2)
with c1:
    db_file    = st.file_uploader('总账本 (.xlsx)', type=['xlsx'])
with c2:
    daily_file = st.file_uploader('今日过磅单 (.xls/.xlsx/.csv)', type=['xls','xlsx','csv'])

# =====================================================
# 5. 弹窗：新增司机
# =====================================================

@st.dialog('➕ 添加新司机')
def add_driver_modal(truck_no):
    new_d = st.text_input('【' + truck_no + '】的新司机姓名：', placeholder='输入姓名')
    if st.button('确认添加', type='primary', use_container_width=True):
        if new_d.strip():
            st.session_state['custom_drv_' + truck_no] = new_d.strip()
            add_log('INFO', '新增司机: ' + new_d.strip() + ' (车号:' + truck_no + ')')
        st.rerun()

# =====================================================
# 6. 主处理流程
# =====================================================

if db_file is not None and daily_file is not None:
    try:
        load_key = 'loaded_' + db_file.name + '_' + daily_file.name
        if load_key not in st.session_state:
            add_log('INFO', '读取总账本: ' + db_file.name)
            add_log('INFO', '读取今日过磅单: ' + daily_file.name)
            st.session_state[load_key] = True

        xls = pd.ExcelFile(db_file)
        sheets = xls.sheet_names

        def read_sheet(name, cols):
            if name in sheets:
                df = pd.read_excel(xls, sheet_name=name)
                missing = [c for c in cols if c not in df.columns]
                if missing:
                    add_log('WARN', '工作表 [' + name + '] 缺少列: ' + str(missing) + '，已用空表替代')
                    return pd.DataFrame(columns=cols)
                return df
            add_log('WARN', '总账本缺少工作表: [' + name + ']，将自动创建')
            return pd.DataFrame(columns=cols)

        df_bal          = read_sheet('客户余额',      ['客户名称', '余额'])
        df_rules        = read_sheet('加工费规则',    ['物料名称', '销售单价', '加工费单价'])
        df_mapping      = read_sheet('物料归类映射',  ['原始名称', '标准名称'])
        df_cust_mapping = read_sheet('客户归类映射',  ['原始名称', '标准名称'])
        df_hist         = read_sheet('过磅明细',      ['单号'])
        df_freight      = read_sheet('公司配送-运费', ['单号','重车时间','车号','收货单位','货物名称','净重','司机姓名','运费单价','运费金额'])
        df_income       = read_sheet('财务收入明细',  ['日期','客户名称','收入类型','金额','备注'])

        mapping_dict = dict(zip(
            df_mapping['原始名称'].astype(str).str.strip(),
            df_mapping['标准名称'].astype(str).str.strip(),
        ))
        
        cust_mapping_dict = dict(zip(
            df_cust_mapping['原始名称'].astype(str).str.strip(),
            df_cust_mapping['标准名称'].astype(str).str.strip(),
        ))

        if daily_file.name.endswith('.csv'):
            df_daily_raw = pd.read_csv(daily_file)
            add_log('INFO', '今日过磅单格式: CSV')
        else:
            df_daily_raw = pd.read_excel(daily_file)
            add_log('INFO', '今日过磅单格式: Excel')

        raw_count = len(df_daily_raw)
        if '状态' in df_daily_raw.columns:
            df_daily_raw = df_daily_raw[~df_daily_raw['状态'].astype(str).str.contains('作废|生产', na=False)]
            filtered_count = raw_count - len(df_daily_raw)
            if filtered_count:
                add_log('INFO', '过滤作废/生产记录 ' + str(filtered_count) + ' 条，剩余 ' + str(len(df_daily_raw)) + ' 条')

        df_daily_raw['单号']       = df_daily_raw.get('单号', '').astype(str).str.replace('.0', '', regex=False)
        df_daily_raw['重车时间']    = df_daily_raw.get('重车时间', '').apply(parse_excel_date)
        df_daily_raw['货物名称']    = df_daily_raw.get('货物名称', '未知物料').astype(str).str.strip()
        df_daily_raw['净重']        = pd.to_numeric(df_daily_raw.get('净重', 0), errors='coerce').fillna(0)
        df_daily_raw['单价']        = pd.to_numeric(df_daily_raw.get('单价', 0), errors='coerce').fillna(0)
        df_daily_raw['金额']        = pd.to_numeric(df_daily_raw.get('金额', 0), errors='coerce').fillna(0)
        df_daily_raw['过磅类型']    = df_daily_raw.get('过磅类型', '').astype(str)
        df_daily_raw['备注']        = df_daily_raw.get('备注', '').fillna('')
        df_daily_raw['备注2']       = ''
        df_daily_raw['_orig_idx']   = df_daily_raw.index
        df_daily_raw['汇报专用名称'] = ''
        add_log('INFO', '字段清洗完成，共 ' + str(len(df_daily_raw)) + ' 条有效记录')

        # ============================================================
        # 第一关：混合料拆分
        # ============================================================
        mixed_mask = df_daily_raw['货物名称'].str.contains(r'\+', na=False)
        if mixed_mask.any():
            mixed_mats = df_daily_raw.loc[mixed_mask, '货物名称'].unique()
            add_log('WARN', '检测到混合料 ' + str(len(mixed_mats)) + ' 种: ' + str(list(mixed_mats)))
            st.warning('⚠️ **第一关：检测到 ' + str(len(mixed_mats)) + ' 种混合料，请配置拆分比例**')

            split_ratios = {}
            for mat in mixed_mats:
                parts = mat.split('+')
                p1 = parts[0].strip()
                p2 = parts[1].strip() if len(parts) > 1 else '其他料'
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    pct1 = st.number_input('【' + mat + '】[' + p1 + '] 占比(%)',
                        min_value=1.0, max_value=99.0, value=50.0, step=1.0, key='s1_' + mat)
                with col_s2:
                    st.number_input('[' + p2 + '] 占比(%)', value=100.0 - pct1, disabled=True, key='s2_' + mat)
                split_ratios[mat] = pct1

            split_confirmed = st.checkbox('✅ 比例已确认，立刻拆分并进入下一步', key='chk_split')

            new_rows = []
            for _, row in df_daily_raw.iterrows():
                mat = str(row['货物名称'])
                if mat not in split_ratios:
                    row['_pre_rounded'] = False
                    new_rows.append(row)
                    continue
                pct1 = split_ratios[mat]
                pct2 = 100.0 - pct1
                parts = mat.split('+')
                mat1 = parts[0].strip()
                mat2 = parts[1].strip() if len(parts) > 1 else '其他'
                r1, r2 = row.copy(), row.copy()
                w_orig, p_orig, a_orig = row['净重'], row['单价'], row['金额']
                t_type, dh = str(row['过磅类型']), str(row['单号'])
                cat = 'retail' if ('微信' in t_type or '现金' in t_type) else ('sign' if '签单' in t_type else 'none')
                if p_orig == 0 and a_orig == 0:
                    w1 = round(w_orig * (pct1 / 100.0), 3)
                    w2 = w_orig - w1
                    r1['净重'], r1['单价'], r1['金额'] = w1, 0.0, 0.0
                    r2['净重'], r2['单价'], r2['金额'] = w2, 0.0, 0.0
                    r1['_pre_rounded'] = r2['_pre_rounded'] = False
                else:
                    true_total = (w_orig * p_orig) if p_orig > 0 else a_orig
                    rounded_total = do_round(true_total, cat)
                    w1 = round(w_orig * (pct1 / 100.0), 3)
                    w2 = w_orig - w1
                    exact_a1 = w1 * p_orig if p_orig > 0 else true_total * (pct1 / 100.0)
                    a1 = do_round(exact_a1, cat)
                    a2 = round(rounded_total - a1, 2)
                    r1['净重'], r1['单价'], r1['金额'] = w1, p_orig, a1
                    r2['净重'], r2['单价'], r2['金额'] = w2, p_orig, a2
                    r1['_pre_rounded'] = r2['_pre_rounded'] = True
                r1['货物名称'], r2['货物名称'] = mat1, mat2
                r1['汇报专用名称'] = r2['汇报专用名称'] = mat
                r1['备注2'] = r2['备注2'] = mat1 + ' ' + str(pct1) + '%+' + mat2 + ' ' + str(pct2) + '% 已拆分单号 ' + dh
                new_rows.extend([r1, r2])

            df_daily_raw = pd.DataFrame(new_rows)
            st.divider()
            if not split_confirmed:
                add_log('WARN', '流程挂起：等待确认混合料拆分比例')
                st.stop()
            else:
                split_log_key = 'log_split_' + report_date.strftime('%Y%m%d%H%M%S')
                if split_log_key not in st.session_state:
                    add_log('INFO', '混合料拆分完成: ' + str(list(mixed_mats)))
                    st.session_state[split_log_key] = True

        # ============================================================
        # 第二关：物料归类映射
        # ============================================================
        df_daily_raw['标准归类']     = df_daily_raw['货物名称'].apply(lambda x: mapping_dict.get(x, x))
        df_daily_raw['汇报专用名称'] = df_daily_raw.apply(
            lambda r: r['标准归类'] if r['汇报专用名称'] == '' else r['汇报专用名称'], axis=1)

        known_originals = set(df_mapping['原始名称'].dropna().astype(str).str.strip().unique())
        known_standards = set(df_rules['物料名称'].dropna().astype(str).str.strip().unique())
        unknown_mats = [
            mat for mat in df_daily_raw['货物名称'].astype(str).str.strip().unique()
            if mat and mat not in known_originals and mat not in known_standards
        ]

        if unknown_mats:
            add_log('WARN', '发现未知物料 ' + str(len(unknown_mats)) + ' 种: ' + str(unknown_mats))
            st.error('🛑 **第二关：发现 ' + str(len(unknown_mats)) + ' 种未知物料，请指定父类**')
            map_options = ['(独立物料)', '(非销售/不计价)'] + sorted(known_standards)
            cols = st.columns(3)
            mapping_inputs = {}
            for i, u_mat in enumerate(unknown_mats):
                with cols[i % 3]:
                    mapping_inputs[u_mat] = st.selectbox(
                        '【' + u_mat + '】', ['(请选择...)'] + map_options, key='map_' + u_mat)
            map_confirmed = st.checkbox('✅ 物料归类已确认', key='chk_map')
            st.divider()
            if not map_confirmed:
                add_log('WARN', '流程挂起：等待用户确认未知物料归类')
                st.stop()
            new_mapping_records = []
            for u_mat, sel_val in mapping_inputs.items():
                if sel_val != '(请选择...)':
                    std_name = u_mat if sel_val == '(独立物料)' else sel_val
                    mapping_dict[u_mat] = std_name
                    new_mapping_records.append({'原始名称': u_mat, '标准名称': std_name})
                    add_log('INFO', '物料映射: [' + u_mat + '] → [' + std_name + ']')
            if new_mapping_records:
                df_mapping = safe_concat([df_mapping, pd.DataFrame(new_mapping_records)])
                df_daily_raw['标准归类'] = df_daily_raw['货物名称'].apply(lambda x: mapping_dict.get(x, x))
                df_daily_raw['汇报专用名称'] = df_daily_raw.apply(
                    lambda r: r['标准归类'] if r['汇报专用名称'] == r['货物名称'] else r['汇报专用名称'], axis=1)
            add_log('INFO', '物料归类映射完成，新增 ' + str(len(new_mapping_records)) + ' 条映射记录')

        # ============================================================
        # 第三关：客户归类映射
        # ============================================================
        if '收货单位' not in df_daily_raw.columns:
            df_daily_raw['收货单位'] = ''

        def fix_shdw(row):
            shdw    = str(row.get('收货单位', '')).strip()
            gb_type = str(row.get('过磅类型', '')).strip()
            if shdw in ('', 'nan'):
                if gb_type == '':
                    return '内部单'
                if '微信' in gb_type or '现金' in gb_type:
                    return '零售客户'
                return '未知客户'
            return shdw

        df_daily_raw['收货单位'] = df_daily_raw.apply(fix_shdw, axis=1)
        df_daily_raw['收货单位'] = df_daily_raw['收货单位'].apply(lambda x: cust_mapping_dict.get(x, x))

        known_custs_base = set(df_bal['客户名称'].dropna().astype(str).str.strip())
        if '收货单位' in df_hist.columns:
            known_custs_base.update(df_hist['收货单位'].dropna().astype(str).str.strip())
        known_custs_base.update(['内部单', '零售客户', '未知客户'])
        known_custs_base.update(df_cust_mapping['标准名称'].dropna().astype(str).str.strip())

        unknown_custs = [c for c in df_daily_raw['收货单位'].astype(str).str.strip().unique()
                         if c and c not in known_custs_base]

        if unknown_custs:
            add_log('WARN', '发现未知客户: ' + str(unknown_custs))
            st.error('🛑 **第三关：发现 ' + str(len(unknown_custs)) + ' 家未知客户，请指定映射**')
            st.markdown('💡 如果它是老客户的别名（如项目部），请选择对应老客户；如果是新客户，请选 `(新建客户)`。建档后可直接在最下方财务登记处使用它！')

            cust_map_options = ['(新建客户)'] + sorted([c for c in known_custs_base if c not in ('内部单', '零售客户', '未知客户')])
            cols = st.columns(3)
            cust_mapping_inputs = {}
            for i, u_cust in enumerate(unknown_custs):
                with cols[i % 3]:
                    cust_mapping_inputs[u_cust] = st.selectbox(
                        '【' + u_cust + '】', ['(请选择...)'] + cust_map_options, key='cmap_' + u_cust)

            cust_map_confirmed = st.checkbox('✅ 客户归类已确认', key='chk_cmap')
            st.divider()
            if not cust_map_confirmed:
                add_log('WARN', '流程挂起：等待用户确认客户归类')
                st.stop()

            new_cust_mapping_records = []
            for u_cust, sel_val in cust_mapping_inputs.items():
                if sel_val != '(请选择...)':
                    std_cust = u_cust if sel_val == '(新建客户)' else sel_val
                    cust_mapping_dict[u_cust] = std_cust
                    new_cust_mapping_records.append({'原始名称': u_cust, '标准名称': std_cust})
                    add_log('INFO', '客户映射: [' + u_cust + '] → [' + std_cust + ']')

            if new_cust_mapping_records:
                df_cust_mapping = safe_concat([df_cust_mapping, pd.DataFrame(new_cust_mapping_records)])
                df_daily_raw['收货单位'] = df_daily_raw['收货单位'].apply(lambda x: cust_mapping_dict.get(x, x))
            add_log('INFO', '客户映射完成')

        # ============================================================
        # 第四关：补填缺失单价
        # ============================================================
        missing_mask = (
            (df_daily_raw['单价'] == 0)
            & (df_daily_raw['金额'] == 0)
            & (df_daily_raw['过磅类型'].str.strip() != '')
            & (df_daily_raw['标准归类'] != '(非销售/不计价)')
        )
        if missing_mask.any():
            missing_groups = df_daily_raw[missing_mask].groupby(['收货单位', '货物名称'])
            add_log('WARN', '检测到 ' + str(missing_mask.sum()) + ' 条记录缺失单价，涉及 ' + str(len(missing_groups)) + ' 个组合')
            st.warning('⚠️ **第四关：' + str(missing_mask.sum()) + ' 条有效销售记录缺失单价，请补充**')
            cols = st.columns(4)
            price_inputs = {}
            for i, ((cust, mat), _) in enumerate(missing_groups):
                with cols[i % 4]:
                    price_inputs[(cust, mat)] = st.number_input(
                        '**' + cust + '**\n' + mat + ' (元/吨)',
                        min_value=0.0, step=1.0, format='%.2f', key='miss_p_' + cust + '_' + mat)
            price_confirmed = st.checkbox('✅ 单价已全部补齐', key='chk_price')
            st.divider()
            if not price_confirmed:
                add_log('WARN', '流程挂起：等待用户补充缺失单价')
                st.stop()
            for idx, row in df_daily_raw[missing_mask].iterrows():
                p_val = price_inputs.get((row['收货单位'], row['货物名称']), 0.0)
                df_daily_raw.at[idx, '单价'] = p_val
            add_log('INFO', '缺失单价已全部补齐')

        # ============================================================
        # 第五关：补填缺失加工费规则
        # ============================================================
        valid_sales  = df_daily_raw[
            (df_daily_raw['过磅类型'].astype(str).str.strip() != '')
            & (df_daily_raw['标准归类'] != '(非销售/不计价)')
        ]
        unique_combos = valid_sales[['标准归类','单价']].drop_duplicates()
        missing_rules = []
        for _, r in unique_combos.iterrows():
            mat   = str(r['标准归类']).strip()
            price = float(r['单价'])
            if df_rules.empty:
                missing_rules.append((mat, price))
                continue
            match = df_rules[
                (df_rules['物料名称'].astype(str).str.strip() == mat)
                & (df_rules['销售单价'].astype(float) == price)
            ]
            if match.empty or pd.isna(match.iloc[0]['加工费单价']) or str(match.iloc[0]['加工费单价']).strip() == '':
                missing_rules.append((mat, price))

        if missing_rules:
            add_log('WARN', '发现 ' + str(len(missing_rules)) + ' 条未收录加工费规则: ' + str(missing_rules))
            st.error('🛑 **第五关：发现 ' + str(len(missing_rules)) + ' 条未收录加工费规则，请补齐**')
            cols = st.columns(4)
            fee_inputs = {}
            for i, (mat, price) in enumerate(missing_rules):
                with cols[i % 4]:
                    fee_inputs[(mat, price)] = st.number_input(
                        '**' + mat + '**\n单价:' + str(price) + '元 → 加工费',
                        min_value=0.0, step=1.0, format='%.2f', key='miss_f_' + mat + '_' + str(price))
            fee_confirmed = st.checkbox('✅ 加工费已补齐，完成最终核算', key='chk_fee')
            st.divider()
            if not fee_confirmed:
                add_log('WARN', '流程挂起：等待用户补充加工费规则')
                st.stop()
            new_rules = [{'物料名称': mat, '销售单价': price, '加工费单价': fee}
                         for (mat, price), fee in fee_inputs.items()]
            df_rules = safe_concat([df_rules, pd.DataFrame(new_rules)])
            for (mat, price), fee in fee_inputs.items():
                add_log('INFO', '新增加工费规则: [' + mat + '] 单价:' + str(price) + ' → 加工费:' + str(fee))

        # ============================================================
        # 最终精确核算
        # ============================================================
        add_log('INFO', '开始最终精确核算...')
        new_amts = []
        for _, row in df_daily_raw.iterrows():
            if row['标准归类'] == '(非销售/不计价)':
                new_amts.append(0.0)
            elif row.get('_pre_rounded', False):
                new_amts.append(row['金额'])
            else:
                w, p, orig_a, t_type = row['净重'], row['单价'], row['金额'], str(row['过磅类型'])
                exact = (w * p) if p > 0 else orig_a
                cat = 'retail' if ('微信' in t_type or '现金' in t_type) else ('sign' if '签单' in t_type else 'none')
                new_amts.append(do_round(exact, cat))
        df_daily_raw['金额'] = new_amts

        def calc_fee_price(row):
            if str(row.get('过磅类型','')).strip() == '' or row.get('标准归类','') == '(非销售/不计价)':
                return 0.0
            match = df_rules[
                (df_rules['物料名称'].astype(str).str.strip() == str(row['标准归类']).strip())
                & (df_rules['销售单价'].astype(float) == float(row['单价']))
            ]
            return float(match.iloc[0]['加工费单价']) if not match.empty else 0.0

        df_daily_raw['加工费单价'] = df_daily_raw.apply(calc_fee_price, axis=1)
        df_daily_raw['加工费']     = df_daily_raw.apply(
            lambda r: do_round(r['净重'] * r['加工费单价'], 'fee') if r['加工费单价'] > 0 else 0.0, axis=1)
        add_log('INFO', '核算完成：总销售额 ' + str(round(df_daily_raw['金额'].sum(),2)) + '，总加工费 ' + str(round(df_daily_raw['加工费'].sum(),2)))

        df_report_base = df_daily_raw.groupby('_orig_idx').agg({
            '单号': 'first', '重车时间': 'first', '车号': 'first',
            '收货单位': 'first', '汇报专用名称': 'first', '标准归类': 'first',
            '过磅类型': 'first', '备注': 'first',
            '净重': 'sum', '金额': 'sum',
        }).reset_index(drop=True)

        df_sales_report = df_report_base[
            (df_report_base['过磅类型'].astype(str).str.strip() != '')
            & (df_report_base['标准归类'] != '(非销售/不计价)')
        ].copy()

        all_known_drivers = sorted([
            str(d) for d in df_freight['司机姓名'].dropna().unique()
            if str(d).strip() and str(d) != 'nan'
        ])
        driver_options = ['(未选择)'] + all_known_drivers + ['➕ 手动输入新司机...']

        hist_custs = df_hist['收货单位'].dropna().astype(str).tolist() if '收货单位' in df_hist.columns else []
        daily_custs = df_sales_report['收货单位'].dropna().astype(str).tolist()
        map_custs = df_cust_mapping['标准名称'].dropna().astype(str).tolist()
        all_known_custs = sorted(set(df_bal['客户名称'].dropna().astype(str).tolist() + hist_custs + daily_custs + map_custs))
        all_known_custs = [c for c in all_known_custs if c.strip() and c not in ('nan','内部单','零售客户','未知客户')]
        cust_options = ['(不录入)'] + all_known_custs + ['➕ 手动输入新客户...']

        # ============================================================
        # 公司配送模块
        # ============================================================
        freight_total  = 0.0
        new_freight_df = pd.DataFrame()
        delivery_mask  = df_sales_report['备注'].astype(str).str.contains('公司配送', na=False)
        has_delivery   = delivery_mask.any()
        delivery_count = delivery_mask.sum()
        expander_title = ('🔴 🚚 检测到 ' + str(delivery_count) + ' 车公司配送，点击展开分配'
                          if has_delivery else '🚚 公司配送与额外运费')

        with st.expander(expander_title, expanded=False):
            if has_delivery:
                add_log('INFO', '检测到 ' + str(delivery_count) + ' 车公司配送记录')
                delivery_df   = df_sales_report[delivery_mask].copy()
                truck_counts  = delivery_df['车号'].value_counts()
                unique_trucks = truck_counts.index.tolist()
                unique_delivery_custs = delivery_df['收货单位'].dropna().unique()

                mem_driver, mem_price = {}, {}
                if not df_freight.empty:
                    for _, r in df_freight.iterrows():
                        if pd.notna(r.get('车号')) and pd.notna(r.get('司机姓名')):
                            mem_driver[str(r['车号'])] = str(r['司机姓名'])
                        if pd.notna(r.get('收货单位')) and pd.notna(r.get('运费单价')):
                            mem_price[str(r['收货单位'])] = float(r['运费单价'])

                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    if st.button('🔄 全选/反选', use_container_width=True):
                        curr = st.session_state.get('batch_sel', False)
                        st.session_state['batch_sel'] = not curr
                        for t in unique_trucks:
                            st.session_state['chk_' + str(t)] = not curr
                        add_log('INFO', '批量' + ('全选' if not curr else '反选') + '配送车辆')
                        st.rerun()
                with col_b2:
                    with st.popover('⚙️ 批量设置运价/司机', use_container_width=True):
                        st.markdown('**批量分配司机**')
                        b_drv = st.selectbox('统一分配司机', driver_options[:-1])
                        if st.button('应用司机(仅勾选)', use_container_width=True):
                            for t in unique_trucks:
                                if st.session_state.get('chk_' + str(t), False) and b_drv != '(未选择)':
                                    st.session_state['custom_drv_' + str(t)] = b_drv
                            add_log('INFO', '批量分配司机: ' + b_drv)
                            st.rerun()
                        st.divider()
                        st.markdown('**按客户设置运价**')
                        b_cust = st.selectbox('目标', ['(对所有已勾选)'] + list(unique_delivery_custs))
                        b_prc  = st.number_input('运价', min_value=0.0, step=1.0, format='%.2f')
                        if st.button('应用运价', use_container_width=True):
                            for t in unique_trucks:
                                t_cust = delivery_df[delivery_df['车号'] == t].iloc[0]['收货单位']
                                if b_cust == '(对所有已勾选)':
                                    if st.session_state.get('chk_' + str(t), False):
                                        st.session_state['p_' + str(t)] = b_prc
                                elif t_cust == b_cust:
                                    st.session_state['p_' + str(t)] = b_prc
                            add_log('INFO', '批量设置运价: ' + str(b_prc) + '元，目标: ' + b_cust)
                            st.rerun()

                driver_map, price_map = {}, {}

                def render_truck_row(t):
                    count    = truck_counts[t]
                    this_df  = delivery_df[delivery_df['车号'] == t]
                    curr_drv = st.session_state.get('custom_drv_' + str(t), mem_driver.get(str(t), ''))
                    curr_prc = st.session_state.get('p_' + str(t), mem_price.get(str(this_df.iloc[0]['收货单位']), 0.0))
                    prefix   = '✅ ' if (curr_drv and curr_prc > 0) else '🚛 '
                    c_chk, c_info, c_drv, c_prc = st.columns([2.5, 2, 2.5, 3])
                    with c_chk:
                        st.checkbox(prefix + str(t) + ' (' + str(count) + '趟)', key='chk_' + str(t))
                    with c_info:
                        with st.popover('📋 明细', use_container_width=True):
                            for _, r in this_df.iterrows():
                                st.markdown(
                                    '📄 **' + str(r['单号']) + '** | 🏢 ' + str(r['收货单位']) + '\n\n' +
                                    '📦 ' + str(r['汇报专用名称']) + ' `' + str(r['净重']) + '吨` | 🕒 ' + str(r['重车时间'])
                                )
                                st.markdown('---')
                    with c_drv:
                        opts = driver_options.copy()
                        if curr_drv and curr_drv not in opts:
                            opts.insert(1, curr_drv)
                        idx_sel = opts.index(curr_drv) if curr_drv in opts else 0
                        d_sel = st.selectbox(
                            '司机_' + str(t), opts, index=idx_sel,
                            key='d_sel_' + str(t), label_visibility='collapsed')
                        if d_sel == '➕ 手动输入新司机...':
                            add_driver_modal(str(t))
                        else:
                            driver_map[t] = d_sel if d_sel != '(未选择)' else ''
                    with c_prc:
                        p_val = st.number_input(
                            '运价_' + str(t), value=curr_prc, step=1.0, format='%.2f',
                            key='p_' + str(t), label_visibility='collapsed', placeholder='¥ 运价')
                        price_map[t] = p_val
                    st.markdown('<hr style="margin:0.25em 0;border-style:dashed;border-color:#eee;"/>', unsafe_allow_html=True)

                for t in unique_trucks[:4]:
                    render_truck_row(t)
                if len(unique_trucks) > 4:
                    with st.expander('↓ 展开剩余 ' + str(len(unique_trucks)-4) + ' 辆车'):
                        for t in unique_trucks[4:]:
                            render_truck_row(t)

                new_freight_records = []
                for _, row in delivery_df.iterrows():
                    t = row['车号']
                    d_name = driver_map.get(t, '')
                    p_val  = price_map.get(t, 0.0)
                    new_freight_records.append({
                        '单号': row['单号'], '重车时间': row['重车时间'], '车号': t,
                        '收货单位': row['收货单位'], '货物名称': row['汇报专用名称'],
                        '净重': row['净重'], '司机姓名': d_name,
                        '运费单价': p_val, '运费金额': do_round(row['净重'] * p_val, 'freight'),
                    })
                new_freight_df = pd.DataFrame(new_freight_records)
                freight_total  = new_freight_df['运费金额'].sum()
                add_log('INFO', '公司配送运费合计: ' + str(round(freight_total,2)) + '元')
            else:
                f_val = st.number_input('今日无配送，如有额外运费请填写(元):', value=0.0, step=1.0, format='%.2f')
                freight_total = do_round(f_val, 'freight')

    # ============================================================
    # 财务资金登记
    # ============================================================
    with st.expander('💰 财务资金登记 (收入 & 预存)', expanded=False):
        tab_income, tab_deposit = st.tabs(['📥 收入登记', '💳 预存登记'])
        today_income_records = []

        def render_income_rows(prefix, type_options, rows_key, btn_key):
            if rows_key not in st.session_state:
                st.session_state[rows_key] = 1
            for i in range(st.session_state[rows_key]):
                c_c, c_t, c_a, c_n = st.columns([2.5, 2, 2.5, 3])
                lv = 'visible' if i == 0 else 'collapsed'
                with c_c:
                    c_sel = st.selectbox('客户名称', cust_options, key=prefix + 'c_sel_' + str(i), label_visibility=lv)
                    if c_sel == '➕ 手动输入新客户...':
                        c_name = st.text_input('新客户', key=prefix + 'c_new_' + str(i), label_visibility='collapsed', placeholder='客户名称')
                    else:
                        c_name = c_sel if c_sel != '(不录入)' else ''
                with c_t:
                    i_type = st.selectbox('类型', type_options, key=prefix + 't_' + str(i), label_visibility=lv)
                with c_a:
                    i_amt = st.number_input('金额(元)', min_value=0.0, step=1.0, format='%.2f', key=prefix + 'a_' + str(i), label_visibility=lv)
                with c_n:
                    i_note = st.text_input('备注', key=prefix + 'n_' + str(i), label_visibility=lv, placeholder='选填')
                st.markdown('<hr style="margin:0.4em 0;border-style:dashed;border-color:#eee;"/>', unsafe_allow_html=True)
                if c_name and i_amt > 0:
                    today_income_records.append({
                        '日期': report_date.strftime('%Y-%m-%d'),
                        '客户名称': c_name, '收入类型': i_type,
                        '金额': i_amt, '备注': i_note,
                    })
            col_add, _ = st.columns([1, 5])
            with col_add:
                if st.button('➕ 新增行', key=btn_key):
                    st.session_state[rows_key] += 1
                    st.rerun()

        with tab_income:
            render_income_rows('inc_', ['微信','现金','银行卡','其他'], 'income_rows', 'add_income_btn')
        with tab_deposit:
            render_income_rows('dep_', ['预存微信','预存银行卡','预存现金','预存备用金','预存其他'], 'deposit_rows', 'add_deposit_btn')

    # ── 资金余额计算 ──────────────────────────────────
    df_cash = df_sales_report[df_sales_report['过磅类型'].str.contains('现金', na=False)]
    df_wx   = df_sales_report[df_sales_report['过磅类型'].str.contains('微信', na=False)]
    df_sign = df_sales_report[df_sales_report['过磅类型'].str.contains('签单', na=False)]

    orig_bal_dict = dict(zip(df_bal['客户名称'], df_bal['余额']))
    deposit_dict  = {}

    retail_wx_amt   = do_round(df_wx['金额'].sum(),   'retail')
    retail_cash_amt = do_round(df_cash['金额'].sum(), 'retail')

    if retail_wx_amt > 0:
        today_income_records.append({'日期': report_date.strftime('%Y-%m-%d'), '客户名称': '零售客户', '收入类型': '零售微信',  '金额': retail_wx_amt,   '备注': '自动汇总'})
    if retail_cash_amt > 0:
        today_income_records.append({'日期': report_date.strftime('%Y-%m-%d'), '客户名称': '零售客户', '收入类型': '零售现金',  '金额': retail_cash_amt, '备注': '自动汇总'})

    for r in today_income_records:
        if ('预存' in r['收入类型'] or r['收入类型'] == '银行卡') and r['客户名称'] != '零售客户':
            deposit_dict[r['客户名称']] = deposit_dict.get(r['客户名称'], 0.0) + float(r['金额'])

    sign_custs = df_sign.groupby('收货单位')
    all_custs  = set(list(orig_bal_dict) + list(deposit_dict) + list(sign_custs.groups))
    bal_dict   = {}
    for c in all_custs:
        spent = do_round(df_sign[df_sign['收货单位'] == c]['金额'].sum(), 'sign') if c in sign_custs.groups else 0.0
        bal_dict[c] = orig_bal_dict.get(c, 0.0) + deposit_dict.get(c, 0.0) - spent

    daily_fee = df_daily_raw[
        (df_daily_raw['过磅类型'].astype(str).str.strip() != '')
        & (df_daily_raw['标准归类'] != '(非销售/不计价)')
    ]['加工费'].sum()

    current_month = report_date.strftime('%Y-%m')
    monthly_fee   = daily_fee
    if not df_hist.empty and '重车时间' in df_hist.columns and '加工费' in df_hist.columns:
        df_hist['加工费'] = pd.to_numeric(df_hist['加工费'], errors='coerce').fillna(0)
        monthly_fee += do_round(
            df_hist[df_hist['重车时间'].astype(str).str.startswith(current_month, na=False)]['加工费'].sum(), 'fee')

    add_log('INFO', '零售微信:' + str(retail_wx_amt) + '元 | 零售现金:' + str(retail_cash_amt) + '元 | 加工费:' + str(round(daily_fee,2)) + '元 | 运费:' + str(round(freight_total,2)) + '元')
    add_log('INFO', '签单客户余额更新: ' + str(len(bal_dict)) + ' 户')

    # ============================================================
    # 汇报文本生成
    # ============================================================
    report = report_date.strftime('%y年%m月%d日') + ' ' + report_time + '\n'

    if len(df_cash) == 0:
        report += '\n现金:无\n'
    else:
        report += '\n现金:' + str(len(df_cash)) + '车' + fmt_weight(df_cash['净重'].sum()) + '吨' + fmt_val(retail_cash_amt,'retail') + '元\n'
        for prod, grp in df_cash.groupby('汇报专用名称'):
            report += prod + ':' + str(len(grp)) + '车' + fmt_weight(grp['净重'].sum()) + '吨' + fmt_val(grp['金额'].sum(),'retail') + '元\n'

    if len(df_wx) == 0:
        report += '\n微信:无\n'
    else:
        report += '\n微信:' + str(len(df_wx)) + '车' + fmt_weight(df_wx['净重'].sum()) + '吨' + fmt_val(retail_wx_amt,'retail') + '元\n'
        for prod, grp in df_wx.groupby('汇报专用名称'):
            report += prod + ':' + str(len(grp)) + '车' + fmt_weight(grp['净重'].sum()) + '吨' + fmt_val(grp['金额'].sum(),'retail') + '元\n'

    if len(df_sign) == 0:
        report += '\n签单:无\n\n'
    else:
        report += '\n签单:' + str(len(df_sign)) + '车' + fmt_weight(df_sign['净重'].sum()) + '吨\n'
        for cust, grp in sign_custs:
            report += cust + ':' + str(len(grp)) + '车' + fmt_weight(grp['净重'].sum()) + '吨\n'
            for prod, p_grp in grp.groupby('汇报专用名称'):
                report += prod + ':' + str(len(p_grp)) + '车' + fmt_weight(p_grp['净重'].sum()) + '吨' + fmt_val(p_grp['金额'].sum(),'sign') + ' 元\n'
            cust_money = do_round(grp['金额'].sum(), 'sign')
            report += '共金额:' + fmt_val(cust_money,'sign') + ' 元\n'
            report += '上日余额:' + fmt_val(orig_bal_dict.get(cust,0.0),'sign') + ' 元\n'
            if deposit_dict.get(cust, 0.0) > 0:
                report += '今日充值:' + fmt_val(deposit_dict[cust],'sign') + ' 元\n'
            report += '当日余额:' + fmt_val(bal_dict.get(cust,0.0),'sign') + ' 元\n\n'

    pure_depositors = [c for c in deposit_dict if c not in sign_custs.groups]
    if pure_depositors:
        report += '【纯充值客户余额刷新】\n'
        for c in pure_depositors:
            report += c + ' 今日充值:' + fmt_val(deposit_dict[c],'sign') + '元 | 最新余额:' + fmt_val(bal_dict.get(c,0.0),'sign') + '元\n\n'

    total_money  = df_sales_report['金额'].sum()
    unsold_count = len(df_daily_raw) - len(df_daily_raw[df_daily_raw['过磅类型'].astype(str).str.strip() != ''])
    unsold_str   = ' (内含未销售单据/废料等 ' + str(unsold_count) + ' 车，已全量留底)' if unsold_count > 0 else ''

    report += ('1,当日销售共计:' + str(len(df_sales_report)) + ' 车' + fmt_weight(df_sales_report['净重'].sum()) + ' 吨 ' +
               fmt_val(total_money,'sign') + ' 元,公司配送运费:' + fmt_val(freight_total,'freight') + '元 ,' +
               '合计:' + fmt_val(total_money + freight_total,'sign') + ' 元。' + unsold_str + '\n')
    report += ('2,当日加工费:' + fmt_val(daily_fee,'fee') + ' 元,' +
               str(report_date.month) + '月1日-' + str(report_date.day) + '日加工费合计:' + fmt_val(monthly_fee,'fee') + ' 元。\n')

    collection_parts    = []
    custom_income_total = 0.0
    if retail_wx_amt   > 0: collection_parts.append('微信零售:' + fmt_val(retail_wx_amt,'retail') + '元')
    if retail_cash_amt > 0: collection_parts.append('现金零售:' + fmt_val(retail_cash_amt,'retail') + '元')

    for r in today_income_records:
        if r['客户名称'] != '零售客户':
            c_name = r['客户名称']
            amt    = float(r['金额'])
            custom_income_total += amt
            label  = c_name + r['收入类型'] if c_name else r['收入类型']
            collection_parts.append(label + ':' + fmt_val(amt,'none') + '元')

    total_collection = retail_wx_amt + retail_cash_amt + custom_income_total
    report += '3,当日合计收款:' + (','.join(collection_parts) if collection_parts else '0元') + ',共计:' + fmt_val(total_collection,'none') + ' 元\n'

    # ============================================================
    # 第二步：显示汇报文本
    # ============================================================
    st.divider()
    st.markdown('##### 📋 第二步：复制每日汇报')
    st.code(report, language='text', line_numbers=False)

    # ── 组装新账本 ────────────────────────────────────
    drop_cols = ['_pre_rounded', '_orig_idx', '汇报专用名称', '标准归类']
    df_daily_raw.drop(columns=[c for c in drop_cols if c in df_daily_raw.columns], inplace=True)

    new_df_bal  = pd.DataFrame(list(bal_dict.items()), columns=['客户名称','余额'])
    keep_cols   = ['单号','重车时间','车号','货物名称','净重','单价','金额','收货单位','过磅类型','备注','加工费单价','加工费','备注2']
    avail_cols  = [c for c in keep_cols if c in df_daily_raw.columns]
    new_df_hist = safe_concat([df_hist, df_daily_raw[avail_cols]])

    new_df_freight  = safe_concat([df_freight, new_freight_df])
    today_income_df = pd.DataFrame(today_income_records) if today_income_records else pd.DataFrame()
    new_df_income   = safe_concat([df_income, today_income_df])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        new_df_bal.to_excel(writer,     sheet_name='客户余额',      index=False)
        df_rules.to_excel(writer,       sheet_name='加工费规则',    index=False)
        df_mapping.to_excel(writer,     sheet_name='物料归类映射',  index=False)
        df_cust_mapping.to_excel(writer,sheet_name='客户归类映射',  index=False)
        new_df_hist.to_excel(writer,    sheet_name='过磅明细',      index=False)
        new_df_freight.to_excel(writer, sheet_name='公司配送-运费', index=False)
        new_df_income.to_excel(writer,  sheet_name='财务收入明细',  index=False)

    success_key = 'success_' + report_date.strftime('%Y%m%d%H%M%S')
    if success_key not in st.session_state:
        add_log('INFO', '✅ 账本生成成功，写入7张工作表，报表日期: ' + report_date.strftime('%Y-%m-%d'))
        st.session_state[success_key] = True

    st.success('✅ 核算完成，可下载更新后的总账本')
    col_btn, _ = st.columns([1, 2])
    with col_btn:
        st.download_button(
            label='💾 下载更新后总账本',
            data=output.getvalue(),
            file_name=report_date.strftime('%Y%m%d') + '_DiBang总账本.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )

except Exception as e:
    tb = traceback.format_exc()
    add_log('ERROR', '程序异常: ' + str(e))
    add_log('ERROR', 'Traceback:\n' + tb)
    st.error(
        '❌ 文件处理失败，请检查：\n'
        '1. 上传文件格式是否正确；\n'
        '2. 总账本是否符合模板要求。\n\n'
        '错误详情：' + str(e) + '\n\n'
        '可点击侧边栏【查看 / 导出系统日志】查看完整堆栈信息。',
        icon='⚠️',
    )

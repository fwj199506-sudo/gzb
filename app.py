import streamlit as st
import pandas as pd
import io
import re
import os

# ================= 页面配置 =================
st.set_page_config(page_title="私募资管产品估值表分析系统", layout="wide")
st.title("📊 私募资管产品估值表精算系统 (底层数组坐标版)")
st.markdown("采用二维数组坐标级定位技术，严格遵循 CAS 科目代码层级分类，解决列名错位与分类混淆问题。")

# ================= 工具函数 =================
def parse_num(val):
    """安全地将各种格式转换为浮点数"""
    if pd.isna(val):
        return 0.0
    try:
        val_str = str(val).replace(',', '').replace(' ', '').replace('，', '').strip()
        if not val_str or val_str in ['-', '--', 'None', 'nan']:
            return 0.0
        if val_str.endswith('%'):
            return float(val_str[:-1]) / 100.0
        return float(val_str)
    except:
        return 0.0

def norm_text(x):
    """统一文本格式：去空格、去全角空格、去首尾横线"""
    s = str(x).replace('\u3000', '').replace(' ', '').strip()
    s = re.sub(r'^[－\-—]+', '', s)
    return s

def clean_code_str(raw_code):
    """清洗证券代码，如 149407 SZ -> 149407.SZ；1103.06.01.196519 SH -> 1103.06.01.196519.SH"""
    code_str = str(raw_code).strip()
    if not code_str or code_str in ['nan', 'None']:
        return ""
    code_str = re.sub(r'\s+([A-Za-z]{2,4})$', r'.\1', code_str)
    return code_str

def top_level_value(row, idx_mkt=-1, idx_cost=-1):
    """
    取一行的主值：
    优先市值 -> 成本 -> 行内最大数字
    """
    vals = [parse_num(x) for x in row]
    if idx_mkt != -1:
        v = parse_num(row[idx_mkt])
        if v != 0:
            return v
    if idx_cost != -1:
        v = parse_num(row[idx_cost])
        if v != 0:
            return v
    return max(vals) if vals else 0.0

def classify_asset(code, name):
    """
    基于 CAS 科目代码层级分类
    优先级：非标特征字 > 明确代码 > 股票/债券代码特征
    """
    c = norm_text(code)
    n = norm_text(name)
    c_upper = c.upper()

    # 1) 非标优先
    if '信托' in n:
        return '信托计划'
    if '资管计划' in n or '资产管理计划' in n or '资产管理' in n:
        return '资管计划'
    if '基金' in n and '公司' not in n:
        return '公募/私募基金'

    # 2) 明确科目代码
    if c.startswith('1102.01') or c.startswith('110201'):
        return '股票'
    if c.startswith('1102.03') or c.startswith('110203'):
        return '债券'
    if c.startswith('1103'):
        return '债券'
    if c.startswith('1104') or '资产支持证券' in n or 'ABS' in c_upper:
        return '债券'
    if c.startswith('1108'):
        return '其他交易性金融资产'
    if c.startswith('1201') or c.startswith('1202') or '买入返售' in n or '逆回购' in n or '质押式' in n:
        return '买入返售(逆回购)'

    # 3) 兜底判定
    if c.startswith('1102'):
        # 常见 A 股/北交所/创业板代码特征
        if re.search(r'\b(00|30|60|68|83|87|43|92)\d{4}\b', c) or '股' in n:
            return '股票'
        if '债' in n:
            return '债券'
        return '其他交易性金融资产'

    return '未分类'

def safe_read_table(file):
    """读取 xls/xlsx/csv"""
    file_name = file.name.lower()
    if file_name.endswith('.csv'):
        encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030']
        last_err = None
        for enc in encodings:
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc, header=None, dtype=str)
            except Exception as e:
                last_err = e
        raise last_err
    else:
        # xls 需要 xlrd；xlsx 用 openpyxl
        try:
            file.seek(0)
            return pd.read_excel(file, header=None, dtype=str, engine='openpyxl')
        except Exception:
            file.seek(0)
            return pd.read_excel(file, header=None, dtype=str, engine='xlrd')

# ================= 核心解析引擎 =================
def process_valuation_files(uploaded_files):
    summary_list = []
    detail_list = []

    for file in uploaded_files:
        file_name = file.name
        base_name = os.path.splitext(file_name)[0]
        product_name = re.sub(r'(_资产估值表.*|_四级.*)$', '', base_name)

        # 1. 原始读取
        try:
            df = safe_read_table(file)
        except Exception as e:
            st.error(f"读取 {file_name} 失败: {str(e)}")
            continue

        # 2. 定位表头坐标
        header_idx = -1
        for i in range(min(30, len(df))):
            row_vals = [norm_text(x) for x in df.iloc[i].values]
            if '科目代码' in row_vals and '科目名称' in row_vals:
                header_idx = i
                break

        if header_idx == -1:
            st.warning(f"跳过 {file_name}：未能在前30行找到标准表头(科目代码/科目名称)")
            continue

        row0 = [norm_text(x) for x in df.iloc[header_idx].values]
        row1 = [norm_text(x) for x in df.iloc[header_idx + 1].values] if header_idx + 1 < len(df) else row0

        # 动态寻找列索引
        def find_idx(keys0, keys1=None):
            if keys1:
                for i in range(len(row0)):
                    if any(k in row0[i] for k in keys0) and any(k in row1[i] for k in keys1):
                        return i
            for i in range(len(row0)):
                if any(k in row0[i] for k in keys0):
                    return i
            return -1

        idx_code = find_idx(['科目代码'])
        idx_name = find_idx(['科目名称'])
        idx_qty = find_idx(['数量'])
        idx_unit_cost = find_idx(['单位成本'])
        idx_price = find_idx(['行情', '市价'])
        idx_cost = find_idx(['成本'], ['本币'])
        idx_mkt = find_idx(['市值'], ['本币'])

        if idx_cost == -1:
            idx_cost = find_idx(['成本'])
        if idx_mkt == -1:
            idx_mkt = find_idx(['市值'])
        if idx_qty == -1:
            idx_qty = 4  # 常见数量列

        # 初始化统计桶
        total_assets = 0.0
        net_assets = 0.0
        bank_deposit = 0.0
        clearing_prov = 0.0
        reverse_repo = 0.0
        stocks = 0.0
        bonds = 0.0
        others = 0.0

        # 3. 逐行解析
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i].values

            # 防止极短行
            if len(row) <= max(idx_code if idx_code != -1 else 0, idx_name if idx_name != -1 else 0):
                continue

            c_raw = str(row[idx_code] if idx_code != -1 else "").strip()
            n_raw = str(row[idx_name] if idx_name != -1 else "").strip()

            c_clean = norm_text(c_raw)
            n_clean = norm_text(n_raw)

            # 结束标志：声明部分，后面不要再解析
            if n_clean.startswith('声明'):
                break

            # 空行跳过
            if not c_clean and not n_clean:
                continue

            row_val = top_level_value(row, idx_mkt, idx_cost)

            # ================= A. 顶层总计项（精确匹配，不能用 contains） =================
            # 资产合计 / 资产净值：只认精确行，防止“其中”与“声明”污染
            if c_clean == '资产合计' or n_clean == '资产合计':
                total_assets = row_val
                continue

            if c_clean == '资产净值' or n_clean == '资产净值':
                net_assets = row_val
                continue

            # 银行存款：你要求新增这一列
            # 这类通常是 1002 顶层行
            if c_clean == '1002' or n_clean == '银行存款':
                bank_deposit = row_val
                continue

            # 结算备付金：只取顶层 1021，避免把 1021.81、1021.B1 等子项重复加总
            if c_clean == '1021' or n_clean == '结算备付金':
                clearing_prov = row_val
                continue

            # 买入返售金融资产：你这份表是 1202，不是 1201
            if c_clean in ['1201', '1202'] or n_clean == '买入返售金融资产':
                reverse_repo = row_val
                continue

            # ================= B. 底层持仓穿透（有数量才进入明细） =================
            v_qty = parse_num(row[idx_qty]) if idx_qty != -1 else 0.0

            # 只有真正有数量的明细才进入穿透统计
            if v_qty > 0:
                asset_type = classify_asset(c_raw, n_raw)

                # 排除明显汇总行
                if any(x in n_clean for x in ['汇总', '合计', '大类', '交易所', '深交所', '上交所']):
                    continue

                # 明细表：只记录资产类持仓
                if c_clean.startswith(('1102', '1103', '1104', '1108', '1201', '1202')):
                    clean_code = "逆回购" if asset_type == '买入返售(逆回购)' else clean_code_str(c_raw)

                    detail_list.append({
                        "所属产品": product_name,
                        "资产类别": asset_type,
                        "代码": clean_code,
                        "资产名称": n_raw,
                        "数量": v_qty,
                        "单位成本": parse_num(row[idx_unit_cost]) if idx_unit_cost != -1 else 0.0,
                        "总成本": parse_num(row[idx_cost]) if idx_cost != -1 else 0.0,
                        "今日行情": parse_num(row[idx_price]) if idx_price != -1 else 0.0,
                        "今日市值": parse_num(row[idx_mkt]) if idx_mkt != -1 else 0.0
                    })

                # 汇总桶
                v_val = parse_num(row[idx_mkt]) if idx_mkt != -1 and parse_num(row[idx_mkt]) != 0 else parse_num(row[idx_cost])

                if asset_type == '股票':
                    stocks += v_val
                elif asset_type == '债券':
                    bonds += v_val
                elif asset_type == '买入返售(逆回购)':
                    reverse_repo += v_val
                elif asset_type in ['信托计划', '资管计划', '公募/私募基金', '其他交易性金融资产']:
                    others += v_val

        # 单个产品解析完毕，压入汇总表
        summary_list.append({
            "产品名称": product_name,
            "银行存款": bank_deposit,
            "结算备付金及保证金": clearing_prov,
            "股票投资金额": stocks,
            "债券投资金额": bonds,
            "买入返售金融资产": reverse_repo,
            "其他交易性金融资产": others,
            "总资产": total_assets,
            "净资产": net_assets
        })

    df_sum = pd.DataFrame(summary_list)
    df_det = pd.DataFrame(detail_list)
    return df_sum, df_det

# ================= 交互渲染 =================
uploaded_files = st.file_uploader(
    "📂 请上传估值表文件 (多选 Excel/CSV)",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("引擎正在进行二维坐标矩阵重构与底层资产穿透..."):
        df_sum, df_det = process_valuation_files(uploaded_files)

    st.success(f"成功解析 {len(uploaded_files)} 个产品！")

    st.subheader("📋 Sheet 1: 产品资产概况")
    df_sum_show = df_sum.copy()

    # 统一展示格式
    for col in df_sum_show.columns:
        if col != "产品名称":
            df_sum_show[col] = df_sum_show[col].apply(lambda x: f"{x:,.2f}")

    st.dataframe(df_sum_show, use_container_width=True)

    st.subheader("🔍 Sheet 2: 底层资产明细")
    df_det_show = df_det.copy()
    if not df_det_show.empty:
        for col in ['数量', '单位成本', '总成本', '今日行情', '今日市值']:
            if col in df_det_show.columns:
                df_det_show[col] = df_det_show[col].apply(lambda x: f"{x:,.2f}")

    st.dataframe(df_det_show, use_container_width=True)

    # 导出文件
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_sum.to_excel(writer, sheet_name='产品概况', index=False)
        df_det.to_excel(writer, sheet_name='资产明细', index=False)

    st.markdown("---")
    st.download_button(
        label="📥 点击下载精算级汇总 Excel 报表",
        data=output.getvalue(),
        file_name="私募产品多维度估值透视表_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("提示：支持批量拖入估值表。算法已升级为精确总计行识别 + 底层明细穿透，避免子项重复累计。")
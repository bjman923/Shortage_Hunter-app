import streamlit as st
import pandas as pd
import os
import json
import re
from datetime import date, timedelta

# ==========================================
# 1. 網頁基本設定
# ==========================================
st.set_page_config(page_title="電池模組缺料分析系統", layout="wide", page_icon="🔋", initial_sidebar_state="expanded")

# ==========================================
# 2. 全域變數與存檔設定
# ==========================================
FILES = {
    "bom": "缺料預估.xlsx",       
    "stock_w08": "庫存明細表.xlsx", 
    "stock_w26": "W26庫存明細表.xlsx"
}
PLAN_FILE = "schedule.json"

if 'read_errors' not in st.session_state: st.session_state.read_errors = {}
if 'debug_logs' not in st.session_state: st.session_state.debug_logs = []

missing = []
for k, f in FILES.items():
    if not os.path.exists(f): missing.append(f)

individual_w08 = {} 
individual_w26 = {}

def rerun_app():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def load_plan():
    if os.path.exists(PLAN_FILE):
        try:
            with open(PLAN_FILE, 'r', encoding='utf-8') as f: return json.load(f)
        except: return []
    return []

def save_plan(data):
    with open(PLAN_FILE, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False)

# ==========================================
# 3. CSS 樣式 (v134.0 品號寬度改為 220px)
# ==========================================
st.markdown("""
<style>
    /* 基礎鎖定 */
    html, body { height: 100vh !important; width: 100vw !important; overflow: hidden !important; font-family: 'Microsoft JhengHei', sans-serif !important; }
    div[data-testid="stAppViewContainer"] { height: 100dvh !important; overflow: hidden !important; width: 100% !important; }
    .main .block-container { padding: 5px !important; max-width: 100% !important; overflow: hidden !important; }
    
    /* footer 隱藏 */
    footer { display: none !important; }

    /* 手機版專屬優化 */
    @media screen and (max-width: 768px) {
        /* Header 顯示設定 (為了側邊欄按鈕) */
        header[data-testid="stHeader"] { 
            display: block !important; 
            background-color: white !important; 
            height: 45px !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
        }
        header[data-testid="stHeader"] button {
            color: black !important;
        }

        section[data-testid="stSidebar"] { z-index: 999999 !important; box-shadow: 2px 0 10px rgba(0,0,0,0.2) !important; }
        
        .app-title { font-size: 18px !important; margin-bottom: 5px !important; white-space: nowrap !important; margin-top: 0px !important; }
        .kpi-container { height: 60px !important; padding: 2px !important; margin-bottom: 5px; background: white; border-radius: 8px; border-left: 4px solid #2c3e50; text-align: center; }
        .kpi-title { font-size: 11px !important; margin: 0; color: #7f8c8d; }
        .kpi-value { font-size: 20px !important; font-weight: 700; color: #2c3e50; }

        /* 表格容器 */
        .table-wrapper { 
            width: 100%; 
            height: calc(100dvh - 200px) !important; 
            overflow: auto !important; 
            margin-top: 5px !important; 
            background: white;
            -webkit-overflow-scrolling: touch; 
        }
        
        /* 寬度設定 */
        table { 
            width: auto !important; 
            min-width: 800px !important; 
            border-collapse: separate; 
            border-spacing: 0; 
            table-layout: fixed !important; 
        }
        
        thead tr th { 
            position: sticky; top: 0; z-index: 50; 
            background-color: #2c3e50; color: white; 
            font-size: 13px !important; 
            padding: 8px 4px !important; 
            white-space: nowrap !important; 
            text-align: center !important;
            border-bottom: 1px solid #ddd;
        }
        
        tbody tr td, 
        tbody tr td > div, 
        tbody tr td > span, 
        tbody tr td > details > summary { 
            font-size: 13px !important; 
            padding: 8px 4px !important; 
            white-space: nowrap !important; /* 絕對不換行 */
            overflow: hidden !important; 
            text-overflow: clip !important; 
            vertical-align: middle !important;
            height: 35px !important;
            line-height: 20px !important;
        }

        details[open] > div {
            white-space: normal !important; 
            height: auto !important;
            overflow: visible !important;
        }
        
        [data-testid="stSidebar"] button { padding: 0px 5px !important; height: 35px !important; font-size: 14px !important; }
    }
    
    /* 電腦版樣式 */
    @media screen and (min-width: 769px) {
        header[data-testid="stHeader"] { display: none !important; }
        .table-wrapper { height: calc(100vh - 260px) !important; overflow: auto; }
        table { min-width: 1000px !important; }
        tbody tr td { font-size: 16px !important; white-space: nowrap !important; }
    }

    /* ★★★ 欄位寬度微調 (品號改為 220px) ★★★ */
    tbody tr td:nth-child(1) { min-width: 60px; text-align: center; }
    tbody tr td:nth-child(2) { min-width: 150px; text-align: left !important; }
    tbody tr td:nth-child(3) { min-width: 80px; text-align: center !important; }
    tbody tr td:nth-child(4) { min-width: 220px; text-align: left; overflow: visible !important; } /* 改為 220px */
    tbody tr td:nth-child(5) { min-width: 200px; text-align: left !important; }
    tbody tr td:nth-child(6) { min-width: 60px; text-align: center !important; }
    tbody tr td:nth-child(7) { min-width: 80px; text-align: center !important; }
    tbody tr td:nth-child(8) { min-width: 80px; text-align: center !important; }
    tbody tr td:nth-child(9) { min-width: 80px; text-align: center !important; }
    tbody tr td:nth-child(10) { min-width: 80px; text-align: center !important; }

    .badge { padding: 2px 6px; border-radius: 4px; font-size: 12px; color: white; font-weight: bold; }
    .badge-ok { background-color: #27ae60; }
    .badge-err { background-color: #c0392b; }
    .sim-table { width: 100%; border: 1px solid #ddd; margin-top: 5px; background: #f9f9f9; }
    .sim-table td { white-space: nowrap !important; }
    .sim-row-short { background-color: #ffebee; color: #c0392b; font-weight: bold; }
    .sim-row-supply { background-color: #e8f5e9; color: #2e7d32; font-weight: bold; }
    div[data-testid="stForm"] button { width: 100%; border-radius: 8px; font-weight: bold; margin-top: 0px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 4. 核心函數
# ==========================================
def get_base_part_no(raw_no):
    s = str(raw_no).strip()
    if len(s) > 0 and s[0] in '0123456789': s = "TW" + s
    if '-' in s: return s.split('-')[0]
    return s

def normalize_key(part_no):
    if pd.isna(part_no): return ""
    s = str(part_no).upper().strip()
    s = s.replace("TW", "").replace("-", "").replace(" ", "")
    return s

def read_excel_auto_header(file_path):
    if not os.path.exists(file_path): return pd.DataFrame()
    try:
        df_preview = pd.read_excel(file_path, header=None, nrows=10, engine='openpyxl')
        target_row = 0
        found = False
        for idx, row in df_preview.iterrows():
            row_str = " ".join(row.astype(str).values)
            if "品號" in row_str: target_row = idx; found = True; break
        return pd.read_excel(file_path, header=target_row, engine='openpyxl')
    except Exception as e:
        st.session_state.read_errors[file_path] = str(e)
        return pd.DataFrame()

def clean_df(df):
    if df.empty: return df
    df.columns = [str(c).strip() for c in df.columns]
    part_col = next((c for c in df.columns if '品號' in c), None)
    if part_col:
        df = df.dropna(subset=[part_col]) 
        df = df[~df[part_col].astype(str).str.contains('小計|合計|總計', na=False)]
    return df

def load_data(files):
    st.session_state.read_errors = {}
    df_bom = read_excel_auto_header(files["bom"])
    df_w08 = read_excel_auto_header(files["stock_w08"])
    df_w26 = read_excel_auto_header(files["stock_w26"])
    if df_w26.empty and files["stock_w26"] not in st.session_state.read_errors:
        st.session_state.debug_logs.append(f"⚠️ {files['stock_w26']} 內容為空或讀取失敗")
    return clean_df(df_bom), clean_df(df_w08), clean_df(df_w26)

def process_mps_file(uploaded_file):
    mps_list = []
    log_msg = []
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        date_col = next((c for c in df.columns if 'Date' in str(c) or '日期' in str(c)), None)
        if not date_col: return [], ["❌ 找不到 [Date] 欄位"]
        
        target_cols = []
        for c in df.columns:
            clean_c = str(c).replace('\n', '').replace(' ', '')
            if '計畫' in clean_c and '產出' in clean_c:
                model_name = clean_c.replace('計畫', '').replace('產出', '').strip()
                if model_name: target_cols.append({'col': c, 'model': model_name})
        if not target_cols: return [], ["⚠️ 找不到任何 [計畫產出] 欄位"]

        today = date.today()
        cutoff_date = today + timedelta(days=1)
        count = 0
        skip_count = 0
        
        for _, row in df.iterrows():
            try:
                raw_date = row[date_col]
                dt_obj = pd.to_datetime(raw_date)
                if dt_obj.date() < cutoff_date:
                    skip_count += 1
                    continue
                plan_date_str = dt_obj.strftime('%Y-%m-%d')
                for t in target_cols:
                    qty = row[t['col']]
                    if pd.notna(qty):
                        try:
                            qty_val = float(qty)
                            if qty_val > 0:
                                mps_list.append({'日期': plan_date_str, '型號': t['model'], '數量': int(qty_val), 'source': 'MPS'})
                                count += 1
                        except: pass
            except: continue
        log_msg.append(f"✅ 匯入 {count} 筆 (已過濾 {skip_count} 筆舊資料)")
        return mps_list, log_msg
    except Exception as e:
        return [], [f"❌ MPS 讀取失敗: {str(e)}"]

def process_supplier_uploads(uploaded_files):
    supply_list = []
    log_msg = []
    if not uploaded_files: return [], []
    for up_file in uploaded_files:
        try:
            df_raw = pd.read_excel(up_file, header=None, engine='openpyxl')
            header_row_idx = -1; part_col_idx = -1
            for r in range(min(15, len(df_raw))):
                row_vals = df_raw.iloc[r].astype(str).values
                for c, val in enumerate(row_vals):
                    if "品號" in val: header_row_idx = r; part_col_idx = c; break
                if header_row_idx != -1: break
            if header_row_idx == -1: log_msg.append(f"❌ {up_file.name}: 未偵測到品號欄"); continue
            date_col_map = {}
            scan_start = max(0, header_row_idx - 1)
            scan_end = min(len(df_raw), header_row_idx + 6)
            for r in range(scan_start, scan_end):
                temp_map = {}
                for c in range(len(df_raw.columns)):
                    val = df_raw.iloc[r, c]
                    try:
                        dt = pd.to_datetime(val, errors='coerce')
                        if pd.notna(dt): temp_map[c] = dt.strftime('%Y-%m-%d')
                    except: continue
                if temp_map: date_col_map = temp_map; break
            if not date_col_map: log_msg.append(f"⚠️ {up_file.name}: 未偵測到日期欄"); continue
            data_start_row = header_row_idx + 1
            count = 0
            for r in range(data_start_row, len(df_raw)):
                try:
                    p_no = str(df_raw.iloc[r, part_col_idx]).strip()
                    if not p_no or p_no.lower() == 'nan': continue
                    for c_idx, date_str in date_col_map.items():
                        qty_val = df_raw.iloc[r, c_idx]
                        try:
                            qty = float(qty_val)
                            if qty > 0:
                                supply_list.append({'date': date_str, 'type': 'supply', 'note': "🚛 到貨", 'part_no': p_no, 'match_key': normalize_key(p_no), 'qty': qty})
                                count += 1
                        except: continue
                except: continue
            log_msg.append(f"✅ {up_file.name}: {count} 筆")
        except Exception as e: log_msg.append(f"❌ {up_file.name}: {str(e)}")
    return supply_list, log_msg

def process_stock(df, store_type):
    if df.empty: return
    try:
        candidates = [c for c in df.columns if '數量' in c]
        stock_cols = [c for c in candidates if '庫存' in c]
        col_q = stock_cols[0] if stock_cols else (candidates[0] if candidates else None)
        if not col_q: return
        col_p = next(c for c in df.columns if '品號' in c)
        if store_type == 'W08':
            col_wh = next((c for c in df.columns if '庫別' in c), None)
            if col_wh: df = df[df[col_wh].astype(str).str.strip() == 'W08']
        df[col_q] = pd.to_numeric(df[col_q], errors='coerce').fillna(0)
        for _, row in df.iterrows():
            raw_p = str(row[col_p]).strip()
            stock_base = get_base_part_no(raw_p)
            qty = row[col_q]
            if store_type == 'W08': individual_w08[stock_base] = individual_w08.get(stock_base, 0) + qty
            else: individual_w26[stock_base] = individual_w26.get(stock_base, 0) + qty
    except: pass

def render_grouped_html_table(grouped_data):
    html = '<div class="table-wrapper"><table style="width:100%;">'
    html += """
    <colgroup>
        <col style="width: 60px">   <col style="width: 150px">  <col style="width: 80px">   <col style="width: 220px">  <col style="width: 200px">  <col style="width: 60px">   <col style="width: 80px">   <col style="width: 80px">   <col style="width: 80px">   <col style="width: 80px">   </colgroup>
    <thead><tr><th>狀態</th><th>首個斷料點</th><th>型號</th><th>品號 / 群組內容</th><th>品名</th><th>用量</th><th>W08</th><th>W26</th><th>總需求</th><th>最終結餘</th></tr></thead><tbody>
    """
    def fmt(n): return f"{int(n):,}"
    for group in grouped_data:
        is_short = group['final_balance'] < 0
        count = len(group['items'])
        is_group = count > 1
        tr_style = 'color: #333;'
        bg_class = 'background-color: #FFEBEE;' if is_short else 'background-color: white;'
        if is_short: tr_style = 'color: #c0392b; font-weight: 500;'
        
        html += f'<tr style="{tr_style} {bg_class}">'
        status_html = '<span class="badge badge-err">缺料</span>' if is_short else '<span class="badge badge-ok">充足</span>'
        html += f'<td class="text-center">{status_html}</td>'
        
        first_shortage = group.get('first_shortage_info', '-')
        shortage_style = "color: #c0392b; font-weight: bold;" if is_short else "color: #aaa;"
        html += f'<td class="text-center" style="{shortage_style}">{first_shortage}</td>'
        
        html += f'<td class="text-center" style="text-align: center !important;">{group["model"]}</td>'
        
        if is_group or group['simulation_logs']:
            details_inner = ""
            if is_group:
                for item in group['items']:
                    details_inner += f'<div style="border-bottom:1px dashed #ccc; padding:6px 0;"><div><span style="color:#444; font-weight:bold;">{item["p_no"]}</span></div><div style="font-size:14px; color:#555;">W08:<b>{fmt(item["w08"])}</b> | W26:<b>{fmt(item["w26"])}</b></div></div>'
            sim_table_html = ""
            if group['simulation_logs']:
                sim_rows = ""
                for log in group['simulation_logs']:
                    row_cls = "sim-row-supply" if log['type'] == 'supply' else ("sim-row-short" if log['balance'] < 0 else "")
                    qty_display = f"+{fmt(log['qty'])}" if log['type'] == 'supply' else f"-{fmt(log['qty'])}"
                    sim_rows += f'<tr class="{row_cls}"><td>{log["date"]}</td><td>{log["note"]}</td><td style="text-align:center;">{qty_display}</td><td style="text-align:center;">{fmt(log["balance"])}</td></tr>'
                sim_table_html = f"""<div class="sim-wrapper" style="margin-top: 10px;"><b style="color:#2c3e50;">📅 MRP模擬：</b><table class="sim-table"><thead><tr><th>日期</th><th>摘要</th><th>變動</th><th>結餘</th></tr></thead><tbody>{sim_rows}</tbody></table></div>"""
            summary_text = f"📦 共用料 ({count})" if is_group else group['items'][0]['p_no']
            details_box = f'<div style="font-size:14px; margin-top:5px; padding-left:5px; border-left:3px solid #ddd;">{details_inner}{sim_table_html}</div>'
            html += f'<td><details><summary>{summary_text}</summary>{details_box}</details></td>'
        else:
            html += f'<td>{group["items"][0]["p_no"]}</td>'

        html += f'<td style="text-align: left !important;">{group["items"][0]["name"]}</td>'
        
        usage = max([i['usage'] for i in group['items']])
        html += f'<td class="text-center" style="text-align: center !important;"><span class="num-font">{usage}</span></td>'
        html += f'<td class="text-center" style="text-align: center !important;"><span class="num-font">{fmt(group["total_w08"])}</span></td>'
        html += f'<td class="text-center" style="text-align: center !important;"><span class="num-font">{fmt(group["total_w26"])}</span></td>'
        html += f'<td class="text-center" style="text-align: center !important;"><span class="num-font">{fmt(group["total_demand"])}</span></td>'
        html += f'<td class="text-center" style="text-align: center !important;"><span class="num-font">{fmt(group["final_balance"])}</span></td></tr>'
    html += '</tbody></table></div>'
    return html

# ==========================================
# 5. 主程式流程
# ==========================================
df_bom_src, df_w08_src, df_w26_src = load_data(FILES)

if df_bom_src is not None:
    if 'plan' not in st.session_state: st.session_state.plan = load_plan()

    try:
        c_model = next(c for c in df_bom_src.columns if '型號' in c)
        c_part = next(c for c in df_bom_src.columns if '品號' in c)
        c_code = next((c for c in df_bom_src.columns if '項目' in c or '代號' in c), None)
        c_name = next((c for c in df_bom_src.columns if '品名' in c), None)
        c_usage = next((c for c in df_bom_src.columns if '用量' in c), None)
    except: st.error("BOM 表欄位偵測失敗"); st.stop()

    if c_code:
        df_bom_src[c_code] = df_bom_src[c_code].fillna('').astype(str)
        df_bom_src['_sort_num'] = df_bom_src[c_code].str.extract('(\d+)').astype(float).fillna(0)
        df_bom_sorted = df_bom_src.sort_values(by=[c_model, '_sort_num', c_part])
    else:
        df_bom_sorted = df_bom_src.sort_values(by=[c_model, c_part])

    unique_models = df_bom_sorted[c_model].dropna().unique().tolist()
    
    with st.sidebar:
        if missing: st.error("⚠️ 檔案缺失！" + str(missing)); st.stop()
        st.header("1. 供應商交期")
        supplier_files = st.file_uploader("上傳供應商 Excel", accept_multiple_files=True, type=['xlsx', 'xls'], key="sup_uploader")
        if supplier_files:
            s_list, s_logs = process_supplier_uploads(supplier_files)
            with st.expander("📊 讀取結果診斷", expanded=False):
                for log in s_logs:
                    if "❌" in log: st.error(log)
                    else: st.success(log)
        else: s_list = []

        st.markdown("---")
        st.header("2. 生產排程")
        
        mps_file = st.file_uploader("📂 上傳排程計畫 (xlsx)", type=['xlsx', 'xls'])
        mps_data = []
        if mps_file:
            mps_data, mps_logs = process_mps_file(mps_file)
            for log in mps_logs:
                if "❌" in log: st.error(log)
                else: st.success(log)

        st.markdown("---")
        st.markdown("**或 手動輸入插單：**")
        with st.form("add_plan"):
            date_in = st.date_input("生產日期", value=date.today())
            m_sel = st.selectbox("選擇型號", unique_models)
            q_str = st.text_input("數量", value="1000")
            if st.form_submit_button("➕ 加入排程"):
                try: q_in = int(q_str)
                except: q_in = 0
                if q_in > 0:
                    st.session_state.plan.append({'日期': date_in.strftime('%Y-%m-%d'), '型號': m_sel, '數量': q_in})
                    save_plan(st.session_state.plan); rerun_app()
        
        if st.session_state.plan:
            st.markdown("###### 📝 手動插單列表")
            sorted_plan = sorted(enumerate(st.session_state.plan), key=lambda x: x[1]['日期'])
            for original_idx, item in sorted_plan:
                c1, c2 = st.columns([5, 1])
                d_str = pd.to_datetime(item['日期']).strftime('%m/%d')
                info_text = f"**{d_str}** | <small>{item['型號']}</small> | **{item['數量']:,}**"
                with c1: st.markdown(info_text, unsafe_allow_html=True)
                with c2:
                    if st.button("✖", key=f"del_{original_idx}"):
                        st.session_state.plan.pop(original_idx)
                        save_plan(st.session_state.plan); rerun_app()
                st.markdown("<hr style='margin: 2px 0; border-top: 1px dashed #eee;'>", unsafe_allow_html=True)
            if st.button("🗑️ 清空手動排程"): st.session_state.plan = []; save_plan([]); rerun_app()
            
        if 'W26庫存明細表.xlsx' in st.session_state.read_errors:
            st.error(f"🔴 W26 讀取失敗！原因：\n{st.session_state.read_errors['W26庫存明細表.xlsx']}")

    process_stock(df_w08_src, 'W08')
    process_stock(df_w26_src, 'W26')

    ledger = {} 
    total_plan_qty = 0
    active_models = [] 
    
    all_plans = []
    if st.session_state.plan:
        for p in st.session_state.plan: p['source'] = '手動'; all_plans.append(p)
    if mps_data: all_plans.extend(mps_data)

    if all_plans:
        sorted_plan_data = sorted(all_plans, key=lambda x: x['日期'])
        active_models = list(set([p['型號'] for p in sorted_plan_data]))
        for item in sorted_plan_data:
            plan_date, plan_model, plan_qty = item['日期'], item['型號'], item['數量']
            source_note = "MPS" if item.get('source') == 'MPS' else "手動"
            total_plan_qty += plan_qty
            model_bom = df_bom_sorted[df_bom_sorted[c_model] == plan_model]
            model_reqs = {}
            for _, r in model_bom.iterrows():
                p_no = str(r[c_part]).strip()
                norm_k = normalize_key(p_no)
                try: usage = float(r.get(c_usage, 0))
                except: usage = 0
                if usage > model_reqs.get(norm_k, 0): model_reqs[norm_k] = usage
            for k, u in model_reqs.items():
                if k not in ledger: ledger[k] = []
                ledger[k].append({'date': plan_date, 'type': 'demand', 'note': f"生產({source_note}): {plan_model}", 'qty': plan_qty * u})

    normalized_map = {}
    for k in ledger.keys():
        norm_k = normalize_key(k)
        if norm_k not in normalized_map: normalized_map[norm_k] = []
        normalized_map[norm_k].append(k)

    if s_list:
        for s in s_list:
            sup_norm_key = s['match_key']
            if sup_norm_key in normalized_map:
                for target_key in normalized_map[sup_norm_key]:
                    ledger[target_key].append(s)
            else:
                if s['part_no'] not in ledger: ledger[s['part_no']] = []
                ledger[s['part_no']].append(s)

    st.markdown(f'<h2 class="app-title">🔋 電池模組缺料分析系統</h2>', unsafe_allow_html=True)

    c_filter, c_search_no, c_search_name = st.columns([1, 1, 1])
    with c_filter: sel_filter = st.selectbox("🔍 篩選機種", ["全部顯示"] + unique_models)
    with c_search_no: search_no = st.text_input("搜尋品號 (Part No.)", "")
    with c_search_name: search_name = st.text_input("搜尋品名 (Name)", "")
    
    if sel_filter == "全部顯示": target_df = df_bom_sorted[df_bom_sorted[c_model].isin(active_models)] if active_models else df_bom_sorted
    else: target_df = df_bom_sorted[df_bom_sorted[c_model] == sel_filter]

    grouped_data = [] 
    current_group_key = None
    current_group_data = None
    
    for _, row in target_df.iterrows():
        p_no = str(row[c_part]).strip()
        bom_base = get_base_part_no(p_no)
        p_code = str(row.get(c_code, '')).strip()
        model = row[c_model]
        if p_code and p_code.lower() != 'nan': group_key = (model, p_code)
        else: group_key = (model, p_no)

        my_w08 = individual_w08.get(bom_base, 0)
        my_w26 = individual_w26.get(bom_base, 0)
        item_data = {'p_no': p_no, 'base': bom_base, 'name': row.get(c_name, ''), 'usage': float(row.get(c_usage, 0)), 'w08': my_w08, 'w26': my_w26, 'net_stock': my_w08 + my_w26}

        if group_key != current_group_key:
            if current_group_data: grouped_data.append(current_group_data)
            current_group_key = group_key
            current_group_data = {'model': model, 'code': p_code, 'items': [item_data], 'req_key': p_code if (p_code and p_code.lower()!='nan') else p_no, 'total_w08': my_w08, 'total_w26': my_w26, 'total_net': my_w08 + my_w26}
        else:
            current_group_data['items'].append(item_data)
            current_group_data['total_w08'] += my_w08
            current_group_data['total_w26'] += my_w26
            current_group_data['total_net'] += (my_w08 + my_w26)
    if current_group_data: grouped_data.append(current_group_data)

    processed_list = []
    shortage_count = 0
    total_items = 0
    
    for g in grouped_data:
        match_no = True if not search_no else any(search_no.lower() in i['p_no'].lower() for i in g['items'])
        match_name = True if not search_name else any(search_name.lower() in i['name'].lower() for i in g['items'])
        
        if match_no and match_name:
            running_balance = g['total_net']
            total_demand = 0
            first_shortage_info = "-"
            simulation_logs = []
            unique_demands = {} 
            supplies = []
            
            for item in g['items']:
                k = normalize_key(item['p_no'])
                if k in ledger:
                    for entry in ledger[k]:
                        if entry['type'] == 'demand':
                            d_key = (entry['date'], entry['note'])
                            if d_key not in unique_demands: unique_demands[d_key] = entry['qty']
                            else: unique_demands[d_key] = max(unique_demands[d_key], entry['qty'])
                        else: supplies.append(entry)
            
            movements = supplies + [{'date': k[0], 'note': k[1], 'type': 'demand', 'qty': v} for k, v in unique_demands.items()]
            movements.sort(key=lambda x: x['date'])
            
            for m in movements:
                if m['type'] == 'demand':
                    running_balance -= m['qty']
                    if m['qty'] > 0: total_demand += m['qty']
                    if running_balance < 0 and first_shortage_info == "-":
                        first_shortage_info = f"{m['date']} ({m['note']})"
                elif m['type'] == 'supply': running_balance += m['qty']
                simulation_logs.append({'date': m['date'], 'note': m['note'], 'type': m['type'], 'qty': m['qty'], 'balance': running_balance})
            
            g['total_demand'] = total_demand
            g['final_balance'] = running_balance
            g['first_shortage_info'] = first_shortage_info
            g['simulation_logs'] = simulation_logs
            if g['final_balance'] < 0: shortage_count += 1
            total_items += 1
            processed_list.append(g)

    if 'show_shortage_only' not in st.session_state: st.session_state.show_shortage_only = False
    def toggle_shortage_view(): st.session_state.show_shortage_only = not st.session_state.show_shortage_only

    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(f"""<div class="kpi-container"><div class="kpi-title">物料項目數</div><div class="kpi-value">{total_items}</div></div>""", unsafe_allow_html=True)
    with c2:
        if st.session_state.show_shortage_only: btn_label = f"🔙 顯示全部\n(目前: {shortage_count} 項缺料)"
        else: btn_label = f"🔥 缺料項目: {shortage_count}\n(點擊只看缺料)"
        st.button(btn_label, on_click=toggle_shortage_view)
    with c3: st.markdown(f"""<div class="kpi-container"><div class="kpi-title">計畫生產總數</div><div class="kpi-value">{total_plan_qty}</div></div>""", unsafe_allow_html=True)

    final_display_list = []
    if st.session_state.show_shortage_only: final_display_list = [g for g in processed_list if g['final_balance'] < 0]
    else: final_display_list = processed_list

    if final_display_list: st.markdown(render_grouped_html_table(final_display_list), unsafe_allow_html=True)
    else:
        if st.session_state.show_shortage_only: st.success("🎉 目前沒有任何缺料項目！")
        else:
            if active_models: st.info("查無符合條件的資料")
            else: st.info("💡 請在左側輸入排程，或選擇「全部顯示」查看所有 BOM。")

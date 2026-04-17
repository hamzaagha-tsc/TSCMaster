import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | Revenue Attribution Portal", layout="wide")

# --- UI STYLING ---
brand_navy, brand_copper = "#102a51", "#c59d5f"
st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy} !important; }}
    [data-testid="stSidebar"] * {{ color: white !important; font-family: 'Arial'; }}
    div.stButton > button:first-child {{ background-color: {brand_copper}; color: white; border-radius: 8px; font-weight: bold; }}
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    </style>
    """, unsafe_allow_html=True)

# --- ROBUST UTILITIES ---
def clean_phone_master(p):
    if pd.isna(p) or str(p).strip() == "": return None
    try:
        if isinstance(p, float): s = format(p, '.0f')
        else: s = str(p).split('.')[0].strip()
        digits = re.sub(r'\D', '', s)
        return digits[-10:] if len(digits) >= 10 else None
    except: return None

def hms_to_sec(t):
    if pd.isna(t) or t == '0' or t == 0: return 0
    s = str(t).strip().replace('.', ':')
    parts = s.split(':')
    try:
        if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        if len(parts) == 2: return int(parts[0])*60 + int(parts[1])
        return 0
    except: return 0

def sec_to_hms(s):
    h = int(s // 3600); m = int((s % 3600) // 60); s = int(s % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def robust_date_parse(s):
    if pd.isna(s): return pd.NaT
    # Tries multiple formats found in TSC reports
    for fmt in ["%Y-%m-%d %H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d"]:
        try: return pd.to_datetime(s, format=fmt)
        except: continue
    return pd.to_datetime(s, dayfirst=True, errors='coerce')

# --- CORE LOGIC ---
def run_revenue_automation(o_raw, call_files_list):
    # 1. AGGREGATE ALL CALLS
    calls_list = []
    for f in call_files_list:
        df = pd.read_csv(f)
        df.columns = df.columns.str.strip()
        
        # Dynamic Header Detection
        time_c = next((c for c in df.columns if 'time' in c.lower() and 'hangup' not in c.lower()), None)
        phone_c = next((c for c in df.columns if 'phone' in c.lower() or 'dst' in c.lower()), None)
        talk_c = next((c for c in df.columns if 'talk time' in c.lower()), None)
        user_c = next((c for c in df.columns if 'user id' in c.lower() or 'username' in c.lower()), None)
        
        temp = df[[time_c, phone_c, talk_c, user_c]].copy()
        temp.columns = ['Raw_Time', 'Raw_Phone', 'Raw_Talk', 'Agent']
        calls_list.append(temp)
    
    all_calls = pd.concat(calls_list, ignore_index=True)
    all_calls['DT'] = all_calls['Raw_Time'].apply(robust_date_parse)
    all_calls['Phone_Clean'] = all_calls['Raw_Phone'].apply(clean_phone_master)
    all_calls['Sec'] = all_calls['Raw_Talk'].apply(hms_to_sec)
    all_calls = all_calls.dropna(subset=['DT', 'Phone_Clean'])

    # 2. PROCESS ORDERS
    o_raw.columns = o_raw.columns.str.strip()
    orders = o_raw.dropna(subset=['Total']).copy()
    time_col = next(c for c in orders.columns if 'created' in c.lower())
    
    orders['DT_Clean'] = orders[time_col].astype(str).str.replace(r'\s\+\d{4}$', '', regex=True)
    orders['DT_Obj'] = pd.to_datetime(orders['DT_Clean'])
    
    orders['B_Ph'] = orders['Billing Phone'].apply(clean_phone_master)
    orders['S_Ph'] = orders['Shipping Phone'].apply(clean_phone_master)
    
    final_output = []
    for _, row in orders.iterrows():
        tags = str(row.get('Tags', '')).lower()
        chan = "Store" if "retail" in tags and "store" in tags else "Website"
        phones = {p for p in [row['B_Ph'], row['S_Ph']] if p}
        
        for ph in phones:
            # Attribution Match
            matches = all_calls[(all_calls['Phone_Clean'] == ph) & (all_calls['DT'] < row['DT_Obj'])]
            base = {
                'Order ID': row['Name'], 'Order Value': float(row['Total']),
                'Order Date': row['DT_Obj'].strftime('%d-%m-%Y'),
                'Order Time': row['DT_Obj'].strftime('%d-%m-%Y %H:%M:%S'),
                'Order Phone': ph, 'Channel': chan
            }

            if matches.empty:
                final_output.append({**base, 'Agent': 'Organic', 'Attributed Revenue': row['Total'], 'Talk Time': '00:00:00'})
                continue

            agent_stats = matches.groupby('Agent')['Sec'].sum().reset_index()

            if row['Total'] >= 100000:
                qual = agent_stats[agent_stats['Sec'] >= 180]
                if qual.empty:
                    final_output.append({**base, 'Agent': 'Organic', 'Attributed Revenue': row['Total'], 'Talk Time': '00:00:00'})
                else:
                    total_s = qual['Sec'].sum()
                    for _, ag in qual.iterrows():
                        final_output.append({**base, 'Agent': ag['Agent'], 'Attributed Revenue': round(row['Total']*(ag['Sec']/total_s),2), 'Talk Time': sec_to_hms(ag['Sec'])})
            else:
                qual = agent_stats[agent_stats['Sec'] >= 60]
                if qual.empty:
                    final_output.append({**base, 'Agent': 'Organic', 'Attributed Revenue': row['Total'], 'Talk Time': '00:00:00'})
                else:
                    for _, ag in qual.iterrows():
                        final_output.append({**base, 'Agent': ag['Agent'], 'Attributed Revenue': row['Total'], 'Talk Time': sec_to_hms(ag['Sec'])})

    return pd.DataFrame(final_output)

# --- UI ---
st.title("💰 TSC Revenue Attribution Portal")
st.markdown("Automating revenue from Jan onwards with specific talk-time thresholds.")

with st.sidebar:
    st.header("Upload Files")
    o_file = st.file_uploader("1. Shopify Orders Raw", type="csv")
    c_files = st.file_uploader("2. Call Data (Upload ALL files together)", type="csv", accept_multiple_files=True)

if o_file and c_files:
    try:
        o_raw = pd.read_csv(o_file)
        with st.spinner("Standardizing and Attributing Revenue..."):
            res = run_revenue_automation(o_raw, c_files)
        
        st.success(f"Processed {len(res)} attribution records.")
        st.dataframe(res, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res.to_excel(writer, sheet_name='Revenue', index=False, startrow=2)
            workbook, worksheet = writer.book, writer.sheets['Revenue']
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            for col_num, value in enumerate(res.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            worksheet.set_column(0, len(res.columns)-1, 22)
            
        st.download_button("📥 Download Final Revenue Report", data=output.getvalue(), file_name="Revenue_Attribution_Final.xlsx")
    except Exception as e: st.error(f"Error: {e}")

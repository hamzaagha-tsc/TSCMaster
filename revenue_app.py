import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | Revenue Attribution", layout="wide")

# --- UI STYLING (Forced Branding) ---
brand_navy, brand_copper = "#102a51", "#c59d5f"
st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy} !important; }}
    [data-testid="stSidebar"] * {{ color: white !important; font-family: 'Arial'; }}
    div.stButton > button:first-child {{ background-color: {brand_copper}; color: white; border-radius: 8px; font-weight: bold; }}
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    </style>
    """, unsafe_allow_html=True)

# --- UTILITIES ---
def clean_phone_master(p):
    """Converts scientific notation and floats to 10-digit strings."""
    if pd.isna(p) or str(p).strip() == "": return None
    try:
        if isinstance(p, float): s = format(p, '.0f')
        else: s = str(p).split('.')[0].strip()
        digits = re.sub(r'\D', '', s)
        return digits[-10:] if len(digits) >= 10 else None
    except: return None

def robust_time_to_sec(t):
    if pd.isna(t) or t == 0: return 0
    s = str(t).strip().replace('.', ':')
    parts = s.split(':')
    try:
        if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        if len(parts) == 2: return int(parts[0])*60 + int(parts[1])
        return 0
    except: return 0

def robust_date_parse(s):
    if pd.isna(s): return pd.NaT
    # Automatically handles YYYY-MM-DD, DD-MM-YYYY, and DD/MM/YYYY
    return pd.to_datetime(s, dayfirst=True, errors='coerce')

def sec_to_hms(s):
    h = int(s // 3600); m = int((s % 3600) // 60); s = int(s % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

# --- CORE LOGIC ---
def process_attribution(o_raw, c_raw):
    # Standardize Headers
    o_raw.columns = o_raw.columns.str.strip()
    c_raw.columns = c_raw.columns.str.strip()

    # 1. PROCESS ORDERS
    orders = o_raw.dropna(subset=['Total']).copy()
    time_col = [c for c in orders.columns if 'created' in c.lower()][0]
    orders['DT_Clean'] = orders[time_col].astype(str).str.replace(r'\s\+\d{4}$', '', regex=True)
    orders['DT_Obj'] = pd.to_datetime(orders['DT_Clean'])
    
    orders['B_Ph'] = orders['Billing Phone'].apply(clean_phone_master)
    orders['S_Ph'] = orders['Shipping Phone'].apply(clean_phone_master)
    
    o_list = []
    for _, row in orders.iterrows():
        tags = str(row.get('Tags', '')).lower()
        chan = "Store" if "retail" in tags and "store" in tags else "Website"
        phones = {p for p in [row['B_Ph'], row['S_Ph']] if p}
        for ph in phones:
            o_list.append({
                'Order ID': row['Name'], 'Value': float(row['Total']),
                'Order_Time': row['DT_Obj'], 'Order_Phone': ph, 'Channel': chan
            })
    o_df = pd.DataFrame(o_list)

    # 2. PROCESS CALLS
    def find_col(keys, df):
        for k in keys:
            for c in df.columns:
                if k.lower() in c.lower(): return c
        return None

    c_time = find_col(['start time', 'call time'], c_raw)
    c_phone = find_col(['dstphone', 'phone number', 'phone'], c_raw)
    c_talk = find_col(['user talk time', 'talk time'], c_raw)
    c_user = find_col(['user id', 'username'], c_raw)

    calls = c_raw.copy()
    calls['Call_DT'] = calls[c_time].apply(robust_date_parse)
    calls['Clean_Ph'] = calls[c_phone].apply(clean_phone_master)
    calls['Sec'] = calls[c_talk].apply(robust_time_to_sec)
    calls = calls[(calls['Sec'] >= 1) & (calls['Clean_Ph'].notna())]

    # 3. MATCHING
    results = []
    for _, order in o_df.iterrows():
        matches = calls[(calls['Clean_Ph'] == order['Order_Phone']) & (calls['Call_DT'] < order['Order_Time'])]
        
        base_row = {
            'Order ID': order['Order ID'], 'Order Value': order['Value'],
            'Order Date': order['Order_Time'].strftime('%d-%m-%Y'),
            'Order Time': order['Order_Time'].strftime('%d-%m-%Y %H:%M:%S'),
            'Order Phone': order['Order_Phone'], 'Channel': order['Channel']
        }

        if matches.empty:
            results.append({**base_row, 'Agent': 'Organic', 'Attributed Revenue': order['Value'], 'Talk Time': '00:00:00'})
            continue

        agent_sums = matches.groupby(c_user)['Sec'].sum().reset_index()

        # Rule B: >= 100k (Proportional split, 3min/180s threshold)
        if order['Value'] >= 100000:
            qual = agent_sums[agent_sums['Sec'] >= 180]
            if qual.empty:
                results.append({**base_row, 'Agent': 'Organic', 'Attributed Revenue': order['Value'], 'Talk Time': '00:00:00'})
            else:
                total_sec = qual['Sec'].sum()
                for _, ag in qual.iterrows():
                    results.append({**base_row, 'Agent': ag[c_user], 'Attributed Revenue': round(order['Value']*(ag['Sec']/total_sec), 2), 'Talk Time': sec_to_hms(ag['Sec'])})
        
        # Rule A: < 100k (Duplicate credit, 1min/60s threshold)
        else:
            qual = agent_sums[agent_sums['Sec'] >= 60]
            if qual.empty:
                results.append({**base_row, 'Agent': 'Organic', 'Attributed Revenue': order['Value'], 'Talk Time': '00:00:00'})
            else:
                for _, ag in qual.iterrows():
                    results.append({**base_row, 'Agent': ag[c_user], 'Attributed Revenue': order['Value'], 'Talk Time': sec_to_hms(ag['Sec'])})

    return pd.DataFrame(results)

# --- UI ---
st.title("💰 Revenue Attribution Hub")
with st.sidebar:
    st.header("Step 1: Data Upload")
    o_file = st.file_uploader("1. Shopify Orders Raw", type="csv")
    c_file = st.file_uploader("2. Call Data (Daily or Historical)", type="csv")

if o_file and c_file:
    try:
        o_raw, c_raw = pd.read_csv(o_file), pd.read_csv(c_file)
        with st.spinner("Analyzing Revenue..."):
            final_df = process_attribution(o_raw, c_raw)
        
        st.success(f"Analysis Complete! Found {len(final_df[final_df['Agent'] != 'Organic'])} Attributed Sales.")
        st.dataframe(final_df, use_container_width=True)

        # Excel Generation
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Revenue', index=False, startrow=2)
            workbook, worksheet = writer.book, writer.sheets['Revenue']
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            for col_num, value in enumerate(final_df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            worksheet.set_column(0, len(final_df.columns)-1, 20)
            
        st.download_button("📥 Download Revenue Report", data=output.getvalue(), file_name="Revenue_Attribution_Final.xlsx")
    except Exception as e: st.error(f"Error: {e}")

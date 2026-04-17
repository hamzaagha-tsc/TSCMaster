import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | Revenue Attribution", layout="wide")

# --- UI STYLING ---
brand_navy = "#102a51"
brand_copper = "#c59d5f"

st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy} !important; }}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p, 
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stRadio label {{
        color: white !important; font-family: 'Arial';
    }}
    div.stButton > button:first-child {{
        background-color: {brand_copper}; color: white; border-radius: 8px; font-weight: bold;
    }}
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    </style>
    """, unsafe_allow_html=True)

# --- ROBUST HELPERS ---
def clean_phone_final(p):
    if pd.isna(p) or p == '': return None
    try:
        # Handle scientific notation and floats correctly
        if isinstance(p, (float, int)):
            s = format(p, '.0f')
        else:
            s = str(p).split('.')[0].strip()
        digits = re.sub(r'\D', '', s)
        return digits[-10:] if len(digits) >= 10 else None
    except:
        return None

def hms_to_sec(t):
    if pd.isna(t) or t == '0' or t == 0: return 0
    try:
        t_str = str(t).strip().replace('.', ':') 
        parts = t_str.split(':')
        if len(parts) == 3: return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        if len(parts) == 2: return int(parts[0]) * 60 + int(parts[1])
        return 0
    except: return 0

def sec_to_hms(s):
    h, m = divmod(s, 3600); m, s = divmod(m, 60)
    return f"{int(h):02d}:{int(m):02d}:{int(s):02d}"

# --- CORE LOGIC ---
def process_revenue(orders_raw, calls_raw):
    orders_raw.columns = orders_raw.columns.str.strip()
    calls_raw.columns = calls_raw.columns.str.strip()

    # 1. CLEAN ORDERS
    # Only keep the summary row (where Total has value)
    orders = orders_raw.dropna(subset=['Total']).copy()
    
    # Strip +0530 and format time
    time_col = [c for c in orders.columns if 'created' in c.lower()][0]
    orders[time_col] = orders[time_col].astype(str).str.replace(r'\s\+\d{4}$', '', regex=True)
    orders['Order_DT_Obj'] = pd.to_datetime(orders[time_col])
    orders['Order Time Str'] = orders['Order_DT_Obj'].dt.strftime('%d-%m-%Y %H:%M:%S')
    
    orders['B_Clean'] = orders['Billing Phone'].apply(clean_phone_final)
    orders['S_Clean'] = orders['Shipping Phone'].apply(clean_phone_final)
    
    expanded_rows = []
    for _, row in orders.iterrows():
        tags = str(row.get('Tags', '')).lower()
        # Search for 'retailstore' or 'retail_store'
        channel = "Store" if "retail" in tags and "store" in tags else "Website"
        
        phones = {p for p in [row['B_Clean'], row['S_Clean']] if p}
        for ph in phones:
            expanded_rows.append({
                'Order ID': row['Name'], 'Order Value': float(row['Total']),
                'Order Date': row['Order_DT_Obj'].strftime('%d-%m-%Y'),
                'Order Time': row['Order_DT_Obj'],
                'Order Time Str': row['Order Time Str'], 
                'Order Phone': ph, 
                'Channel': channel
            })
    orders_df = pd.DataFrame(expanded_rows)

    # 2. CLEAN CALLS
    def find_col(possible, df):
        for p in possible:
            for c in df.columns:
                if p.lower() in c.lower(): return c
        return None

    call_time_col = find_col(['call time', 'start time'], calls_raw)
    phone_col = find_col(['phone number', 'phone', 'dstphone'], calls_raw)
    talk_col = find_col(['user talk time', 'talk time'], calls_raw)
    user_col = find_col(['user id', 'username'], calls_raw)

    calls = calls_raw.copy()
    calls['Call_DT_Obj'] = pd.to_datetime(calls[call_time_col], dayfirst=True, errors='coerce')
    calls['Clean_Phone'] = calls[phone_col].apply(clean_phone_final)
    calls['Talk_Sec'] = calls[talk_col].apply(hms_to_sec)
    
    # Filter for connected calls
    calls = calls[(calls['Talk_Sec'] >= 1) & (calls['Clean_Phone'].notna())].copy()

    # 3. ATTRIBUTION
    final_data = []
    for _, order in orders_df.iterrows():
        match_calls = calls[(calls['Clean_Phone'] == order['Order Phone']) & 
                           (calls['Call_DT_Obj'] < order['Order Time'])]
        
        if match_calls.empty:
            final_data.append({**order, 'Agent': 'Organic', 'Attributed Revenue': order['Order Value'], 'Total Talk Time': '00:00:00'})
            continue

        agent_stats = match_calls.groupby(user_col)['Talk_Sec'].sum().reset_index()
        
        if order['Order Value'] >= 100000:
            qualified = agent_stats[agent_stats['Talk_Sec'] >= 180].copy()
            if qualified.empty:
                final_data.append({**order, 'Agent': 'Organic', 'Attributed Revenue': order['Order Value'], 'Total Talk Time': '00:00:00'})
            else:
                total_q = qualified['Talk_Sec'].sum()
                for _, ag in qualified.iterrows():
                    share = ag['Talk_Sec'] / total_q
                    final_data.append({**order, 'Agent': ag[user_col], 'Attributed Revenue': round(order['Order Value'] * share, 2), 'Total Talk Time': sec_to_hms(ag['Talk_Sec'])})
        else:
            qualified = agent_stats[agent_stats['Talk_Sec'] >= 60].copy()
            if qualified.empty:
                final_data.append({**order, 'Agent': 'Organic', 'Attributed Revenue': order['Order Value'], 'Total Talk Time': '00:00:00'})
            else:
                for _, ag in qualified.iterrows():
                    final_data.append({**order, 'Agent': ag[user_col], 'Attributed Revenue': order['Order Value'], 'Total Talk Time': sec_to_hms(ag['Talk_Sec'])})

    return pd.DataFrame(final_data)

# --- UI ---
st.title("💰 Revenue Attribution Automation")
with st.sidebar:
    st.header("Upload Files")
    o_file = st.file_uploader("1. Shopify Orders Raw", type="csv")
    c_file = st.file_uploader("2. Call Attribution Data", type="csv")

if o_file and c_file:
    try:
        o_raw, c_raw = pd.read_csv(o_file), pd.read_csv(c_file)
        result_df = process_revenue(o_raw, c_raw)
        
        display_df = result_df.drop(columns=['Order Time']).rename(columns={'Order Time Str': 'Order Time'})
        st.success("Attribution Calculated!")
        st.dataframe(display_df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, sheet_name='Revenue', index=False, startrow=2)
            workbook, worksheet = writer.book, writer.sheets['Revenue']
            fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1})
            for col_num, value in enumerate(display_df.columns.values):
                worksheet.write(2, col_num, value, fmt)
            worksheet.set_column(0, len(display_df.columns)-1, 22)
            
        st.download_button("📥 Download Revenue Excel", data=output.getvalue(), file_name="TSC_Revenue_Attribution.xlsx")
    except Exception as e: st.error(f"Error: {e}")

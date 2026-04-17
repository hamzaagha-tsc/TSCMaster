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

# --- HELPERS ---
def clean_phone(p):
    if pd.isna(p): return None
    digits = re.sub(r'\D', '', str(p))
    return digits[-10:] if len(digits) >= 10 else None

def hms_to_sec(t):
    if pd.isna(t) or t == '0' or t == 0: return 0
    try:
        t_str = str(t).strip()
        # Handle formats like 00:02:30 or 00:02.5
        t_str = t_str.replace('.', ':') 
        parts = t_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        return 0
    except: return 0

def sec_to_hms(s):
    h, m = divmod(s, 3600); m, s = divmod(m, 60)
    return f"{int(h):02d}:{int(m):02d}:{int(s):02d}"

# --- CORE LOGIC ---
def process_revenue(orders_raw, calls_raw):
    # Standardize Headers
    orders_raw.columns = orders_raw.columns.str.strip()
    calls_raw.columns = calls_raw.columns.str.strip()

    # 1. CLEAN ORDERS
    orders = orders_raw.dropna(subset=['Total']).copy()
    # Find the right time column (Created at)
    time_col = [c for c in orders.columns if 'created' in c.lower()][0]
    
    orders[time_col] = orders[time_col].str.replace(r'\s\+\d{4}$', '', regex=True)
    orders['Order_DT_Obj'] = pd.to_datetime(orders[time_col])
    orders['Order Date'] = orders['Order_DT_Obj'].dt.strftime('%d-%m-%Y')
    orders['Order Time Str'] = orders['Order_DT_Obj'].dt.strftime('%d-%m-%Y %H:%M:%S')
    
    # Phone Logic
    orders['B_Phone'] = orders['Billing Phone'].apply(clean_phone) if 'Billing Phone' in orders.columns else None
    orders['S_Phone'] = orders['Shipping Phone'].apply(clean_phone) if 'Shipping Phone' in orders.columns else None
    
    expanded_rows = []
    for _, row in orders.iterrows():
        channel = "Store" if "retail_store" in str(row.get('Tags', '')).lower() else "Website"
        phones = {p for p in [row['B_Phone'], row['S_Phone']] if p}
        for ph in phones:
            expanded_rows.append({
                'Order ID': row['Name'], 'Order Value': float(row['Total']),
                'Order Date': row['Order Date'], 'Order Time': row['Order_DT_Obj'],
                'Order Time Str': row['Order Time Str'], 'Order Phone': ph, 'Channel': channel
            })
    orders_df = pd.DataFrame(expanded_rows)

    # 2. CLEAN CALLS
    calls = calls_raw.copy()
    # Find the right columns even if names vary slightly
    call_time_col = [c for c in calls.columns if 'call time' in c.lower()][0]
    phone_col = [c for c in calls.columns if 'phone' in c.lower()][0]
    talk_col = [c for c in calls.columns if 'talk time' in c.lower()][0]
    user_col = [c for c in calls.columns if 'user id' in c.lower() or 'username' in c.lower()][0]

    calls['Call_DT_Obj'] = pd.to_datetime(calls[call_time_col], dayfirst=True, errors='coerce')
    calls['Clean_Phone'] = calls[phone_col].apply(clean_phone)
    calls['Talk_Sec'] = calls[talk_col].apply(hms_to_sec)
    
    # Filter for connected calls with valid phone
    calls = calls[(calls['Talk_Sec'] >= 1) & (calls['Clean_Phone'].notna())].copy()

    # 3. ATTRIBUTION
    final_data = []
    for _, order in orders_df.iterrows():
        # Match calls for this phone that happened BEFORE order
        match_calls = calls[(calls['Clean_Phone'] == order['Order Phone']) & 
                           (calls['Call_DT_Obj'] < order['Order Time'])]
        
        if match_calls.empty:
            final_data.append({**order, 'Agent': 'Organic', 'Attributed Revenue': order['Order Value'], 'Total Talk Time': '00:00:00'})
            continue

        # Aggregate Talk Time per Agent
        agent_stats = match_calls.groupby(user_col)['Talk_Sec'].sum().reset_index()
        
        # Scenario B: >= 100,000 (Split logic, 3min/180s criteria)
        if order['Order Value'] >= 100000:
            qualified = agent_stats[agent_stats['Talk_Sec'] >= 180].copy()
            if qualified.empty:
                final_data.append({**order, 'Agent': 'Organic', 'Attributed Revenue': order['Order Value'], 'Total Talk Time': '00:00:00'})
            else:
                total_q_sec = qualified['Talk_Sec'].sum()
                for _, ag in qualified.iterrows():
                    share = ag['Talk_Sec'] / total_q_sec
                    final_data.append({**order, 'Agent': ag[user_col], 'Attributed Revenue': round(order['Order Value'] * share, 2), 'Total Talk Time': sec_to_hms(ag['Talk_Sec'])})
        
        # Scenario A: < 100,000 (Duplicate logic, 1min/60s criteria)
        else:
            qualified = agent_stats[agent_stats['Talk_Sec'] >= 60].copy()
            if qualified.empty:
                final_data.append({**order, 'Agent': 'Organic', 'Attributed Revenue': order['Order Value'], 'Total Talk Time': '00:00:00'})
            else:
                for _, ag in qualified.iterrows():
                    final_data.append({**order, 'Agent': ag[user_col], 'Attributed Revenue': order['Order Value'], 'Total Talk Time': sec_to_hms(ag['Talk_Sec'])})

    return pd.DataFrame(final_data)

# --- APP UI ---
st.title("💰 Revenue Attribution Automation")

with st.sidebar:
    st.header("Upload Files")
    order_file = st.file_uploader("1. Shopify Orders Raw", type="csv")
    call_file = st.file_uploader("2. Call Attribution Data", type="csv")

if order_file and call_file:
    try:
        o_raw = pd.read_csv(order_file)
        c_raw = pd.read_csv(call_file)
        
        with st.spinner("Calculating Attribution..."):
            result_df = process_revenue(o_raw, c_raw)
            # Cleanup for display
            display_df = result_df.drop(columns=['Order Time']).rename(columns={'Order Time Str': 'Order Time'})
            
        st.success("Analysis Complete!")
        st.dataframe(display_df, use_container_width=True)

        # Excel Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, sheet_name='Attribution', index=False, startrow=2)
            workbook = writer.book
            worksheet = writer.sheets['Attribution']
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1})
            data_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 11, 'border': 1})
            
            for col_num, value in enumerate(display_df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            worksheet.set_column(0, len(display_df.columns)-1, 22)
            
        st.download_button("📥 Download Final Report", data=output.getvalue(), file_name="Revenue_Attribution.xlsx")
    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Ensure your files have columns for 'Call Time', 'Phone', 'Talk Time', and 'User ID'.")

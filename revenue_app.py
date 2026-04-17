import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | Revenue Attribution", layout="wide")

# --- UI STYLING ---
brand_navy, brand_copper = "#102a51", "#c59d5f"
st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy} !important; }}
    [data-testid="stSidebar"] * {{ color: white !important; }}
    div.stButton > button:first-child {{ background-color: {brand_copper}; color: white; border-radius: 8px; font-weight: bold; }}
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    .debug-box {{ padding: 10px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9; font-family: monospace; }}
    </style>
    """, unsafe_allow_html=True)

# --- ROBUST UTILITIES ---
def force_10_digit_phone(p):
    """Converts any phone format (float, scientific, string) to last 10 digits."""
    if pd.isna(p) or str(p).strip() == "": return None
    # 1. Remove .0 if it's a float-string and scientific notation junk
    s = str(p).split('.')[0].split('e')[0].strip()
    # 2. Keep only digits
    digits = re.sub(r'\D', '', s)
    # 3. Take last 10
    return digits[-10:] if len(digits) >= 10 else None

def smart_time_to_sec(t):
    """Handles HH:MM:SS, MM:SS, and MM.SS formats."""
    if pd.isna(t) or t == 0: return 0
    s = str(t).strip().replace('.', ':')
    parts = s.split(':')
    try:
        if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        if len(parts) == 2: return int(parts[0])*60 + int(parts[1])
        return 0
    except: return 0

def sec_to_hms(s):
    h, m = divmod(int(s), 3600); m, s = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

# --- THE ENGINE ---
def run_attribution(o_raw, c_raw):
    # Standardize columns
    o_raw.columns = o_raw.columns.str.strip()
    c_raw.columns = c_raw.columns.str.strip()

    # 1. PROCESS ORDERS
    orders = o_raw.dropna(subset=['Total']).copy()
    # Handle Shopify '+0530' and parse
    time_col = [c for c in orders.columns if 'created' in c.lower()][0]
    orders['Clean_Time_Str'] = orders[time_col].astype(str).str.replace(r'\s\+\d{4}$', '', regex=True)
    orders['DT_Obj'] = pd.to_datetime(orders['Clean_Time_Str'])
    
    # Phone Sanitization
    orders['B_Clean'] = orders['Billing Phone'].apply(force_10_digit_phone)
    orders['S_Clean'] = orders['Shipping Phone'].apply(force_10_digit_phone)
    
    orders_processed = []
    for _, row in orders.iterrows():
        chan = "Store" if "retailstore" in str(row.get('Tags', '')).replace('_','').lower() else "Website"
        phones = {p for p in [row['B_Clean'], row['S_Clean']] if p}
        for ph in phones:
            orders_processed.append({
                'Order ID': row['Name'], 'Value': float(row['Total']),
                'Order_Time': row['DT_Obj'], 'Order_Phone': ph, 'Channel': chan
            })
    o_df = pd.DataFrame(orders_processed)

    # 2. PROCESS CALLS
    calls = c_raw.copy()
    # Map headers dynamically
    c_time_col = next((c for c in calls.columns if 'call time' in c.lower()), None)
    c_phone_col = next((c for c in calls.columns if 'phone' in c.lower()), None)
    c_talk_col = next((c for c in calls.columns if 'talk time' in c.lower()), None)
    c_user_col = next((c for c in calls.columns if 'user id' in c.lower() or 'username' in c.lower()), None)

    # Date parsing (Very important: checking for Day-First)
    calls['Call_DT'] = pd.to_datetime(calls[c_time_col], dayfirst=True, errors='coerce')
    calls['Clean_Phone'] = calls[c_phone_col].apply(force_10_digit_phone)
    calls['Talk_Sec'] = calls[c_talk_col].apply(smart_time_to_sec)
    
    # Keep only connected calls with a valid phone
    calls = calls[(calls['Talk_Sec'] >= 1) & (calls['Clean_Phone'].notna())].copy()

    # 3. DEBUG STATS (To show in UI)
    st.write("### 🔍 Debug Console")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Orders Samples:**")
        st.write(o_df[['Order ID', 'Order_Phone', 'Order_Time']].head(3))
    with col2:
        st.write("**Calls Samples:**")
        st.write(calls[['Clean_Phone', 'Call_DT', 'Talk_Sec']].head(3))

    # 4. MATCHING
    final_output = []
    for _, order in o_df.iterrows():
        # FIND MATCHES: Same phone AND Call occurred BEFORE Order
        match = calls[(calls['Clean_Phone'] == order['Order_Phone']) & (calls['Call_DT'] < order['Order_Time'])]
        
        if match.empty:
            final_output.append({**order, 'Agent': 'Organic', 'Revenue': order['Value'], 'Talk_Time': '00:00:00'})
            continue

        # Group by agent for the window
        agent_sums = match.groupby(c_user_col)['Talk_Sec'].sum().reset_index()

        # RULE B: >= 100k (Proportionate Split, 3min threshold)
        if order['Value'] >= 100000:
            qualified = agent_sums[agent_sums['Talk_Sec'] >= 180]
            if qualified.empty:
                final_output.append({**order, 'Agent': 'Organic', 'Revenue': order['Value'], 'Talk_Time': '00:00:00'})
            else:
                total_q_sec = qualified['Talk_Sec'].sum()
                for _, ag in qualified.iterrows():
                    share = ag['Talk_Sec'] / total_q_sec
                    final_output.append({**order, 'Agent': ag[c_user_col], 'Revenue': round(order['Value']*share, 2), 'Talk_Time': sec_to_hms(ag['Talk_Sec'])})
        
        # RULE A: < 100k (Multi-Credit, 1min threshold)
        else:
            qualified = agent_sums[agent_sums['Talk_Sec'] >= 60]
            if qualified.empty:
                final_output.append({**order, 'Agent': 'Organic', 'Revenue': order['Value'], 'Talk_Time': '00:00:00'})
            else:
                for _, ag in qualified.iterrows():
                    final_output.append({**order, 'Agent': ag[c_user_col], 'Revenue': order['Value'], 'Talk_Time': sec_to_hms(ag['Talk_Sec'])})

    return pd.DataFrame(final_output)

# --- APP UI ---
st.title("💰 Standalone Revenue Attribution")

with st.sidebar:
    st.header("Files")
    o_file = st.file_uploader("Upload Shopify Raw CSV", type="csv")
    c_file = st.file_uploader("Upload Call Data CSV", type="csv")

if o_file and c_file:
    try:
        o_raw, c_raw = pd.read_csv(o_file), pd.read_csv(c_file)
        result = run_attribution(o_raw, c_raw)
        
        # Format for display
        final_df = result.copy()
        final_df['Order_Time'] = final_df['Order_Time'].dt.strftime('%d-%m-%Y %H:%M:%S')
        
        st.success(f"Processed {len(result)} records.")
        st.dataframe(final_df, use_container_width=True)

        # Excel Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Attribution', index=False, startrow=2)
            workbook = writer.book
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1})
            worksheet = writer.sheets['Attribution']
            for col_num, value in enumerate(final_df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
        
        st.download_button("📥 Download Attribution Report", data=output.getvalue(), file_name="Revenue_Report.xlsx")
    except Exception as e:
        st.error(f"Error during processing: {e}")

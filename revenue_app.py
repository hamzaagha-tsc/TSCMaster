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
    user_col = [c for c in calls.columns if 'user id' in c.lower

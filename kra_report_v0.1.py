import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | KRA Portal", layout="wide")

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

# --- HELPERS ---
def hms_to_sec(t):
    if pd.isna(t) or t == '0' or t == 0: return 0
    try:
        parts = str(t).strip().split(':')
        if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        if len(parts) == 2: return int(parts[0])*60 + int(parts[1])
        return 0
    except: return 0

def sec_to_hms(s):
    h = int(s // 3600); m = int((s % 3600) // 60); s = int(s % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

# --- CORE LOGIC ---
def generate_kra(df):
    # 1. CLEAN HEADERS & DATA
    df.columns = df.columns.str.strip()
    
    # Identify required columns
    time_col = next(c for c in df.columns if 'start time' in c.lower())
    user_col = next(c for c in df.columns if 'user name' in c.lower())
    phone_col = next(c for c in df.columns if 'phone' in c.lower())
    talk_col = next(c for c in df.columns if 'talk time' in c.lower())
    
    # Extract Latest Timestamp for filename
    latest_time_dt = pd.to_datetime(df[time_col], dayfirst=True).max()
    ts_str = latest_time_dt.strftime('%d-%m-%Y_%H%M%S')
    
    # Pre-process Columns
    df['Date'] = pd.to_datetime(df[time_col], dayfirst=True).dt.date
    df['Agent Name'] = df[user_col].str.replace('AA ', '', regex=False)
    df['Secs'] = df[talk_col].apply(hms_to_sec)
    
    # 2. AGGREGATION (Per Agent Per Day)
    # Unique Outbound: Total unique phones dialed
    unq_ob = df.groupby(['Date', 'Agent Name'])[phone_col].nunique().reset_index(name='Unique Outbound Calls')
    
    # Unique Connected: Unique phones where talk time > 0
    unq_cc = df[df['Secs'] > 0].groupby(['Date', 'Agent Name'])[phone_col].nunique().reset_index(name='Unique Connected Calls')
    
    # Total Talk Time
    talk_sum = df.groupby(['Date', 'Agent Name'])['Secs'].sum().reset_index(name='Talk_Secs')
    
    # 3. MERGE & SORT
    kra = pd.merge(unq_ob, unq_cc, on=['Date', 'Agent Name'], how='left').fillna(0)
    kra = pd.merge(kra, talk_sum, on=['Date', 'Agent Name'], how='left').fillna(0)
    
    # Final cleanup
    kra['Total Talk Time'] = kra['Talk_Secs'].apply(sec_to_hms)
    kra = kra.sort_values(by='Unique Connected Calls', ascending=False)
    
    # Select final columns in order
    final_kra = kra[['Agent Name', 'Unique Outbound Calls', 'Unique Connected Calls', 'Total Talk Time', 'Date', 'Talk_Secs']]
    return final_kra, ts_str

# --- UI ---
st.title("📊 KRA Report Generator")
st.markdown("Simple 1-step tool to generate Daily KRA Reports from Custom Sales Reports.")

with st.sidebar:
    st.header("Step 1")
    uploaded_file = st.file_uploader("Upload Custom Sales Report (CSV)", type="csv")

if uploaded_file:
    try:
        raw_df = pd.read_csv(uploaded_file)
        kra_data, timestamp = generate_kra(raw_df)
        
        # Display Preview (Removing helper columns for UI)
        st.subheader(f"KRA Data Preview (Sorted by Connected Calls)")
        st.dataframe(kra_data.drop(columns=['Talk_Secs']), use_container_width=True)

        # EXCEL GENERATION WITH RAG
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # We don't include Talk_Secs in the excel sheet, but we use it for formatting
            excel_df = kra_data.drop(columns=['Talk_Secs'])
            excel_df.to_excel(writer, sheet_name='KRA Report', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['KRA Report']
            
            # FORMATS
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            data_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            
            # Apply Headers
            for col_num, value in enumerate(excel_df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            # Apply Data Style & Width
            worksheet.set_column('A:E', 25, data_fmt)
            
            # --- RAG COLOR GRADING (3-Color Scale) ---
            # Range: Red (Low) -> Yellow (Mid) -> Green (High)
            color_scale = {'type': '3_color_scale', 
                           'min_color': "#FF0000", 'mid_color': "#FFFF00", 'max_color': "#00B050"}
            
            last_row = len(excel_df) + 2
            # B: Unique Outbound
            worksheet.conditional_format(3, 1, last_row, 1, color_scale)
            # C: Unique Connected
            worksheet.conditional_format(3, 2, last_row, 2, color_scale)
            # D: Total Talk Time (Based on seconds - we apply it to the column)
            # Since formatting strings is hard in xlsxwriter based on other columns, 
            # we apply the scale to the hidden Talk_Secs and formatting to the visible column manually 
            # or just scale the visible calls. For "Talk Time", we scale the values.
            worksheet.conditional_format(3, 3, last_row, 3, color_scale)
            
        st.download_button(
            label="📥 Download KRA Report",
            data=output.getvalue(),
            file_name=f"KRA_Report_Till_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error: {e}")

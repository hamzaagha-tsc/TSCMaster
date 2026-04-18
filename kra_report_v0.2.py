import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | KRA Portal", layout="wide")

# --- UI BRANDING (Adaptive & High Contrast) ---
brand_navy, brand_copper = "#102a51", "#c59d5f"
st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy} !important; }}
    [data-testid="stSidebar"] * {{ color: white !important; font-family: 'Arial'; }}
    div.stButton > button:first-child {{
        background-color: {brand_copper}; color: white; border-radius: 8px; font-weight: bold; border: none;
    }}
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    </style>
    """, unsafe_allow_html=True)

# --- ROBUST UTILITIES ---
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

# --- CORE KRA ENGINE ---
def process_kra_logic(df):
    df.columns = df.columns.str.strip()
    
    # 1. Identify Columns
    # We use exact names from your 16.04.csv for maximum reliability
    time_col = 'Start Time' if 'Start Time' in df.columns else next(c for c in df.columns if 'time' in c.lower())
    user_col = 'User Name' if 'User Name' in df.columns else next(c for c in df.columns if 'user' in c.lower())
    phone_col = 'dstPhone' if 'dstPhone' in df.columns else next(c for c in df.columns if 'phone' in c.lower())
    talk_col = 'Talk Time' if 'Talk Time' in df.columns else next(c for c in df.columns if 'talk' in c.lower())

    # 2. Extract Date into Separate Column (CRITICAL FIX)
    df['dt_parsed'] = pd.to_datetime(df[time_col], dayfirst=True)
    df['Report Date'] = df['dt_parsed'].dt.date
    
    # Get filename timestamp from last call
    last_call_ts = df['dt_parsed'].max().strftime('%d-%m-%Y_%H%M%S')
    
    # 3. Clean Agent Names (Remove "AA " prefix)
    df['Clean Agent'] = df[user_col].astype(str).str.replace('AA ', '', regex=False).str.strip()
    
    # 4. Clean Talk Time to Seconds
    df['Talk_Secs'] = df[talk_col].apply(hms_to_sec)
    
    # 5. AGGREGATION (Grouping by Date and Agent)
    # Unique Outbound Calls
    ob_agg = df.groupby(['Report Date', 'Clean Agent'])[phone_col].nunique().reset_index(name='Unique Outbound Calls')
    
    # Unique Connected Calls (where Talk Time > 0)
    cc_agg = df[df['Talk_Secs'] > 0].groupby(['Report Date', 'Clean Agent'])[phone_col].nunique().reset_index(name='Unique Connected Calls')
    
    # Total Talk Time
    tt_agg = df.groupby(['Report Date', 'Clean Agent'])['Talk_Secs'].sum().reset_index(name='Total_Talk_Secs')
    
    # 6. MERGE & SORT
    final = pd.merge(ob_agg, cc_agg, on=['Report Date', 'Clean Agent'], how='left').fillna(0)
    final = pd.merge(final, tt_agg, on=['Report Date', 'Clean Agent'], how='left').fillna(0)
    
    # Format Time back to HMS
    final['Total Talk Time'] = final['Total_Talk_Secs'].apply(sec_to_hms)
    
    # Final Column Order and Sorting (Connected Calls Z->A)
    final = final.sort_values(by=['Report Date', 'Unique Connected Calls'], ascending=[True, False])
    
    report = final.rename(columns={'Clean Agent': 'Agent Name', 'Report Date': 'Date'})
    return report[['Agent Name', 'Unique Outbound Calls', 'Unique Connected Calls', 'Total Talk Time', 'Date', 'Total_Talk_Secs']], last_call_ts

# --- UI SECTION ---
st.title("KRA Reports - Test")
with st.sidebar:
    st.markdown("TSC Portal")
    st.divider()
    up_file = st.file_uploader("Upload Custom Sales Report (CSV)", type="csv")

if up_file:
    try:
        raw_df = pd.read_csv(up_file)
        kra_out, ts_name = process_kra_logic(raw_df)
        
        st.success(f"Report Generated!")
        # Display Preview
        st.dataframe(kra_out.drop(columns=['Total_Talk_Secs', 'Date']), use_container_width=True)

        # Excel Export with RAG Color Grading
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df = kra_out.drop(columns=['Total_Talk_Secs'])
            export_df.to_excel(writer, sheet_name='KRA Report', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['KRA Report']
            
            # Format styles
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            data_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            
            # Headers
            for col_num, value in enumerate(export_df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            worksheet.set_column('A:E', 25, data_fmt)
            
            # RAG (Red-Amber-Green) Heatmap logic
            rag_logic = {'type': '3_color_scale', 'min_color': "#FF0000", 'mid_color': "#FFFF00", 'max_color': "#00B050"}
            
            last_row = len(export_df) + 2
            # Apply RAG to Unique Outbound, Unique Connected, and Total Talk Time
            worksheet.conditional_format(3, 1, last_row, 1, rag_logic) # Outbound
            worksheet.conditional_format(3, 2, last_row, 2, rag_logic) # Connected
            worksheet.conditional_format(3, 3, last_row, 3, rag_logic) # Talk Time
            
        st.download_button(
            label="📥 Download KRA Excel",
            data=output.getvalue(),
            file_name=f"KRA_Report_Till_{ts_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error: {e}")

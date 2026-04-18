import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="TSC | KRA Portal", layout="wide")

# --- UI PROTECTION (Theme Adaptive & No External Icons) ---
brand_navy, brand_copper = "#102a51", "#c59d5f"
st.markdown(f"""
    <style>
    /* Force Sidebar Visibility in both Light & Dark Mode */
    [data-testid="stSidebar"] {{
        background-color: {brand_navy} !important;
    }}
    [data-testid="stSidebar"] * {{
        color: white !important;
    }}
    /* Main Area Styling */
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    
    /* Button Styling */
    div.stButton > button:first-child {{
        background-color: {brand_copper};
        color: white;
        border-radius: 8px;
        font-weight: bold;
        border: none;
    }}
    </style>
    """, unsafe_allow_html=True)

# --- UTILITIES ---
def clean_phone_master(p):
    """Handles scientific notation and floats to return 10-digit strings."""
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

# --- CORE KRA LOGIC ---
def process_kra_report(df):
    df.columns = df.columns.str.strip()
    
    # 1. Dynamic Column Mapping
    time_c = next(c for c in df.columns if 'time' in c.lower() and 'hangup' not in c.lower())
    user_c = next(c for c in df.columns if 'user name' in c.lower())
    phone_c = next(c for c in df.columns if 'phone' in c.lower() or 'dst' in c.lower())
    talk_c = next(c for c in df.columns if 'talk time' in c.lower())
    
    # 2. Date and Time Sanitization
    # Force DayFirst=True and normalize to remove all time components for grouping
    df['dt_obj'] = pd.to_datetime(df[time_c], dayfirst=True, errors='coerce')
    df['Clean_Date'] = df['dt_obj'].dt.normalize() 
    
    # Generate timestamp for filename from the last actual call
    filename_ts = df['dt_obj'].max().strftime('%d-%m-%Y_%H%M%S')
    
    # 3. Data Cleaning
    df['Clean_Agent'] = df[user_c].astype(str).str.replace('AA ', '', regex=False).str.strip()
    df['Clean_Phone'] = df[phone_c].apply(clean_phone_master)
    df['Secs'] = df[talk_c].apply(hms_to_sec)
    
    # 4. Grouped Aggregation
    # B: Unique Outbound (Total unique phones dialed per agent per day)
    ob = df.groupby(['Clean_Agent', 'Clean_Date'])['Clean_Phone'].nunique().reset_index(name='Unique Outbound Calls')
    
    # C: Unique Connected (Unique phones where Talk Time >= 1 sec)
    cc = df[df['Secs'] > 0].groupby(['Clean_Agent', 'Clean_Date'])['Clean_Phone'].nunique().reset_index(name='Unique Connected Calls')
    
    # D: Total Talk Time
    tt = df.groupby(['Clean_Agent', 'Clean_Date'])['Secs'].sum().reset_index(name='Talk_Secs')
    
    # 5. Merge, Format and Sort
    final = pd.merge(ob, cc, on=['Clean_Agent', 'Clean_Date'], how='left').fillna(0)
    final = pd.merge(final, tt, on=['Clean_Agent', 'Clean_Date'], how='left').fillna(0)
    
    final['Total Talk Time'] = final['Talk_Secs'].apply(sec_to_hms)
    
    # Sort by Connected Calls (Z->A)
    final = final.sort_values(by=['Clean_Date', 'Unique Connected Calls'], ascending=[True, False])
    
    # Map back to requested columns
    report = final.rename(columns={'Clean_Agent': 'Agent Name', 'Clean_Date': 'Date'})
    return report[['Agent Name', 'Unique Outbound Calls', 'Unique Connected Calls', 'Total Talk Time', 'Date', 'Talk_Secs']], filename_ts

# --- UI SECTION ---
st.title("📊 KRA Performance Report")
st.markdown("Upload the **Custom Sales Report** to generate the graded performance file.")

with st.sidebar:
    st.markdown("### 🏢 TSC Administration")
    st.divider()
    uploaded_file = st.file_uploader("Upload CSV File", type="csv")
    st.info("The report will auto-clean 'AA ' prefixes and apply RAG color grading.")

if uploaded_file:
    try:
        raw_df = pd.read_csv(uploaded_file)
        kra_res, timestamp = process_kra_report(raw_df)
        
        st.success(f"Report Generated! Timestamp: {timestamp}")
        # Preview for User (Hiding helper column)
        st.dataframe(kra_res.drop(columns=['Talk_Secs']), use_container_width=True)

        # Excel Export with RAG
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df = kra_res.drop(columns=['Talk_Secs'])
            export_df.to_excel(writer, sheet_name='KRA Report', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['KRA Report']
            
            # Format Definitions
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            data_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
            
            # Write Headers
            for col_num, value in enumerate(export_df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            worksheet.set_column('A:E', 22, data_fmt)
            
            # RAG (Red-Amber-Green) Conditional Formatting
            rag_colors = {
                'type': '3_color_scale',
                'min_color': "#FF0000", # Red
                'mid_color': "#FFFF00", # Yellow
                'max_color': "#00B050"  # Green
            }
            
            last_row = len(export_df) + 2
            # Apply RAG to Columns B, C, and D
            worksheet.conditional_format(3, 1, last_row, 1, rag_colors) # Outbound
            worksheet.conditional_format(3, 2, last_row, 2, rag_colors) # Connected
            
            # To apply RAG to Talk Time (Column D), we scale based on the numeric seconds
            # Excel formats the background based on the relative values in the column
            worksheet.conditional_format(3, 3, last_row, 3, rag_colors) 
            
        st.download_button(
            label="📥 Download KRA Excel",
            data=output.getvalue(),
            file_name=f"KRA_Report_Till_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error Processing Data: {e}")

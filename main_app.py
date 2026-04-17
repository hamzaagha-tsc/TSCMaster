import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- APP CONFIGURATION ---
st.set_page_config(
    page_title="TSC | Performance Portal", 
    page_icon="https://thesleepcompany.in/cdn/shop/files/fav-icon_32x32.png", 
    layout="wide"
)

# --- THE SLEEP COMPANY BRANDING ---
brand_navy = "#102a51"
brand_copper = "#c59d5f"

st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy}; }}
    [data-testid="stSidebar"] * {{ color: white; }}
    div.stButton > button:first-child {{
        background-color: {brand_copper}; color: white; border-radius: 5px; border: none; font-weight: bold;
    }}
    h1, h2, h3 {{ color: {brand_navy}; font-family: 'Arial'; }}
    </style>
    """, unsafe_allow_html=True)

# --- UNIVERSAL HELPERS ---
def hms_to_sec(t):
    if pd.isna(t) or t == '0' or t == 0: return 0
    try:
        t_clean = str(t).split('.')[0].strip() 
        parts = t_clean.split(':')
        if len(parts) == 3: return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        return 0
    except: return 0

def sec_to_hms(seconds):
    if pd.isna(seconds) or seconds <= 0: return "00:00:00"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def to_excel_formatted_multi(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # STYLES: Arial 11, Yellow Background
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
        footer_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
        data_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'})
        title_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14, 'font_color': brand_navy})

        for sheet_name, content in sheets_dict.items():
            df = content['df']
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
            worksheet = writer.sheets[sheet_name]
            worksheet.write(0, 0, content['title'], title_fmt)
            worksheet.set_column(0, len(df.columns) - 1, 20)

            # Apply Headers
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)

            # Apply Data & Yellow Footer Logic
            last_row_idx = len(df) + 2
            for r_idx in range(len(df)):
                curr_row = r_idx + 3
                is_sum = content.get('has_summary', False) and (curr_row == last_row_idx)
                style = footer_fmt if is_sum else data_fmt
                for c_idx in range(len(df.columns)):
                    worksheet.write(curr_row, c_idx, df.iloc[r_idx, c_idx], style)
    return output.getvalue()

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.image("https://thesleepcompany.in/cdn/shop/files/Logo_White_300x.png?v=1614330134", width=200)
    st.header("Department Portal")
    app_mode = st.radio("Choose App", ["Sales Performance", "Pre-Sales SLA & Breaks"])
    st.divider()

# ==========================================
# 1. SALES TEAM SECTION
# ==========================================
if app_mode == "Sales Performance":
    st.title("🚀 Sales Team Performance")
    with st.sidebar:
        p_file = st.file_uploader("1. Productivity Summary", type="csv")
        s_file = st.file_uploader("2. Session Details", type="csv")
        c_file = st.file_uploader("3. Custom Sales Report", type="csv")

    if p_file and s_file and c_file:
        try:
            prod, sess, sales = pd.read_csv(p_file), pd.read_csv(s_file

import streamlit as st
import pandas as pd
import io

# --- APP CONFIGURATION ---
st.set_page_config(
    page_title="TSC | Performance Portal", 
    page_icon="https://thesleepcompany.in/cdn/shop/files/new_logo.webp?v=1706780127&width=600", 
    layout="wide"
)

# --- THE SLEEP COMPANY BRANDING & UI ---
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
    .stDataFrame {{ font-family: 'Arial'; }}
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

def to_excel_formatted(df, sheet_name, report_title, has_summary_row=True):
    """Generates an Excel file with Arial 11 and Yellow headers/footers."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Formats
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'
        })
        footer_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1
        })
        data_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 11, 'border': 1})
        title_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14, 'font_color': brand_navy})

        # Apply Title
        worksheet.write(0, 0, report_title, title_fmt)

        # Apply Header Format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(2, col_num, value, header_fmt)

        # Apply Data & Footer Format
        last_row = len(df) + 2
        for r in range(3, last_row + 1):
            fmt = footer_fmt if (has_summary_row and r == last_row) else data_fmt
            for c in range(len(df.columns)):
                val = df.iloc[r-3, c]
                worksheet.write(r, c, val, fmt)

        worksheet.set_column(0, len(df.columns)-1, 18)
    return output.getvalue()

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.image("https://thesleepcompany.in/cdn/shop/files/Logo_White_300x.png?v=1614330134", width=200)
    st.header("Sales Department Portal")
    app_mode = st.radio("Choose App", ["Sales Performance", "Pre-Sales SLA"])
    st.divider()

# ==========================================
# SALES TEAM SECTION
# ==========================================
if app_mode == "Sales Performance":
    st.title("🚀 Sales Team Automation")
    with st.sidebar:
        p_file = st.file_uploader("1. Productivity Summary", type="csv")
        s_file = st.file_uploader("2. Session Details", type="csv")
        c_file = st.file_uploader("3. Custom Sales Report", type="csv")

    if p_file and s_file and c_file:
        prod, sess, sales = pd.read_csv(p_file), pd.read_csv(s_file), pd.read_csv(c_file)
        # Cleaning & Logic (Simplified for clarity)
        for d in [prod, sess, sales]: d.columns = d.columns.str.strip()
        prod['Date'] = pd.to_datetime(prod['Interval Start'], dayfirst=True, errors='coerce').dt.date
        sess['Date'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.date
        sales['Date'] = pd.to_datetime(sales['Start Time'], dayfirst=True, errors='coerce').dt.date
        
        # Break & Sales Logic
        sess['BS'] = sess['Break Duration'].apply(hms_to_sec)
        bp = sess.dropna(subset=['Break Reason']).pivot_table(index=['Date', 'User ID'], columns='Break Reason', values='BS', aggfunc='sum', fill_value=0).reset_index()
        sales['TS'] = sales['Talk Time'].apply(hms_to_sec)
        sales['IC'] = sales['TS'] >= 1
        sa = sales.groupby(['Date', 'User ID']).agg(T_OB=('call Id', 'count'), U_OB=('dstPhone', 'nunique')).reset_index()
        ca = sales[sales['IC']].groupby(['Date', 'User ID']).agg(C_OB=('call Id', 'count'), U_CC=('dstPhone', 'nunique')).reset_index()
        sf = pd.merge(sa, ca, on=['Date', 'User ID'], how='left').fillna(0)

        # Mapping & Final Table
        map_cols = {'Total Staffed Duration': 'Staffed Duration','Total Ready Duration': 'Ready Duration','Total Break Duration': 'Total Break Duration','Total Idle Time': 'Idle Time','Total Talk Time in Interval': 'Talk Time','Total ACW Duration in Interval': 'ACW Duration'}
        for k, v in map_cols.items(): prod[v+'_sec'] = prod[k].apply(hms_to_sec) if k in prod.columns else 0
        pf = prod.groupby(['Date', 'User ID', 'User Name']).agg({f"{n}_sec": "sum" for n in map_cols.values()}).reset_index()
        res = pd.merge(pf, sf, on=['Date', 'User ID'], how='left')
        res = pd.merge(res, bp, on=['Date', 'User ID'], how='left').fillna(0)

        # Formatting Output
        for c in [x for x in res.columns if x.endswith('_sec') or x in bp.columns]:
            if c not in ['Date', 'User ID']: res[c.replace('_sec','')] = res[c].apply(sec_to_hms)
        
        final = res[['Date', 'User Name', 'User ID', 'Staffed Duration', 'Ready Duration', 'Total Break Duration', 'Idle Time', 'Talk Time', 'ACW Duration', 'T_OB', 'C_OB', 'U_OB', 'U_CC']]
        final.columns = ['Date', 'Agent Name', 'Email', 'Staffed', 'Ready', 'Breaks', 'Idle', 'Talk', 'ACW', 'Total Calls', 'Connected', 'Unique OB', 'Unique CC']
        
        st.dataframe(final, use_container_width=True)
        excel_data = to_excel_formatted(final, "Sales_Report", "Sales Team Performance Database", has_summary_row=False)
        st.download_button("📥 Download Excel (Arial/Yellow)", data=excel_data, file_name="Sales_Report.xlsx")

# ==========================================
# PRE-SALES SECTION
# ==========================================
elif app_mode == "Pre-Sales SLA":
    st.title("📞 Pre-Sales SLA Report")
    with st.sidebar:
        acd_file = st.file_uploader("Upload ACD Call Details CSV", type="csv")

    if acd_file:
        df = pd.read_csv(acd_file)
        df.columns = df.columns.str.strip()
        df['Call Time'] = pd.to_datetime(df['Call Time'], dayfirst=True)
        df['Is_Ans'] = df['Username'].notna() & (df['Username'].astype(str).str.strip() != '')
        df['Hour'] = df['Call Time'].dt.hour

        # Hourly Stats
        hourly = df.groupby('Hour').agg(Received=('Call ID', 'count'), Answered=('Is_Ans', 'sum')).reset_index()
        hourly['Missed'] = hourly['Received'] - hourly['Answered']
        hourly['SLA %'] = (hourly['Answered'] / hourly['Received'] * 100).round(2).astype(str) + '%'
        hourly['Interval'] = hourly['Hour'].apply(lambda x: f"{x:02d}:00 - {x+1:02d}:00")
        
        # Add Total Footer Row
        totals = pd.DataFrame([{
            'Interval': 'Grand Total',
            'Received': hourly['Received'].sum(),
            'Answered': hourly['Answered'].sum(),
            'Missed': hourly['Missed'].sum(),
            'SLA %': f"{(hourly['Answered'].sum()/hourly['Received'].sum()*100):.2f}%"
        }])
        final_hourly = pd.concat([hourly[['Interval', 'Received', 'Answered', 'Missed', 'SLA %']], totals], ignore_index=True)

        st.subheader("Hourly SLA Breakdown")
        st.dataframe(final_hourly, use_container_width=True)

        excel_sla = to_excel_formatted(final_hourly, "PreSales_SLA", "Pre-Sales Inbound SLA Report", has_summary_row=True)
        st.download_button("📥 Download Excel (Arial/Yellow)", data=excel_sla, file_name="Pre_Sales_SLA_Report.xlsx")

import streamlit as st
import pandas as pd
import io

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

def to_excel_formatted(df, sheet_name, report_title, has_summary_row=False):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # FORMATS: Arial 11 + Yellow Background
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'
        })
        footer_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFFF00', 'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'
        })
        data_fmt = workbook.add_format({
            'font_name': 'Arial', 'font_size': 11, 'border': 1, 'align': 'center'
        })
        title_fmt = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 14, 'font_color': brand_navy
        })

        # Set Column Widths & Title
        worksheet.write(0, 0, report_title, title_fmt)
        worksheet.set_column(0, len(df.columns) - 1, 22)

        # Apply Header Format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(2, col_num, value, header_fmt)

        # Apply Data & Footer (Summary) Format
        last_row_idx = len(df) + 2
        for r_idx in range(len(df)):
            current_excel_row = r_idx + 3
            # If it's the last row AND we requested a summary highlight
            is_summary = has_summary_row and (current_excel_row == last_row_idx)
            row_style = footer_fmt if is_summary else data_fmt
            
            for c_idx in range(len(df.columns)):
                val = df.iloc[r_idx, c_idx]
                worksheet.write(current_excel_row, c_idx, val, row_style)
                
    return output.getvalue()

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.image("https://thesleepcompany.in/cdn/shop/files/Logo_White_300x.png?v=1614330134", width=200)
    st.header("Department Portal")
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
        try:
            prod, sess, sales = pd.read_csv(p_file), pd.read_csv(s_file), pd.read_csv(c_file)
            for d in [prod, sess, sales]: d.columns = d.columns.str.strip()
            
            # Date Parsing
            prod['Date'] = pd.to_datetime(prod['Interval Start'], dayfirst=True, errors='coerce').dt.date
            sess['Date'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.date
            sales['Date'] = pd.to_datetime(sales['Start Time'], dayfirst=True, errors='coerce').dt.date
            
            # Sales & Break Logic
            sales['TS'] = sales['Talk Time'].apply(hms_to_sec)
            sales['IC'] = sales['TS'] >= 1
            sa = sales.groupby(['Date', 'User ID']).agg(T_OB=('call Id', 'count'), U_OB=('dstPhone', 'nunique')).reset_index()
            ca = sales[sales['IC']].groupby(['Date', 'User ID']).agg(C_OB=('call Id', 'count'), U_CC=('dstPhone', 'nunique')).reset_index()
            sf = pd.merge(sa, ca, on=['Date', 'User ID'], how='left').fillna(0)

            # Map Duration Columns
            map_cols = {'Total Staffed Duration': 'Staffed','Total Ready Duration': 'Ready','Total Break Duration': 'Breaks','Total Idle Time': 'Idle','Total Talk Time in Interval': 'Talk','Total ACW Duration in Interval': 'ACW'}
            for k, v in map_cols.items(): 
                if k in prod.columns: prod[v+'_sec'] = prod[k].apply(hms_to_sec)
                else: prod[v+'_sec'] = 0

            pf = prod.groupby(['Date', 'User ID', 'User Name']).agg({f"{n}_sec": "sum" for n in map_cols.values()}).reset_index()
            res = pd.merge(pf, sf, on=['Date', 'User ID'], how='left').fillna(0)

            # Convert to HMS
            for c in [x for x in res.columns if x.endswith('_sec')]:
                res[c.replace('_sec','')] = res[c].apply(sec_to_hms)
            
            # Final Column Cleanup
            final_sales = res[['Date', 'User Name', 'User ID', 'Staffed', 'Ready', 'Breaks', 'Idle', 'Talk', 'ACW', 'T_OB', 'C_OB', 'U_OB', 'U_CC']]
            final_sales.columns = ['Date', 'Agent Name', 'Email', 'Staffed', 'Ready', 'Breaks', 'Idle', 'Talk', 'ACW', 'Total Calls', 'Connected', 'Unique OB', 'Unique CC']
            
            # Final Summary Row for Sales
            sales_totals = pd.DataFrame([{
                'Date': 'Grand Total', 'Agent Name': '-', 'Email': '-', 'Staffed': '-', 'Ready': '-', 'Breaks': '-', 'Idle': '-', 'Talk': '-', 'ACW': '-',
                'Total Calls': final_sales['Total Calls'].sum(),
                'Connected': final_sales['Connected'].sum(),
                'Unique OB': final_sales['Unique OB'].sum(),
                'Unique CC': final_sales['Unique CC'].sum()
            }])
            final_sales_with_total = pd.concat([final_sales, sales_totals], ignore_index=True)

            st.dataframe(final_sales_with_total, use_container_width=True)
            xl_sales = to_excel_formatted(final_sales_with_total, "Sales_Performance", "Sales Team Performance Report", has_summary_row=True)
            st.download_button("📥 Download Sales Excel", data=xl_sales, file_name="Sales_Performance_Report.xlsx")
        except Exception as e: st.error(f"Error: {e}")

# ==========================================
# PRE-SALES SECTION
# ==========================================
elif app_mode == "Pre-Sales SLA":
    st.title("📞 Pre-Sales SLA Report")
    with st.sidebar:
        acd_file = st.file_uploader("Upload ACD Call Details CSV", type="csv")

    if acd_file:
        try:
            df = pd.read_csv(acd_file)
            df.columns = df.columns.str.strip()
            df['Call Time'] = pd.to_datetime(df['Call Time'], dayfirst=True)
            # LOGIC: If Username exists, it was answered.
            df['Is_Ans'] = df['Username'].notna() & (df['Username'].astype(str).str.strip() != '')
            df['Hour'] = df['Call Time'].dt.hour

            hourly = df.groupby('Hour').agg(Received=('Call ID', 'count'), Answered=('Is_Ans', 'sum')).reset_index()
            hourly['Missed'] = hourly['Received'] - hourly['Answered']
            hourly['SLA %'] = (hourly['Answered'] / hourly['Received'] * 100).round(2).astype(str) + '%'
            hourly['Interval'] = hourly['Hour'].apply(lambda x: f"{x:02d}:00 - {x+1:02d}:00")
            
            # Grand Total Logic
            totals = pd.DataFrame([{
                'Interval': 'Grand Total',
                'Received': hourly['Received'].sum(),
                'Answered': hourly['Answered'].sum(),
                'Missed': hourly['Missed'].sum(),
                'SLA %': f"{(hourly['Answered'].sum()/hourly['Received'].sum()*100):.2f}%"
            }])
            final_hourly = pd.concat([hourly[['Interval', 'Received', 'Answered', 'Missed', 'SLA %']], totals], ignore_index=True)

            st.subheader("Inbound SLA Summary")
            st.dataframe(final_hourly, use_container_width=True)

            xl_pre = to_excel_formatted(final_hourly, "SLA_Report", "Pre-Sales Inbound SLA Management Report", has_summary_row=True)
            st.download_button("📥 Download Pre-Sales Excel", data=xl_pre, file_name="Pre_Sales_SLA_Report.xlsx")
        except Exception as e: st.error(f"Error: {e}")

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

def to_excel_formatted_multi(sheets_dict, report_title):
    """
    sheets_dict: { 'SheetName': {'df': dataframe, 'title': 'Table Title', 'has_summary': bool} }
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # FORMATS
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

            # Apply Data & Footer
            last_row_idx = len(df) + 2
            for r_idx in range(len(df)):
                curr_row = r_idx + 3
                is_sum = content['has_summary'] and (curr_row == last_row_idx)
                style = footer_fmt if is_sum else data_fmt
                for c_idx in range(len(df.columns)):
                    worksheet.write(curr_row, c_idx, df.iloc[r_idx, c_idx], style)
                    
    return output.getvalue()

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.image("https://thesleepcompany.in/cdn/shop/files/new_logo.webp?v=1706780127", width=200)
    st.header("CRM Department Portal")
    app_mode = st.radio("Choose App", ["Sales Performance", "Pre-Sales Break and SLA"])
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
            prod, sess, sales = pd.read_csv(p_file), pd.read_csv(s_file), pd.read_csv(c_file)
            for d in [prod, sess, sales]: d.columns = d.columns.str.strip()
            
            prod['Date'] = pd.to_datetime(prod['Interval Start'], dayfirst=True, errors='coerce').dt.date
            sess['Date'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.date
            sales['Date'] = pd.to_datetime(sales['Start Time'], dayfirst=True, errors='coerce').dt.date
            
            sales['TS'] = sales['Talk Time'].apply(hms_to_sec)
            sales['IC'] = sales['TS'] >= 1
            sa = sales.groupby(['Date', 'User ID']).agg(T_OB=('call Id', 'count'), U_OB=('dstPhone', 'nunique')).reset_index()
            ca = sales[sales['IC']].groupby(['Date', 'User ID']).agg(C_OB=('call Id', 'count'), U_CC=('dstPhone', 'nunique')).reset_index()
            sf = pd.merge(sa, ca, on=['Date', 'User ID'], how='left').fillna(0)

            map_cols = {'Total Staffed Duration': 'Staffed','Total Ready Duration': 'Ready','Total Break Duration': 'Breaks','Total Idle Time': 'Idle','Total Talk Time in Interval': 'Talk','Total ACW Duration in Interval': 'ACW'}
            for k, v in map_cols.items(): prod[v+'_sec'] = prod[k].apply(hms_to_sec) if k in prod.columns else 0

            pf = prod.groupby(['Date', 'User ID', 'User Name']).agg({f"{n}_sec": "sum" for n in map_cols.values()}).reset_index()
            res = pd.merge(pf, sf, on=['Date', 'User ID'], how='left').fillna(0)

            for c in [x for x in res.columns if x.endswith('_sec')]:
                res[c.replace('_sec','')] = res[c].apply(sec_to_hms)
            
            final_sales = res[['Date', 'User Name', 'User ID', 'Staffed', 'Ready', 'Breaks', 'Idle', 'Talk', 'ACW', 'T_OB', 'C_OB', 'U_OB', 'U_CC']]
            final_sales.columns = ['Date', 'Agent Name', 'Email', 'Staffed', 'Ready', 'Breaks', 'Idle', 'Talk', 'ACW', 'Total Calls', 'Connected', 'Unique OB', 'Unique CC']
            
            st.dataframe(final_sales, use_container_width=True)
            
            sheets = {'Sales_Report': {'df': final_sales, 'title': 'Sales Performance Report', 'has_summary': False}}
            xl = to_excel_formatted_multi(sheets, "Sales Report")
            st.download_button("📥 Download Sales Excel", data=xl, file_name="Sales_Report.xlsx")
        except Exception as e: st.error(f"Error: {e}")

# ==========================================
# 2. PRE-SALES BREAK & SLA SECTION
# ==========================================
elif app_mode == "Pre-Sales Break and SLA":
    st.title("📞 Pre-Sales Break & SLA Report")
    with st.sidebar:
        acd_file = st.file_uploader("1. Upload ACD Call Details (SLA)", type="csv")
        sess_file = st.file_uploader("2. Upload Session Details (Breaks)", type="csv")

    if acd_file and sess_file:
        try:
            # --- SLA LOGIC ---
            acd = pd.read_csv(acd_file)
            acd.columns = acd.columns.str.strip()
            acd['Call Time'] = pd.to_datetime(acd['Call Time'], dayfirst=True)
            acd['Date'] = acd['Call Time'].dt.date
            acd['Hour'] = acd['Call Time'].dt.hour
            acd['Is_Ans'] = acd['Username'].notna() & (acd['Username'].astype(str).str.strip() != '')

            # Filter Completed Hours
            now_h = datetime.now().hour
            today = datetime.now().date()
            acd_f = acd[~((acd['Date'] == today) & (acd['Hour'] >= now_h))]

            # SLA Hour-wise
            h_sla = acd_f.groupby('Hour').agg(Received=('Call ID', 'count'), Answered=('Is_Ans', 'sum')).reset_index()
            h_sla['Missed'] = h_sla['Received'] - h_sla['Answered']
            h_sla['SLA %'] = (h_sla['Answered'] / h_sla['Received'] * 100).round(2).astype(str) + '%'
            h_sla['Interval'] = h_sla['Hour'].apply(lambda x: f"{x:02d}:00 - {x+1:02d}:00")
            h_sla_final = h_sla[['Interval', 'Received', 'Answered', 'Missed', 'SLA %']]
            
            # SLA Day-wise
            d_sla = acd_f.groupby('Date').agg(Received=('Call ID', 'count'), Answered=('Is_Ans', 'sum')).reset_index()
            d_sla['Missed'] = d_sla['Received'] - d_sla['Answered']
            d_sla['SLA %'] = (d_sla['Answered'] / d_sla['Received'] * 100).round(2).astype(str) + '%'
            
            # --- BREAK LOGIC ---
            sess = pd.read_csv(sess_file)
            sess.columns = sess.columns.str.strip()
            sess['Date'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.date
            sess['Hour'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.hour
            sess['BS'] = sess['Break Duration'].apply(hms_to_sec)
            
            # Breaks Day-wise (Agent Summary)
            d_brk = sess.dropna(subset=['Break Reason']).pivot_table(index=['Date', 'Username'], columns='Break Reason', values='BS', aggfunc='sum', fill_value=0).reset_index()
            for c in [x for x in d_brk.columns if x not in ['Date', 'Username']]:
                d_brk[c] = d_brk[c].apply(sec_to_hms)

            # Breaks Hour-wise
            h_brk = sess.dropna(subset=['Break Reason']).pivot_table(index=['Date', 'Hour'], columns='Break Reason', values='BS', aggfunc='sum', fill_value=0).reset_index()
            h_brk['Interval'] = h_brk['Hour'].apply(lambda x: f"{x:02d}:00 - {x+1:02d}:00")
            cols_to_move = [c for c in h_brk.columns if c not in ['Date', 'Hour', 'Interval']]
            h_brk_final = h_brk[['Date', 'Interval'] + cols_to_move]
            for c in cols_to_move: h_brk_final[c] = h_brk_final[c].apply(sec_to_hms)

            # --- DISPLAY & EXPORT ---
            st.subheader("SLA Analysis")
            st.dataframe(h_sla_final, use_container_width=True)
            
            st.subheader("Agent Break Analysis")
            st.dataframe(d_brk, use_container_width=True)

            sheets = {
                'SLA_Analysis': {'df': h_sla_final, 'title': 'Hourly SLA Report (Completed Hours)', 'has_summary': False},
                'SLA_Daywise': {'df': d_sla, 'title': 'Daily SLA Summary', 'has_summary': False},
                'Break_Daywise': {'df': d_brk, 'title': 'Agent Break Summary (Day-wise)', 'has_summary': False},
                'Break_Hourwise': {'df': h_brk_final, 'title': 'Hourly Break Distribution', 'has_summary': False}
            }
            xl_presales = to_excel_formatted_multi(sheets, "Pre-Sales Break and SLA Report")
            st.download_button("📥 Download Combined Pre-Sales Excel", data=xl_presales, file_name="Pre_Sales_Break_SLA_Report.xlsx")
            
        except Exception as e: st.error(f"Error: {e}")

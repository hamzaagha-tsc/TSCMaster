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

# --- THE SLEEP COMPANY BRANDING (Aggressive CSS Fix for Light Mode) ---
brand_navy = "#102a51"
brand_copper = "#c59d5f"

st.markdown(f"""
    <style>
    /* 1. FORCE SIDEBAR BACKGROUND & TEXT */
    [data-testid="stSidebar"] {{
        background-color: {brand_navy} !important;
    }}
    /* Force all text inside sidebar to be white */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] .stRadio label,
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {{
        color: white !important;
    }}
    
    /* 2. STYLE BUTTONS */
    div.stButton > button:first-child {{
        background-color: {brand_copper};
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: bold;
    }}
    
    /* 3. GENERAL STYLING */
    h1, h2, h3 {{
        color: {brand_navy};
        font-family: 'Arial';
    }}
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

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)

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
    st.header("🏢 Department Portal")
    app_mode = st.radio("Choose App", ["Sales Performance", "Pre-Sales SLA & Breaks"])
    st.divider()

# ==========================================
# 1. SALES TEAM SECTION (Now with Break Details)
# ==========================================
if app_mode == "Sales Performance":
    st.title("🚀 Sales Detailed Performance")
    with st.sidebar:
        p_file = st.file_uploader("1. Productivity Summary", type="csv")
        s_file = st.file_uploader("2. Session Details", type="csv")
        c_file = st.file_uploader("3. Custom Sales Report", type="csv")

    if p_file and s_file and c_file:
        try:
            prod, sess, sales = pd.read_csv(p_file), pd.read_csv(s_file), pd.read_csv(c_file)
            for d in [prod, sess, sales]: d.columns = d.columns.str.strip()
            
            # Dates
            prod['Date'] = pd.to_datetime(prod['Interval Start'], dayfirst=True, errors='coerce').dt.date
            sess['Date'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.date
            sales['Date'] = pd.to_datetime(sales['Start Time'], dayfirst=True, errors='coerce').dt.date
            
            # Sales & Unique Calls
            sales['TS'] = sales['Talk Time'].apply(hms_to_sec)
            sales['IC'] = sales['TS'] >= 1
            sa = sales.groupby(['Date', 'User ID']).agg(T_OB=('call Id', 'count'), U_OB=('dstPhone', 'nunique')).reset_index()
            ca = sales[sales['IC']].groupby(['Date', 'User ID']).agg(C_OB=('call Id', 'count'), U_CC=('dstPhone', 'nunique')).reset_index()
            sf = pd.merge(sa, ca, on=['Date', 'User ID'], how='left').fillna(0)

            # --- DETAILED BREAKS FOR SALES ---
            sess['BS'] = sess['Break Duration'].apply(hms_to_sec)
            break_pivot = sess.dropna(subset=['Break Reason']).pivot_table(
                index=['Date', 'User ID'], columns='Break Reason', values='BS', aggfunc='sum', fill_value=0
            ).reset_index()

            # Productivity
            map_cols = {'Total Staffed Duration': 'Staffed','Total Ready Duration': 'Ready','Total Break Duration': 'Total Breaks','Total Idle Time': 'Idle','Total Talk Time in Interval': 'Talk','Total ACW Duration in Interval': 'ACW'}
            for k, v in map_cols.items(): 
                prod[v+'_sec'] = prod[k].apply(hms_to_sec) if k in prod.columns else 0

            pf = prod.groupby(['Date', 'User ID', 'User Name']).agg({f"{n}_sec": "sum" for n in map_cols.values()}).reset_index()
            res = pd.merge(pf, sf, on=['Date', 'User ID'], how='left').fillna(0)
            res = pd.merge(res, break_pivot, on=['Date', 'User ID'], how='left').fillna(0)

            # Convert to HMS
            time_cols = [x for x in res.columns if x.endswith('_sec')]
            break_cols = [x for x in break_pivot.columns if x not in ['Date', 'User ID']]
            for c in time_cols + break_cols:
                res[c.replace('_sec','')] = res[c].apply(sec_to_hms)
            
            final_order = ['Date', 'User Name', 'User ID', 'Staffed', 'Ready', 'Total Breaks'] + break_cols + ['Idle', 'Talk', 'ACW', 'T_OB', 'C_OB', 'U_OB', 'U_CC']
            result = res[[c for c in final_order if c in res.columns]].copy()
            
            st.dataframe(result, use_container_width=True)
            xl = to_excel_formatted_multi({'Sales_Report': {'df': result, 'title': 'Sales Detailed Performance', 'has_summary': False}})
            st.download_button("📥 Download Detailed Sales Excel", data=xl, file_name="Sales_Detailed_Report.xlsx")
        except Exception as e: st.error(f"Error: {e}")

# ==========================================
# 2. PRE-SALES SECTION
# ==========================================
elif app_mode == "Pre-Sales SLA & Breaks":
    st.title("📞 Pre-Sales Performance Hub")
    with st.sidebar:
        pre_mode = st.radio("Report Type", ["Only SLA", "Both SLA and Break Reports"])
        st.divider()
        acd_file = st.file_uploader("1. ACD Call Details (SLA)", type="csv")
        sess_file = st.file_uploader("2. Session Details (Breaks)", type="csv") if "Both" in pre_mode else None

    if acd_file:
        try:
            acd = pd.read_csv(acd_file)
            acd.columns = acd.columns.str.strip()
            acd['Call Time'] = pd.to_datetime(acd['Call Time'], dayfirst=True)
            acd['Date'] = acd['Call Time'].dt.date
            acd['Hour'] = acd['Call Time'].dt.hour
            acd['Is_Ans'] = acd['Username'].notna() & (acd['Username'].astype(str).str.strip() != '')

            # Smart Filter: Completed Hours only
            max_t = acd['Call Time'].max()
            acd_f = acd[~((acd['Date'] == max_t.date()) & (acd['Hour'] >= max_t.hour))]

            if acd_f.empty: st.warning("All data is from the current hour. No completed hours found yet.")
            else:
                h_sla = acd_f.groupby('Hour').agg(Rec=('Call ID', 'count'), Ans=('Is_Ans', 'sum')).reset_index()
                h_sla['Miss'] = h_sla['Rec'] - h_sla['Ans']
                h_sla['SLA %'] = h_sla.apply(lambda r: f"{(r['Ans']/r['Rec']*100):.2f}%" if r['Rec'] > 0 else "0.00%", axis=1)
                h_sla['Interval'] = h_sla['Hour'].apply(lambda x: f"{x:02d}:00 - {x+1:02d}:00")
                
                tr, ta = h_sla['Rec'].sum(), h_sla['Ans'].sum()
                t_h = pd.DataFrame([{'Interval': 'Grand Total', 'Rec': tr, 'Ans': ta, 'Miss': tr-ta, 'SLA %': f"{(ta/tr*100):.2f}%" if tr>0 else "0.00%"}])
                h_final = pd.concat([h_sla[['Interval', 'Rec', 'Ans', 'Miss', 'SLA %']], t_h], ignore_index=True)

                sheets = {'Hourly_SLA': {'df': h_final, 'title': 'Hourly SLA (Completed)', 'has_summary': True}}

                if "Both" in pre_mode and sess_file:
                    s_df = pd.read_csv(sess_file)
                    s_df.columns = s_df.columns.str.strip()
                    s_df['BS'] = s_df['Break Duration'].apply(hms_to_sec)
                    s_df['Date'] = pd.to_datetime(s_df['Login Time'], dayfirst=True, errors='coerce').dt.date
                    d_brk = s_df.dropna(subset=['Break Reason']).pivot_table(index=['Date', 'Username'], columns='Break Reason', values='BS', aggfunc='sum', fill_value=0).reset_index()
                    for c in [x for x in d_brk.columns if x not in ['Date', 'Username']]: d_brk[c] = d_brk[c].apply(sec_to_hms)
                    sheets['Break_Summary'] = {'df': d_brk, 'title': 'Agent Break Details', 'has_summary': False}

                st.subheader("SLA Overview")
                st.dataframe(h_final, use_container_width=True)
                xl_data = to_excel_formatted_multi(sheets)
                st.download_button("📥 Download Pre-Sales Excel", data=xl_data, file_name="Pre_Sales_SLA_Break_Report.xlsx")
        except Exception as e: st.error(f"Error: {e}")

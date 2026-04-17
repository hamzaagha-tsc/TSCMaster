import streamlit as st
import pandas as pd
import numpy as np

# --- APP CONFIGURATION ---
st.set_page_config(
    page_title="TSC | Agent Performance", 
    page_icon="https://thesleepcompany.in/cdn/shop/files/fav-icon_32x32.png", 
    layout="wide"
)
st.title ("TSC Unified Agent Database - Hamza Agha")

# --- THE SLEEP COMPANY BRANDING ---
brand_navy = "#102a51"
brand_copper = "#c59d5f"

st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ background-color: {brand_navy}; }}
    [data-testid="stSidebar"] * {{ color: white; }}
    div.stButton > button:first-child {{
        background-color: {brand_copper};
        color: white;
        border-radius: 5px;
        border: none;
    }}
    h1, h2, h3 {{ color: {brand_navy}; }}
    </style>
    """, unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---
def hms_to_sec(t):
    if pd.isna(t) or t == '0' or t == 0: return 0
    try:
        t_clean = str(t).split('.')[0].strip() 
        parts = t_clean.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        return 0
    except:
        return 0

def sec_to_hms(seconds):
    if pd.isna(seconds) or seconds <= 0: return "00:00:00"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

# --- SIDEBAR UPLOADS ---
with st.sidebar:
    st.header("Upload Reports")
    prod_file = st.file_uploader("1. Productivity Summary (Agent Productivity Interval Report)", type="csv")
    sess_file = st.file_uploader("2. Session Details (Agent Session Details Report)", type="csv")
    sales_file = st.file_uploader("3. Custom Calls Report (Custom Report Sales)", type="csv")

if prod_file and sess_file and sales_file:
    try:
        # 1. Load Data
        prod = pd.read_csv(prod_file)
        sess = pd.read_csv(sess_file)
        sales = pd.read_csv(sales_file)

        # Clean Column Names
        for df in [prod, sess, sales]:
            df.columns = df.columns.str.strip()

        # 2. Cleaning Data Rows
        for df in [prod, sess, sales]:
            df.dropna(subset=['User ID'], inplace=True)
            df['User ID'] = df['User ID'].astype(str).str.lower().str.strip()

        # Process Dates
        prod['Date'] = pd.to_datetime(prod['Interval Start'], dayfirst=True, errors='coerce').dt.date
        sess['Date'] = pd.to_datetime(sess['Login Time'], dayfirst=True, errors='coerce').dt.date
        sales['Date'] = pd.to_datetime(sales['Start Time'], dayfirst=True, errors='coerce').dt.date

        # 3. Process Individual Breaks
        sess['Break_Sec'] = sess['Break Duration'].apply(hms_to_sec)
        sess_breaks = sess.dropna(subset=['Break Reason']).copy()
        break_pivot = sess_breaks.pivot_table(
            index=['Date', 'User ID'], 
            columns='Break Reason', 
            values='Break_Sec', 
            aggfunc='sum', 
            fill_value=0
        ).reset_index()

        # 4. Process Sales Metrics
        sales['Talk_Sec'] = sales['Talk Time'].apply(hms_to_sec)
        sales['Is_Connected'] = sales['Talk_Sec'] >= 1
        
        sales_agg = sales.groupby(['Date', 'User ID']).agg(
            Total_OB_Calls=('call Id', 'count'),
            Unq_OB_Calls=('dstPhone', 'nunique')
        ).reset_index()

        conn_agg = sales[sales['Is_Connected']].groupby(['Date', 'User ID']).agg(
            Connected_OB_Calls=('call Id', 'count'),
            Unq_CC_Calls=('dstPhone', 'nunique')
        ).reset_index()
        sales_final = pd.merge(sales_agg, conn_agg, on=['Date', 'User ID'], how='left').fillna(0)

        # 5. Process Productivity (Mapping Actual CSV Names to Your Report Names)
        # We explicitly link your required names to the CSV's long names
        mapping = {
            'Total Staffed Duration': 'Staffed Duration',
            'Total Ready Duration': 'Ready Duration',
            'Total Break Duration': 'Total Break Duration',
            'Total Idle Time': 'Idle Time',
            'Total Talk Time in Interval': 'Talk Time',
            'Total ACW Duration in Interval': 'ACW Duration'
        }
        
        for csv_col, target_name in mapping.items():
            if csv_col in prod.columns:
                prod[target_name + '_sec'] = prod[csv_col].apply(hms_to_sec)
            else:
                st.warning(f"Note: '{csv_col}' not found in file. Using 0s.")
                prod[target_name + '_sec'] = 0

        # Aggregate Productivity
        prod_final = prod.groupby(['Date', 'User ID', 'User Name']).agg({
            f"{name}_sec": "sum" for name in mapping.values()
        }).reset_index()

        # 6. Final Merge
        final_df = pd.merge(prod_final, sales_final, on=['Date', 'User ID'], how='left')
        final_df = pd.merge(final_df, break_pivot, on=['Date', 'User ID'], how='left').fillna(0)

        # 7. Convert Sec to HH:MM:SS
        sec_cols = [c for c in final_df.columns if c.endswith('_sec')]
        break_cols = [c for c in break_pivot.columns if c not in ['Date', 'User ID']]
        
        for col in sec_cols + break_cols:
            clean_name = col.replace('_sec', '').replace('Total ', '')
            final_df[clean_name] = final_df[col].apply(sec_to_hms)

        # 8. Column Selection
        desired_order = [
            'Date', 'User Name', 'User ID', 'Staffed Duration', 'Ready Duration', 'Total Break Duration',
            'After Call Work', 'Lunch', 'First Break', 'Last Break', 'Meeting', 'Miscellaneous',
            'Idle Time', 'Talk Time', 'ACW Duration', 
            'Total_OB_Calls', 'Connected_OB_Calls', 'Unq_OB_Calls', 'Unq_CC_Calls'
        ]
        
        final_columns = [col for col in desired_order if col in final_df.columns]
        result = final_df[final_columns].copy()
        
        # Numeric Clean up
        count_cols = ['Total_OB_Calls', 'Connected_OB_Calls', 'Unq_OB_Calls', 'Unq_CC_Calls']
        for col in count_cols:
            if col in result.columns:
                result[col] = result[col].astype(int)

        # KPI Metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("Active Agents", len(result['User ID'].unique()))
        m2.metric("Total Connected Calls", result['Connected_OB_Calls'].sum())
        m3.metric("Unique Customers Reached", result['Unq_CC_Calls'].sum())

        st.success("Report Generated Successfully!")
        st.dataframe(result, use_container_width=True)

        csv = result.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Performance Database", data=csv, file_name="TSC_Performance_Report.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Error: {e}")

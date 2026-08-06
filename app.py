import streamlit as st
import pandas as pd

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="NexSys Operations Analyzer", layout="wide")
st.title("NexSys Operations & Financial Analyzer")
st.markdown("Detailed tabular analysis across Parts Usage, Jobs, Invoices, and Timesheets for **Simple Installs** and **Water Heaters**.")

# --- HELPER FUNCTIONS ---
def map_tech_name(fleet_name):
    """Normalizes fleet and user names across datasets."""
    name = str(fleet_name).strip()
    if "Bill" in name: return "Bill"
    if "Bryan" in name: return "Bryan Pickett"
    if "Carmen" in name: return "Carmen Tripodi"
    if "Erik" in name: return "Erik Tange"
    if "Matt S" in name or "Matt's" in name or "Mathew" in name or "Matt Schlosser" in name: return "Matt Schlosser"
    if "Sean" in name: return "Sean Marble"
    if "Tanner" in name: return "Tanner LaForge"
    return name

@st.cache_data
def load_google_sheet(url):
    """Reads a public Google Sheet URL into a dictionary of DataFrames."""
    if not url:
        return None
    base_url = url.split('/edit')[0]
    export_url = f"{base_url}/export?format=xlsx"
    try:
        xls = pd.read_excel(export_url, sheet_name=None)
        return xls
    except Exception as e:
        st.error(f"Error reading Google Sheet. Ensure link permissions are public ('Anyone with the link can view'). Details: {e}")
        return None

def process_parts_df(df, business_unit):
    """Cleans parts inventory DataFrame from Google Sheets and adjusts for returned items."""
    if df is None or df.empty:
        return pd.DataFrame()
        
    df = df.copy()
    if 'Transferred To' not in df.columns or 'Total Value' not in df.columns:
        return pd.DataFrame()
        
    if df['Total Value'].dtype == object:
        df['Total Value'] = df['Total Value'].astype(str).replace(r'[\$,]', '', regex=True)
    
    df['Total Value'] = pd.to_numeric(df['Total Value'], errors='coerce').fillna(0)
    
    if 'Qty' in df.columns:
        df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)

    # --- RETURN DEDUCTION LOGIC ---
    # Items marked as 'Return', 'Returned', or 'In' are subtracted from usage totals
    if 'Direction' in df.columns:
        is_return = df['Direction'].astype(str).str.strip().str.lower().isin(['return', 'returned', 'in'])
        df.loc[is_return, 'Total Value'] = -df.loc[is_return, 'Total Value']
        if 'Qty' in df.columns:
            df.loc[is_return, 'Qty'] = -df.loc[is_return, 'Qty']

    df['Transferred To'] = df['Transferred To'].astype(str).str.replace("Matt's TransitFleet", "Matt S")
    df['Tech'] = df['Transferred To'].apply(map_tech_name)
    df['Business Unit'] = business_unit
    return df

def read_uploaded_csv(file_obj):
    """Reads uploaded CSV file safely by resetting the stream pointer."""
    if file_obj is None:
        return None
    file_obj.seek(0)
    try:
        df = pd.read_csv(file_obj, header=1)
        # Verify that header=1 parsed valid columns, otherwise fallback to header=0
        if not any(col in df.columns for col in ['Business Unit', 'Business Unit.1', 'User', 'Assigned Team Members']):
            file_obj.seek(0)
            df = pd.read_csv(file_obj, header=0)
        return df
    except Exception:
        file_obj.seek(0)
        return pd.read_csv(file_obj, header=0)

# --- SIDEBAR FILTERS & DATA SOURCES ---
st.sidebar.header("📁 Data Sources")

sheet_url = st.sidebar.text_input(
    "Google Sheets Parts URL", 
    value="https://docs.google.com/spreadsheets/d/1OR4mEgviGglKNLwinPLnc8NB3FrN7VH9rt5qAvo8RRs/edit?usp=sharing"
)

uploaded_jobs = st.sidebar.file_uploader("Upload 'all jobs.csv'", type=['csv'])
uploaded_invoices = st.sidebar.file_uploader("Upload 'invoices.csv'", type=['csv'])
uploaded_timesheets = st.sidebar.file_uploader("Upload 'timesheets.csv'", type=['csv'])

TARGET_BUS = ['Lowes - Simple Installs', 'Lowes - Water Heaters']

# --- PRE-PROCESS ALL DATA ONCE ---
# 1. Parts Data from Google Sheets
sheets_dict = load_google_sheet(sheet_url)
df_parts = pd.DataFrame()
if sheets_dict:
    sheet_names = list(sheets_dict.keys())
    simple_sheet = sheet_names[0] if len(sheet_names) > 0 else None
    wh_sheet = sheet_names[1] if len(sheet_names) > 1 else simple_sheet
    
    df_simple_p = process_parts_df(sheets_dict[simple_sheet], 'Lowes - Simple Installs') if simple_sheet else pd.DataFrame()
    df_wh_p = process_parts_df(sheets_dict[wh_sheet], 'Lowes - Water Heaters') if wh_sheet else pd.DataFrame()
    df_parts = pd.concat([df_simple_p, df_wh_p], ignore_index=True)

# 2. Jobs Data
raw_jobs_df = read_uploaded_csv(uploaded_jobs)
jobs_filtered = pd.DataFrame()
if raw_jobs_df is not None and 'Business Unit' in raw_jobs_df.columns:
    jobs_filtered = raw_jobs_df[raw_jobs_df['Business Unit'].isin(TARGET_BUS)].copy()
    jobs_filtered['Invoice Amount'] = pd.to_numeric(jobs_filtered['Total Invoice Amount'], errors='coerce').fillna(0)

# 3. Invoices Data (Excludes Draft & Void)
raw_inv_df = read_uploaded_csv(uploaded_invoices)
inv_filtered = pd.DataFrame()
if raw_inv_df is not None:
    bu_col = 'Business Unit.1' if 'Business Unit.1' in raw_inv_df.columns else ('Business Unit' if 'Business Unit' in raw_inv_df.columns else None)
    if bu_col:
        inv_filtered = raw_inv_df[
            raw_inv_df[bu_col].isin(TARGET_BUS) & 
            ~raw_inv_df['Status'].astype(str).str.lower().str.contains('draft|void', na=False)
        ].copy()
        inv_filtered['Invoice Total'] = pd.to_numeric(inv_filtered['Invoice Total'], errors='coerce').fillna(0)
        inv_filtered['Business Unit Clean'] = inv_filtered[bu_col]

# 4. Timesheets Data
ts_df = read_uploaded_csv(uploaded_timesheets)
if ts_df is not None and 'Clock In Date/Time' in ts_df.columns:
    ts_df['In'] = pd.to_datetime(ts_df['Clock In Date/Time'], errors='coerce')
    ts_df['Out'] = pd.to_datetime(ts_df['Clock Out Date/Time'], errors='coerce')
    ts_df['Hours'] = (ts_df['Out'] - ts_df['In']).dt.total_seconds() / 3600.0

# --- TABS ---
tab_exec, tab_parts, tab_jobs, tab_inv, tab_ts = st.tabs([
    "📈 Executive Summary Table",
    "⚙️ Parts Usage",
    "📋 Jobs Analysis",
    "💳 Invoices Analysis",
    "⏱️ Timesheets Analysis"
])

# --- TAB 1: EXECUTIVE SUMMARY TABLE ---
with tab_exec:
    st.header("Technician Level Master Summary Table")
    st.markdown("Consolidated view combining net parts cost (factoring in returns), job counts, invoice revenue, and timesheet hours.")
    
    exec_data = {}
    
    if not df_parts.empty:
        p_agg = df_parts.groupby('Tech')['Total Value'].sum()
        for tech, val in p_agg.items():
            exec_data.setdefault(tech, {})['Net Parts Cost'] = val

    if not jobs_filtered.empty and 'Assigned Team Members' in jobs_filtered.columns:
        j_agg = jobs_filtered.groupby('Assigned Team Members')['#ID'].count()
        for tech, val in j_agg.items():
            exec_data.setdefault(tech, {})['Jobs Completed'] = val

    if not inv_filtered.empty and 'Assigned Team Members' in inv_filtered.columns:
        i_agg = inv_filtered.groupby('Assigned Team Members')['Invoice Total'].sum()
        for tech, val in i_agg.items():
            exec_data.setdefault(tech, {})['Total Revenue'] = val

    if ts_df is not None and 'User' in ts_df.columns:
        t_agg = ts_df.groupby('User')['Hours'].sum()
        for tech, val in t_agg.items():
            exec_data.setdefault(tech, {})['Logged Hours'] = val

    if exec_data:
        master_df = pd.DataFrame.from_dict(exec_data, orient='index').fillna(0).reset_index()
        master_df.rename(columns={'index': 'Technician / Team'}, inplace=True)
        
        if 'Net Parts Cost' in master_df.columns:
            master_df['Net Parts Cost'] = master_df['Net Parts Cost'].map('${:,.2f}'.format)
        if 'Total Revenue' in master_df.columns:
            master_df['Total Revenue'] = master_df['Total Revenue'].map('${:,.2f}'.format)
        if 'Logged Hours' in master_df.columns:
            master_df['Logged Hours'] = master_df['Logged Hours'].map('{:,.2f} hrs'.format)
        if 'Jobs Completed' in master_df.columns:
            master_df['Jobs Completed'] = master_df['Jobs Completed'].astype(int)

        st.dataframe(master_df, use_container_width=True, hide_index=True)
    else:
        st.info("Upload source files in the sidebar to populate the master summary table.")

# --- TAB 2: PARTS USAGE ---
with tab_parts:
    st.header("Parts Usage Analysis (Net Usage)")
    st.caption("Note: Items marked with Direction = 'Return' or 'Returned' are automatically subtracted from parts costs and quantities.")
    
    if not df_parts.empty:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Net Parts Value by Technician & Business Unit")
            t_parts = df_parts.groupby(['Tech', 'Business Unit'])['Total Value'].sum().reset_index()
            t_parts.rename(columns={'Total Value': 'Net Parts Cost'}, inplace=True)
            t_parts['Net Parts Cost'] = t_parts['Net Parts Cost'].map('${:,.2f}'.format)
            st.dataframe(t_parts, use_container_width=True, hide_index=True)
            
        with col2:
            st.subheader("Net Parts Value by Business Unit")
            bu_parts = df_parts.groupby('Business Unit')['Total Value'].agg(['count', 'sum']).reset_index()
            bu_parts.columns = ['Business Unit', 'Line Items', 'Net Parts Cost']
            bu_parts['Net Parts Cost'] = bu_parts['Net Parts Cost'].map('${:,.2f}'.format)
            st.dataframe(bu_parts, use_container_width=True, hide_index=True)

        st.subheader("Item-Level Net Quantity Used")
        item_parts = df_parts.groupby(['Business Unit', 'SKU', 'Item'])['Qty'].sum().reset_index()
        item_parts.rename(columns={'Qty': 'Net Qty Used'}, inplace=True)
        st.dataframe(item_parts, use_container_width=True, hide_index=True)
    else:
        st.info("Google Sheet parts data not loaded.")

# --- TAB 3: JOBS ANALYSIS ---
with tab_jobs:
    st.header("Jobs Analysis")
    if not jobs_filtered.empty:
        st.subheader("Job Summary by Tech, Business Unit & Job Title")
        j_summary = jobs_filtered.groupby(['Assigned Team Members', 'Business Unit', 'Title']).agg(
            Job_Count=('#ID', 'count'),
            Total_Invoice_Amount=('Invoice Amount', 'sum')
        ).reset_index()
        j_summary.rename(columns={'Assigned Team Members': 'Technician', 'Title': 'Job Title'}, inplace=True)
        j_summary['Total_Invoice_Amount'] = j_summary['Total_Invoice_Amount'].map('${:,.2f}'.format)
        
        st.dataframe(j_summary.sort_values(by=['Technician', 'Business Unit']), use_container_width=True, hide_index=True)

        st.subheader("Job Summary by Business Unit Level")
        bu_jobs = jobs_filtered.groupby('Business Unit').agg(
            Total_Jobs=('#ID', 'count'),
            Total_Billed=('Invoice Amount', 'sum')
        ).reset_index()
        bu_jobs['Total_Billed'] = bu_jobs['Total_Billed'].map('${:,.2f}'.format)
        st.dataframe(bu_jobs, use_container_width=True, hide_index=True)
    else:
        st.info("Upload 'all jobs.csv' in the sidebar.")

# --- TAB 4: INVOICES ANALYSIS ---
with tab_inv:
    st.header("Invoices Analysis")
    st.caption("Note: Invoices with status 'Draft' or 'Void' are excluded from analysis.")
    
    if not inv_filtered.empty:
        st.subheader("Invoices Summary by Tech, Business Unit & Parent Job Title")
        inv_sum = inv_filtered.groupby(['Assigned Team Members', 'Business Unit Clean', 'Parent Job Title', 'Status']).agg(
            Invoice_Count=('#ID', 'count'),
            Total_Amount=('Invoice Total', 'sum')
        ).reset_index()
        
        inv_sum.rename(columns={'Assigned Team Members': 'Technician', 'Business Unit Clean': 'Business Unit'}, inplace=True)
        inv_sum['Total_Amount'] = inv_sum['Total_Amount'].map('${:,.2f}'.format)
        
        st.dataframe(inv_sum.sort_values(by=['Technician', 'Business Unit']), use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Invoice Totals by Business Unit")
            bu_inv = inv_filtered.groupby('Business Unit Clean').agg(
                Invoices_Count=('#ID', 'count'),
                Total_Revenue=('Invoice Total', 'sum')
            ).reset_index()
            bu_inv.rename(columns={'Business Unit Clean': 'Business Unit'}, inplace=True)
            bu_inv['Total_Revenue'] = bu_inv['Total_Revenue'].map('${:,.2f}'.format)
            st.dataframe(bu_inv, use_container_width=True, hide_index=True)

        with col2:
            st.subheader("Invoice Breakdown by Status")
            status_inv = inv_filtered.groupby('Status').agg(
                Count=('#ID', 'count'),
                Total_Value=('Invoice Total', 'sum')
            ).reset_index()
            status_inv['Total_Value'] = status_inv['Total_Value'].map('${:,.2f}'.format)
            st.dataframe(status_inv, use_container_width=True, hide_index=True)
    else:
        st.info("Upload 'invoices.csv' in the sidebar.")

# --- TAB 5: TIMESHEETS ANALYSIS ---
with tab_ts:
    st.header("Technician Timesheets Analysis")
    if ts_df is not None and not ts_df.empty and 'User' in ts_df.columns:
        st.subheader("Technician Hours Summary")
        ts_sum = ts_df.groupby('User').agg(
            Shift_Count=('Clock In Date/Time', 'count'),
            Total_Hours=('Hours', 'sum'),
            Avg_Shift_Hours=('Hours', 'mean')
        ).reset_index()
        
        ts_sum.rename(columns={'User': 'Technician'}, inplace=True)
        ts_sum['Total_Hours'] = ts_sum['Total_Hours'].map('{:,.2f} hrs'.format)
        ts_sum['Avg_Shift_Hours'] = ts_sum['Avg_Shift_Hours'].map('{:,.2f} hrs'.format)
        
        st.dataframe(ts_sum.sort_values(by='Shift_Count', ascending=False), use_container_width=True, hide_index=True)

        st.subheader("Detailed Shift Logs")
        show_cols = [c for c in ['User', 'Clock In Date/Time', 'Clock Out Date/Time', 'Hours', 'Clock In Notes', 'Clock Out Notes'] if c in ts_df.columns]
        st.dataframe(ts_df[show_cols], use_container_width=True, hide_index=True)
    else:
        st.info("Upload 'timesheets.csv' in the sidebar.")

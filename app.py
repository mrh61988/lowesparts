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
    """Cleans parts inventory DataFrame from Google Sheets and adjusts for returns."""
    df = df.copy()
    if 'Transferred To' not in df.columns or 'Total Value' not in df.columns:
        return pd.DataFrame()
        
    if df['Total Value'].dtype == object:
        df['Total Value'] = df['Total Value'].astype(str).replace(r'[\$,]', '', regex=True)
    
    df['Total Value'] = pd.to_numeric(df['Total Value'], errors='coerce').fillna(0)
    
    if 'Qty' in df.columns:
        df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)

    # --- RETURN DEDUCTION LOGIC ---
    # If Direction is 'Return' or 'Returned', treat as negative (returned to warehouse/inventory)
    if 'Direction' in df.columns:
        is_return = df['Direction'].astype(str).str.strip().str.lower().isin(['return', 'returned', 'in'])
        df.loc[is_return, 'Total Value'] = -df.loc[is_return, 'Total Value']
        if 'Qty' in df.columns:
            df.loc[is_return, 'Qty'] = -df.loc[is_return, 'Qty']

    # Clean up Tech names
    df['Transferred To'] = df['Transferred To'].astype(str).str.replace("Matt's TransitFleet", "Matt S")
    df['Tech'] = df['Transferred To'].apply(map_tech_name)
    df['Business Unit'] = business_unit
    return df

# --- SIDEBAR FILTERS & DATA SOURCES ---
st.sidebar.header("📁 Data Sources")

sheet_url = st.sidebar.text_input(
    "Google Sheets Parts URL", 
    value="https://docs.google.com/spreadsheets/d/1OR4mEgviGglKNLwinPLnc8NB3FrN7VH9rt5qAvo8RRs/edit?usp=sharing"
)

uploaded_jobs = st.sidebar.file_uploader("Upload 'all jobs.csv'", type=['csv'])
uploaded_invoices = st.sidebar.file_uploader("Upload 'invoices.csv'", type=['csv'])
uploaded_timesheets = st.sidebar.file_uploader("Upload 'timesheets.csv'", type=['csv'])

# Target Business Units
TARGET_BUS = ['Lowes - Simple Installs', 'Lowes - Water Heaters']

# --- TABS FOR ANALYSIS ---
tab_exec, tab_parts, tab_jobs, tab_inv, tab_ts = st.tabs([
    "📈 Executive Summary Table",
    "⚙️ Parts Usage",
    "📋 Jobs Analysis",
    "💳 Invoices Analysis",
    "⏱️ Timesheets Analysis"
])

# Load Parts from Google Sheet
sheets_dict = load_google_sheet(sheet_url)
df_parts = pd.DataFrame()

if sheets_dict:
    sheet_names = list(sheets_dict.keys())
    simple_sheet = sheet_names[0] if len(sheet_names) > 0 else None
    wh_sheet = sheet_names[1] if len(sheet_names) > 1 else simple_sheet
    
    df_simple_p = process_parts_df(sheets_dict[simple_sheet], 'Lowes - Simple Installs') if simple_sheet else pd.DataFrame()
    df_wh_p = process_parts_df(sheets_dict[wh_sheet], 'Lowes - Water Heaters') if wh_sheet else pd.DataFrame()
    df_parts = pd.concat([df_simple_p, df_wh_p], ignore_index=True)

# --- TAB 1: EXECUTIVE SUMMARY TABLE ---
with tab_exec:
    st.header("Technician Level Master Summary Table")
    st.markdown("Consolidated view combining net parts cost (factoring in returns), job counts, invoice revenue, and timesheet hours.")
    
    exec_data = {}
    
    # Net Parts cost per tech
    if not df_parts.empty:
        p_agg = df_parts.groupby('Tech')['Total Value'].sum()
        for tech, val in p_agg.items():
            exec_data.setdefault(tech, {})['Net Parts Cost'] = val

    # Jobs count per tech
    if uploaded_jobs is not None:
        try:
            jobs_df = pd.read_csv(uploaded_jobs, header=1)
            if 'Business Unit' not in jobs_df.columns:
                uploaded_jobs.seek(0)
                jobs_df = pd.read_csv(uploaded_jobs)
            j_filtered = jobs_df[jobs_df['Business Unit'].isin(TARGET_BUS)]
            j_agg = j_filtered.groupby('Assigned Team Members')['#ID'].count()
            for tech, val in j_agg.items():
                exec_data.setdefault(tech, {})['Jobs Completed'] = val
        except Exception:
            pass

    # Invoice total per tech (excluding draft/void)
    if uploaded_invoices is not None:
        try:
            inv_df = pd.read_csv(uploaded_invoices, header=1)
            bu_col = 'Business Unit.1' if 'Business Unit.1' in inv_df.columns else 'Business Unit'
            if bu_col in inv_df.columns:
                inv_filtered = inv_df[
                    inv_df[bu_col].isin(TARGET_BUS) & 
                    ~inv_df['Status'].str.lower().str.contains('draft|void', na=False)
                ].copy()
                inv_filtered['Invoice Total'] = pd.to_numeric(inv_filtered['Invoice Total'], errors='coerce').fillna(0)
                i_agg = inv_filtered.groupby('Assigned Team Members')['Invoice Total'].sum()
                for tech, val in i_agg.items():
                    exec_data.setdefault(tech, {})['Total Revenue'] = val
        except Exception:
            pass

    # Timesheet hours per tech
    if uploaded_timesheets is not None:
        try:
            ts_df = pd.read_csv(uploaded_timesheets)
            ts_df['In'] = pd.to_datetime(ts_df['Clock In Date/Time'], errors='coerce')
            ts_df['Out'] = pd.to_datetime(ts_df['Clock Out Date/Time'], errors='coerce')
            ts_df['Hours'] = (ts_df['Out'] - ts_df['In']).dt.total_seconds() / 3600.0
            t_agg = ts_df.groupby('User')['Hours'].sum()
            for tech, val in t_agg.items():
                exec_data.setdefault(tech, {})['Logged Hours'] = val
        except Exception:
            pass

    if exec_data:
        master_df = pd.DataFrame.from_dict(exec_data, orient='index').fillna(0).reset_index()
        master_df.rename(columns={'index': 'Technician / Team'}, inplace=True)
        
        # Format columns
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
    st.caption("Note: Items marked with Direction = 'Return' are automatically subtracted from parts costs and quantities.")
    
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
    if uploaded_jobs is not None:
        try:
            jobs_df = pd.read_csv(uploaded_jobs, header=1)
            if 'Business Unit' not in jobs_df.columns:
                uploaded_jobs.seek(0)
                jobs_df = pd.read_csv(uploaded_jobs)
            
            jobs_filtered = jobs_df[jobs_df['Business Unit'].isin(TARGET_BUS)].copy()
            jobs_filtered['Invoice Amount'] = pd.to_numeric(jobs_filtered['Total Invoice Amount'], errors='coerce').fillna(0)

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

        except Exception as e:
            st.error(f"Error processing jobs file: {e}")
    else:
        st.info("Upload 'all jobs.csv' in the sidebar.")

# --- TAB 4: INVOICES ANALYSIS ---
with tab_inv:
    st.header("Invoices Analysis")
    st.caption("Note: Invoices with status 'Draft' or 'Void' are excluded from analysis.")
    
    if uploaded_invoices is not None:
        try:
            inv_df = pd.read_csv(uploaded_invoices, header=1)
            bu_col = 'Business Unit.1' if 'Business Unit.1' in inv_df.columns else 'Business Unit'
            
            if bu_col in inv_df.columns:
                # Filter BU and Status (Exclude Draft / Void)
                inv_filtered = inv_df[
                    inv_df[bu_col].isin(TARGET_BUS) & 
                    ~inv_df['Status'].str.lower().str.contains('draft|void', na=False)
                ].copy()
                
                inv_filtered['Invoice Total'] = pd.to_numeric(inv_filtered['Invoice Total'], errors='coerce').fillna(0)

                st.subheader("Invoices Summary by Tech, Business Unit & Parent Job Title")
                inv_sum = inv_filtered.groupby(['Assigned Team Members', bu_col, 'Parent Job Title', 'Status']).agg(
                    Invoice_Count=('#ID', 'count'),
                    Total_Amount=('Invoice Total', 'sum')
                ).reset_index()
                
                inv_sum.rename(columns={'Assigned Team Members': 'Technician', bu_col: 'Business Unit'}, inplace=True)
                inv_sum['Total_Amount'] = inv_sum['Total_Amount'].map('${:,.2f}'.format)
                
                st.dataframe(inv_sum.sort_values(by=['Technician', 'Business Unit']), use_container_width=True, hide_index=True)

                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Invoice Totals by Business Unit")
                    bu_inv = inv_filtered.groupby(bu_col).agg(
                        Invoices_Count=('#ID', 'count'),
                        Total_Revenue=('Invoice Total', 'sum')
                    ).reset_index()
                    bu_inv.rename(columns={bu_col: 'Business Unit'}, inplace=True)
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

        except Exception as e:
            st.error(f"Error processing invoices file: {e}")
    else:
        st.info("Upload 'invoices.csv' in the sidebar.")

# --- TAB 5: TIMESHEETS ANALYSIS ---
with tab_ts:
    st.header("Technician Timesheets Analysis")
    if uploaded_timesheets is not None:
        try:
            ts_df = pd.read_csv(uploaded_timesheets)
            ts_df['In'] = pd.to_datetime(ts_df['Clock In Date/Time'], errors='coerce')
            ts_df['Out'] = pd.to_datetime(ts_df['Clock Out Date/Time'], errors='coerce')
            ts_df['Hours'] = (ts_df['Out'] - ts_df['In']).dt.total_seconds() / 3600.0

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
            st.dataframe(
                ts_df[['User', 'Clock In Date/Time', 'Clock Out Date/Time', 'Hours', 'Clock In Notes', 'Clock Out Notes']], 
                use_container_width=True, 
                hide_index=True
            )

        except Exception as e:
            st.error(f"Error processing timesheets file: {e}")
    else:
        st.info("Upload 'timesheets.csv' in the sidebar.")

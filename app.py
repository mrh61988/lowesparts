import streamlit as st
import pandas as pd
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="NexSys Operations & Payroll Analyzer", layout="wide")
st.title("NexSys Operations & Financial Analyzer")
st.markdown("Detailed tabular analysis across Parts Usage, Jobs, Invoices, Timesheets, and **Replenishment Efficiency** for **Simple Installs** and **Water Heaters**.")

# --- VALID TECHNICIANS LIST & PAY STRUCTURE ---
PAY_STRUCTURE = {
    "Nate Smith": {"type": "Hourly", "rate": 22.50, "location": "Phoenix"},
    "Bill Black": {"type": "Hourly", "rate": 25.00, "location": "Phoenix"},
    "Sean Marble": {"type": "Salary", "annual": 70000.0, "location": "Phoenix"},
    "Tanner LaForge": {"type": "Hourly", "rate": 25.00, "location": "Phoenix"},
    "Erik Tange": {"type": "Commission", "rate": 0.34, "location": "Phoenix"},
    "Bryan Pickett": {"type": "Commission", "rate": 0.34, "location": "Phoenix"},
    "Matt Schlosser": {"type": "Hourly", "rate": 25.00, "location": "Phoenix"},
    "Mathew Hodges": {"type": "Salary", "annual": 65000.0, "location": "Tucson"}
}

VALID_TECHS = list(PAY_STRUCTURE.keys())

# --- HELPER FUNCTIONS ---
def normalize_single_tech(name):
    """Normalizes fleet, contractor, or user names using exact word boundaries."""
    name_str = str(name).strip()
    name_lower = name_str.lower()
    
    if re.search(r'\bbill\b', name_lower): return "Bill Black"
    if re.search(r'\bbryan\b', name_lower): return "Bryan Pickett"
    if re.search(r'\berik\b', name_lower): return "Erik Tange"
    if re.search(r'\bmatt\b', name_lower) and not re.search(r'\bmathew\b', name_lower): return "Matt Schlosser"
    if re.search(r'\bmathew\b', name_lower) or re.search(r'\bhodges\b', name_lower): return "Mathew Hodges"
    if re.search(r'\bnate\b', name_lower) or re.search(r'\bnathan\b', name_lower): return "Nate Smith"
    if re.search(r'\bsean\b', name_lower): return "Sean Marble"
    if re.search(r'\btanner\b', name_lower): return "Tanner LaForge"
    return name_str

def clean_and_filter_techs(tech_str):
    """Splits multi-tech strings, maps names, and retains only valid techs."""
    if pd.isna(tech_str) or not str(tech_str).strip():
        return None
    
    parts = [p.strip() for p in str(tech_str).split(',')]
    valid_parts = []
    
    for p in parts:
        norm = normalize_single_tech(p)
        if norm in VALID_TECHS:
            if norm not in valid_parts:
                valid_parts.append(norm)
                
    if valid_parts:
        return ", ".join(valid_parts)
    return None

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
    if 'Direction' in df.columns:
        is_return = df['Direction'].astype(str).str.strip().str.lower().isin(['return', 'returned', 'in'])
        df.loc[is_return, 'Total Value'] = -df.loc[is_return, 'Total Value']
        if 'Qty' in df.columns:
            df.loc[is_return, 'Qty'] = -df.loc[is_return, 'Qty']

    df['Transferred To'] = df['Transferred To'].astype(str).str.replace("Matt's TransitFleet", "Matt S")
    df['Tech'] = df['Transferred To'].apply(clean_and_filter_techs)
    df = df[df['Tech'].notna()].copy()
    df['Business Unit'] = business_unit
    return df

def read_uploaded_csv(file_obj):
    """Reads uploaded CSV file safely by resetting the stream pointer."""
    if file_obj is None:
        return None
    file_obj.seek(0)
    try:
        df = pd.read_csv(file_obj, header=1)
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
    jobs_filtered['Tech Clean'] = jobs_filtered['Assigned Team Members'].apply(clean_and_filter_techs)
    jobs_filtered = jobs_filtered[jobs_filtered['Tech Clean'].notna()].copy()
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
        inv_filtered['Tech Clean'] = inv_filtered['Assigned Team Members'].apply(clean_and_filter_techs)
        inv_filtered = inv_filtered[inv_filtered['Tech Clean'].notna()].copy()
        inv_filtered['Invoice Total'] = pd.to_numeric(inv_filtered['Invoice Total'], errors='coerce').fillna(0)
        inv_filtered['Business Unit Clean'] = inv_filtered[bu_col]

# 4. Timesheets Data with Weekly Overtime Logic
ts_df = read_uploaded_csv(uploaded_timesheets)
if ts_df is not None and 'Clock In Date/Time' in ts_df.columns:
    ts_df['Tech Clean'] = ts_df['User'].apply(clean_and_filter_techs)
    ts_df = ts_df[ts_df['Tech Clean'].notna()].copy()
    ts_df['In'] = pd.to_datetime(ts_df['Clock In Date/Time'], errors='coerce')
    ts_df['Out'] = pd.to_datetime(ts_df['Clock Out Date/Time'], errors='coerce')
    ts_df['Hours'] = (ts_df['Out'] - ts_df['In']).dt.total_seconds() / 3600.0
    ts_df['Week'] = ts_df['In'].dt.to_period('W-SUN')

# --- CALCULATE METRICS, REVENUE, HOURS & OVERTIME PER INDIVIDUAL TECH ---
tech_metrics = {
    t: {
        "Revenue": 0.0, 
        "Hours": 0.0, 
        "RegHours": 0.0, 
        "OTHours": 0.0, 
        "Jobs": 0, 
        "PartsCost": 0.0
    } for t in VALID_TECHS
}

if not df_parts.empty:
    for _, row in df_parts.iterrows():
        t_clean = row['Tech']
        if pd.notna(t_clean) and t_clean in tech_metrics:
            tech_metrics[t_clean]["PartsCost"] += row['Total Value']

if not inv_filtered.empty:
    for _, row in inv_filtered.iterrows():
        t_clean = row['Tech Clean']
        if pd.isna(t_clean): continue
        t_list = [x.strip() for x in t_clean.split(',')]
        split_rev = row['Invoice Total'] / len(t_list)
        for t in t_list:
            if t in tech_metrics:
                tech_metrics[t]["Revenue"] += split_rev

if not jobs_filtered.empty:
    for _, row in jobs_filtered.iterrows():
        t_clean = row['Tech Clean']
        if pd.isna(t_clean): continue
        t_list = [x.strip() for x in t_clean.split(',')]
        for t in t_list:
            if t in tech_metrics:
                tech_metrics[t]["Jobs"] += 1

if ts_df is not None and not ts_df.empty:
    weekly_hrs = ts_df.groupby(['Tech Clean', 'Week'])['Hours'].sum().reset_index()
    
    for t in VALID_TECHS:
        t_weeks = weekly_hrs[weekly_hrs['Tech Clean'] == t]
        tot_hrs = 0.0
        tot_reg = 0.0
        tot_ot = 0.0
        for _, w_row in t_weeks.iterrows():
            w_hrs = w_row['Hours']
            tot_hrs += w_hrs
            tot_reg += min(40.0, w_hrs)
            tot_ot += max(0.0, w_hrs - 40.0)
            
        tech_metrics[t]["Hours"] = tot_hrs
        tech_metrics[t]["RegHours"] = tot_reg
        tech_metrics[t]["OTHours"] = tot_ot

# --- TABS ---
tab_exec, tab_parts, tab_jobs, tab_inv, tab_ts, tab_test = st.tabs([
    "📈 Executive Summary Table",
    "⚙️ Parts Usage",
    "📋 Jobs Analysis",
    "💳 Invoices Analysis",
    "⏱️ Timesheets Analysis",
    "🧪 Test Section: BU Efficiency"
])

# --- TAB 1: EXECUTIVE SUMMARY TABLE ---
with tab_exec:
    st.header("Technician Level Master Summary Table")
    st.markdown("Consolidated view for active technicians combining net parts cost, job counts, invoice revenue, regular/overtime hours, and calculated pay.")
    
    exec_rows = []
    for t in sorted(VALID_TECHS):
        m = tech_metrics[t]
        p_info = PAY_STRUCTURE[t]
        p_type = p_info["type"]
        
        # Calculate Pay with Weekly Overtime
        if p_type == "Hourly":
            rate = p_info["rate"]
            pay = (m["RegHours"] * rate) + (m["OTHours"] * rate * 1.5)
        elif p_type == "Commission":
            pay = m["Revenue"] * p_info["rate"]
        elif p_type == "Salary":
            pay = p_info["annual"] / 12.0
            
        exec_rows.append({
            "Technician": t,
            "Jobs Completed": m["Jobs"],
            "Reg Hours": m["RegHours"],
            "OT Hours": m["OTHours"],
            "Total Hours": m["Hours"],
            "Net Parts Cost": m["PartsCost"],
            "Attributed Revenue": m["Revenue"],
            "Gross Pay (July 2026)": pay
        })

    master_df = pd.DataFrame(exec_rows)
    display_master = master_df.copy()
    display_master["Reg Hours"] = display_master["Reg Hours"].map('{:,.2f} hrs'.format)
    display_master["OT Hours"] = display_master["OT Hours"].map('{:,.2f} hrs'.format)
    display_master["Total Hours"] = display_master["Total Hours"].map('{:,.2f} hrs'.format)
    display_master["Net Parts Cost"] = display_master["Net Parts Cost"].map('${:,.2f}'.format)
    display_master["Attributed Revenue"] = display_master["Attributed Revenue"].map('${:,.2f}'.format)
    display_master["Gross Pay (July 2026)"] = display_master["Gross Pay (July 2026)"].map('${:,.2f}'.format)

    st.dataframe(display_master, use_container_width=True, hide_index=True)

# --- TAB 2: PARTS USAGE ---
with tab_parts:
    st.header("Parts Usage Analysis (Net Usage)")
    st.caption("Filtered exclusively for active technicians. Items marked as 'Return' are subtracted.")
    
    if not df_parts.empty:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Net Parts Value by Technician & Business Unit")
            t_parts = df_parts.groupby(['Tech', 'Business Unit'])['Total Value'].sum().reset_index()
            t_parts.rename(columns={'Tech': 'Technician', 'Total Value': 'Net Parts Cost'}, inplace=True)
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
        j_summary = jobs_filtered.groupby(['Tech Clean', 'Business Unit', 'Title']).agg(
            Job_Count=('#ID', 'count'),
            Total_Invoice_Amount=('Invoice Amount', 'sum')
        ).reset_index()
        j_summary.rename(columns={'Tech Clean': 'Technician', 'Title': 'Job Title'}, inplace=True)
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
    st.caption("Filtered exclusively for active technicians (Draft & Void invoices excluded).")
    
    if not inv_filtered.empty:
        st.subheader("Invoices Summary by Tech, Business Unit & Parent Job Title")
        inv_sum = inv_filtered.groupby(['Tech Clean', 'Business Unit Clean', 'Parent Job Title', 'Status']).agg(
            Invoice_Count=('#ID', 'count'),
            Total_Amount=('Invoice Total', 'sum')
        ).reset_index()
        
        inv_sum.rename(columns={'Tech Clean': 'Technician', 'Business Unit Clean': 'Business Unit'}, inplace=True)
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
    if ts_df is not None and not ts_df.empty and 'Tech Clean' in ts_df.columns:
        st.subheader("Technician Hours Summary with Weekly Overtime Breakdown")
        
        ts_weekly_summary = []
        for t in sorted(VALID_TECHS):
            m = tech_metrics[t]
            if m["Hours"] > 0:
                ts_weekly_summary.append({
                    "Technician": t,
                    "Total Hours": m["Hours"],
                    "Regular Hours (<=40/wk)": m["RegHours"],
                    "Overtime Hours (>40/wk)": m["OTHours"],
                    "Overtime % of Total": (m["OTHours"] / m["Hours"] * 100) if m["Hours"] > 0 else 0.0
                })

        ts_sum_df = pd.DataFrame(ts_weekly_summary)
        disp_ts_sum = ts_sum_df.copy()
        disp_ts_sum["Total Hours"] = disp_ts_sum["Total Hours"].map('{:,.2f} hrs'.format)
        disp_ts_sum["Regular Hours (<=40/wk)"] = disp_ts_sum["Regular Hours (<=40/wk)"].map('{:,.2f} hrs'.format)
        disp_ts_sum["Overtime Hours (>40/wk)"] = disp_ts_sum["Overtime Hours (>40/wk)"].map('{:,.2f} hrs'.format)
        disp_ts_sum["Overtime % of Total"] = disp_ts_sum["Overtime % of Total"].map('{:.1f}%'.format)
        
        st.dataframe(disp_ts_sum, use_container_width=True, hide_index=True)
    else:
        st.info("Upload 'timesheets.csv' in the sidebar.")

# --- TAB 6: TEST SECTION - BU LEVEL EFFICIENCY ---
with tab_test:
    st.header("🧪 Test Section: BU-Level Replenishment Efficiency & Material Ratios")
    st.markdown("""
    This section explicitly separates **Simple Installs** and **Water Heaters** to evaluate technician replenishment 
    intensity against expected business unit ratios. *Note: Mathew Hodges is based in Tucson and does not pull from the main warehouse.*
    """)

    def get_bu_efficiency_table(bu_name, max_material_ratio_threshold):
        bu_rows = []
        for t in sorted(VALID_TECHS):
            if not df_parts.empty:
                p_sub = df_parts[(df_parts['Tech'] == t) & (df_parts['Business Unit'] == bu_name)]
                parts_cost = p_sub['Total Value'].sum()
            else:
                parts_cost = 0.0

            j_count = 0
            if not jobs_filtered.empty:
                j_sub = jobs_filtered[jobs_filtered['Business Unit'] == bu_name]
                for _, r in j_sub.iterrows():
                    t_list = [x.strip() for x in r['Tech Clean'].split(',')]
                    if t in t_list:
                        j_count += 1

            rev = 0.0
            if not inv_filtered.empty:
                i_sub = inv_filtered[inv_filtered['Business Unit Clean'] == bu_name]
                for _, r in i_sub.iterrows():
                    t_list = [x.strip() for x in r['Tech Clean'].split(',')]
                    if t in t_list:
                        rev += r['Invoice Total'] / len(t_list)

            cost_per_job = (parts_cost / j_count) if j_count > 0 else 0.0
            mat_pct = (parts_cost / rev * 100) if rev > 0 else 0.0

            if t == "Mathew Hodges":
                flag = "🌵 Tucson Tech (No Warehouse Restocks)"
            elif j_count > 0 and parts_cost == 0:
                flag = "⚠️ Zero Parts Restocked"
            elif mat_pct > max_material_ratio_threshold:
                flag = f"🔴 High Material % (>{max_material_ratio_threshold:.0f}%)"
            elif mat_pct > 0 and mat_pct < 1.0 and j_count > 5:
                flag = "🟡 Low Material % (Unreported Transfers?)"
            elif j_count == 0 and parts_cost == 0:
                flag = "⚪ No Jobs in BU"
            else:
                flag = "🟢 Normal Range"

            if j_count > 0 or parts_cost > 0:
                bu_rows.append({
                    "Technician": t,
                    "Jobs Completed": j_count,
                    "Attributed Revenue": rev,
                    "Net Replenishment Cost": parts_cost,
                    "Replenishment / Job": cost_per_job,
                    "Material % of Revenue": mat_pct,
                    "Operational Flag": flag
                })

        df_res = pd.DataFrame(bu_rows)
        if not df_res.empty:
            df_res["Attributed Revenue"] = df_res["Attributed Revenue"].map('${:,.2f}'.format)
            df_res["Net Replenishment Cost"] = df_res["Net Replenishment Cost"].map('${:,.2f}'.format)
            df_res["Replenishment / Job"] = df_res["Replenishment / Job"].map('${:,.2f}'.format)
            df_res["Material % of Revenue"] = df_res["Material % of Revenue"].map('{:.2f}%'.format)
        return df_res

    st.subheader("1. Lowes - Simple Installs (Expected Material Ratio: 2.0% – 8.0%)")
    simple_eff_df = get_bu_efficiency_table('Lowes - Simple Installs', max_material_ratio_threshold=8.0)
    if not simple_eff_df.empty:
        st.dataframe(simple_eff_df, use_container_width=True, hide_index=True)
    else:
        st.info("No data available for Simple Installs.")

    st.subheader("2. Lowes - Water Heaters (Expected Material Ratio: 2.5% – 12.0%)")
    wh_eff_df = get_bu_efficiency_table('Lowes - Water Heaters', max_material_ratio_threshold=12.0)
    if not wh_eff_df.empty:
        st.dataframe(wh_eff_df, use_container_width=True, hide_index=True)
    else:
        st.info("No data available for Water Heaters.")

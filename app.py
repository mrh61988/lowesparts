import streamlit as st
import pandas as pd
import numpy as np
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="NexSys Operations & Payroll Analyzer", layout="wide")

# --- INJECT CUSTOM CSS FOR TEXT WRAPPING & ZERO HORIZONTAL SCROLLING ---
st.markdown("""
<style>
/* Force text wrapping and fit container for Streamlit tables and dataframes */
div[data-testid="stDataFrame"] td, div[data-testid="stDataFrame"] th {
    white-space: normal !important;
    word-wrap: break-word !important;
}
.stTable td, .stTable th {
    white-space: normal !important;
    word-wrap: break-word !important;
}
/* Ensure dataframe container uses full width cleanly */
div[data-testid="stDataFrame"] {
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

st.title("NexSys Operations & Financial Analyzer")
st.markdown("Detailed tabular analysis across Parts Usage, Jobs, Invoices, Timesheets, Warehouse Inventory Targets, and **Replenishment Efficiency** for **Simple Installs** and **Water Heaters**.")

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
def find_col(df, keyword_lists):
    """Fuzzy column matcher handling hidden line breaks, newlines, and alternate names in Google Sheets."""
    for keywords in keyword_lists:
        for c in df.columns:
            c_clean = str(c).lower().replace('\n', ' ').strip()
            if all(k.lower() in c_clean for k in keywords):
                return c
    return None

def clean_sku(val):
    """Robust SKU cleaner handling integers, floats like 40699.0, strings, and whitespace."""
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

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

def get_first_valid_tech(tech_str):
    """Attributes 100% of credit to the FIRST listed valid active technician."""
    if pd.isna(tech_str) or not str(tech_str).strip():
        return None
    
    parts = [p.strip() for p in str(tech_str).split(',')]
    for p in parts:
        norm = normalize_single_tech(p)
        if norm in VALID_TECHS:
            return norm
    return None

def parse_min_max(val):
    """Parses warehouse min/max formats (e.g. '70 / 140', '4-8', '4 8')."""
    val_str = str(val).strip()
    if '/' in val_str:
        parts = [p.strip() for p in val_str.split('/')]
        c_min = parts[0] if parts[0] not in ['—', '-', '', 'nan', 'None'] else '0'
        c_max = parts[1] if len(parts) > 1 and parts[1] not in ['—', '-', '', 'nan', 'None'] else '0'
        try:
            return int(float(c_min)), int(float(c_max))
        except:
            pass
            
    nums = re.findall(r'\d+', val_str)
    if len(nums) >= 2:
        return int(nums[0]), int(nums[1])
    elif len(nums) == 1:
        return int(nums[0]), 0
        
    return 0, 0

def map_dept_to_bu(dept):
    """Standardizes Department string from Nexsys Min/Max sheet to Business Unit."""
    dept_str = str(dept).lower().strip()
    if 'simple' in dept_str:
        return 'Lowes - Simple Installs'
    if 'water' in dept_str or 'heater' in dept_str:
        return 'Lowes - Water Heaters'
    return 'Other'

@st.cache_data(ttl=15)
def fetch_live_google_sheet(url):
    """Reads a public Google Sheet URL into a dictionary of DataFrames."""
    if not url:
        return None
    base_url = url.split('/edit')[0]
    export_url = f"{base_url}/export?format=xlsx"
    try:
        xls = pd.read_excel(export_url, sheet_name=None)
        return xls
    except Exception as e:
        st.error(f"Error reading Google Sheet. Ensure link permissions are public. Details: {e}")
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

    if 'SKU' in df.columns:
        df['SKU'] = df['SKU'].apply(clean_sku)

    df['Transferred To'] = df['Transferred To'].astype(str).str.replace("Matt's TransitFleet", "Matt S")
    df['Tech'] = df['Transferred To'].apply(get_first_valid_tech)
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

def detect_dataset_date_range(ts_df, inv_filtered, df_parts):
    """Automatically detects overall date range across timesheets, invoices, and parts transfers."""
    all_dates = []

    if ts_df is not None and not ts_df.empty and 'In' in ts_df.columns:
        valid_ts_dates = ts_df['In'].dropna()
        if not valid_ts_dates.empty:
            all_dates.extend(valid_ts_dates.tolist())

    if inv_filtered is not None and not inv_filtered.empty:
        for col in inv_filtered.columns:
            if 'date' in str(col).lower():
                parsed = pd.to_datetime(inv_filtered[col], errors='coerce').dropna()
                if not parsed.empty:
                    all_dates.extend(parsed.tolist())
                    break

    if df_parts is not None and not df_parts.empty:
        for col in df_parts.columns:
            if 'date' in str(col).lower():
                parsed = pd.to_datetime(df_parts[col], errors='coerce').dropna()
                if not parsed.empty:
                    all_dates.extend(parsed.tolist())
                    break

    if all_dates:
        min_d = min(all_dates)
        max_d = max(all_dates)
        days = (max_d - min_d).days + 1
        days = max(1, days)
        weeks = days / 7.0

        if days <= 10:
            label = f"{days}-Day Period ({min_d.strftime('%b %d')}–{max_d.strftime('%b %d')})"
        elif days <= 18:
            label = f"2-Week Period ({min_d.strftime('%b %d')}–{max_d.strftime('%b %d')})"
        else:
            label = f"Period ({min_d.strftime('%b %d')}–{max_d.strftime('%b %d')})"

        return days, weeks, label, min_d, max_d

    return 31, 31.0 / 7.0, "Monthly (Default)", None, None

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
sheets_dict = fetch_live_google_sheet(sheet_url)
df_parts = pd.DataFrame()
df_current_minmax = pd.DataFrame()

if sheets_dict:
    sheet_names = list(sheets_dict.keys())
    simple_sheet, wh_parts_sheet, wh_units_sheet, minmax_sheet = None, None, None, None
    
    for s in sheet_names:
        sl = s.lower()
        if 'min' in sl and 'max' in sl:
            minmax_sheet = s
        elif 'simple' in sl:
            simple_sheet = s
        elif 'warehouse' in sl and 'water heater' in sl:
            wh_units_sheet = s
        elif 'water heater' in sl and 'part' in sl:
            wh_parts_sheet = s
            
    if not simple_sheet and len(sheet_names) > 0: simple_sheet = sheet_names[0]
    if not wh_parts_sheet and len(sheet_names) > 1: wh_parts_sheet = sheet_names[1]
    if not wh_units_sheet and len(sheet_names) > 2: wh_units_sheet = sheet_names[2]
    
    if simple_sheet == minmax_sheet: simple_sheet = None
    if wh_parts_sheet == minmax_sheet: wh_parts_sheet = None
    if wh_units_sheet == minmax_sheet: wh_units_sheet = None
    
    df_simple_p = process_parts_df(sheets_dict[simple_sheet], 'Lowes - Simple Installs') if simple_sheet else pd.DataFrame()
    df_wh_p = process_parts_df(sheets_dict[wh_parts_sheet], 'Lowes - Water Heaters (Parts)') if wh_parts_sheet else pd.DataFrame()
    df_wh_u = process_parts_df(sheets_dict[wh_units_sheet], 'Lowes - Water Heaters (Units)') if wh_units_sheet else pd.DataFrame()
    
    df_parts = pd.concat([df_simple_p, df_wh_p, df_wh_u], ignore_index=True)
    
    if minmax_sheet and minmax_sheet in sheets_dict:
        raw_minmax = sheets_dict[minmax_sheet].copy()
        raw_minmax.columns = raw_minmax.columns.astype(str).str.replace('\n', ' ').str.strip()
        
        c_sku = find_col(raw_minmax, [['sku'], ['item #'], ['part']])
        c_minmax = find_col(raw_minmax, [['warehouse', 'min'], ['min', 'max'], ['min/max']])
        c_qty = find_col(raw_minmax, [['qty'], ['quantity'], ['on hand'], ['stock']])
        c_dept = find_col(raw_minmax, [['department'], ['business unit'], ['bu']])
        c_item = find_col(raw_minmax, [['item name'], ['description'], ['item']])

        if c_sku:
            raw_minmax['SKU_clean'] = raw_minmax[c_sku].apply(clean_sku)
            if c_minmax:
                parsed_mins_maxs = raw_minmax[c_minmax].apply(parse_min_max).tolist()
                raw_minmax[['Current Min', 'Current Max']] = parsed_mins_maxs
            else:
                raw_minmax['Current Min'] = 0
                raw_minmax['Current Max'] = 0
                
            raw_minmax['Current On Hand'] = pd.to_numeric(raw_minmax[c_qty], errors='coerce').fillna(0).astype(int) if c_qty else 0
            raw_minmax['Business Unit_sheet'] = raw_minmax[c_dept].apply(map_dept_to_bu) if c_dept else 'Unknown'
            raw_minmax['Item Name'] = raw_minmax[c_item] if c_item else ''
            
            df_current_minmax = raw_minmax[['SKU_clean', 'Item Name', 'Business Unit_sheet', 'Current On Hand', 'Current Min', 'Current Max']].copy()

# 2. Jobs Data
raw_jobs_df = read_uploaded_csv(uploaded_jobs)
jobs_filtered = pd.DataFrame()
if raw_jobs_df is not None and 'Business Unit' in raw_jobs_df.columns:
    jobs_filtered = raw_jobs_df[raw_jobs_df['Business Unit'].isin(TARGET_BUS)].copy()
    jobs_filtered['Tech Clean'] = jobs_filtered['Assigned Team Members'].apply(get_first_valid_tech)
    jobs_filtered = jobs_filtered[jobs_filtered['Tech Clean'].notna()].copy()
    jobs_filtered['Invoice Amount'] = pd.to_numeric(jobs_filtered['Total Invoice Amount'], errors='coerce').fillna(0)

# 3. Invoices Data
raw_inv_df = read_uploaded_csv(uploaded_invoices)
inv_filtered = pd.DataFrame()
if raw_inv_df is not None:
    bu_col = 'Business Unit.1' if 'Business Unit.1' in raw_inv_df.columns else ('Business Unit' if 'Business Unit' in raw_inv_df.columns else None)
    if bu_col:
        inv_filtered = raw_inv_df[
            raw_inv_df[bu_col].isin(TARGET_BUS) & 
            ~raw_inv_df['Status'].astype(str).str.lower().str.contains('draft|void', na=False)
        ].copy()
        inv_filtered['Tech Clean'] = inv_filtered['Assigned Team Members'].apply(get_first_valid_tech)
        inv_filtered = inv_filtered[inv_filtered['Tech Clean'].notna()].copy()
        inv_filtered['Invoice Total'] = pd.to_numeric(inv_filtered['Invoice Total'], errors='coerce').fillna(0)
        inv_filtered['Business Unit Clean'] = inv_filtered[bu_col]

# 4. Timesheets Data
ts_df = read_uploaded_csv(uploaded_timesheets)
if ts_df is not None and 'Clock In Date/Time' in ts_df.columns:
    ts_df['Tech Clean'] = ts_df['User'].apply(get_first_valid_tech)
    ts_df = ts_df[ts_df['Tech Clean'].notna()].copy()
    ts_df['In'] = pd.to_datetime(ts_df['Clock In Date/Time'], errors='coerce')
    ts_df['Out'] = pd.to_datetime(ts_df['Clock Out Date/Time'], errors='coerce')
    ts_df['Hours'] = (ts_df['Out'] - ts_df['In']).dt.total_seconds() / 3600.0
    ts_df['Week'] = ts_df['In'].dt.to_period('W-SUN')

# --- AUTOMATIC DATE RANGE DETECTION ---
total_days, total_weeks, period_label, min_date, max_date = detect_dataset_date_range(ts_df, inv_filtered, df_parts)

# Sidebar Feedback Banner
st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Auto-Detected Timeframe")
if min_date and max_date:
    st.sidebar.success(
        f"**Range:** {min_date.strftime('%b %d, %Y')} – {max_date.strftime('%b %d, %Y')}\n\n"
        f"**Duration:** {total_days} Days ({total_weeks:.2f} Weeks)"
    )
else:
    st.sidebar.info("Using default monthly baseline until data files with dates are uploaded.")

# Display active tech roster in sidebar
st.sidebar.markdown("---")
st.sidebar.markdown("### 👷 Tech Roster")
for t, p in PAY_STRUCTURE.items():
    loc_tag = "🌵 Tucson" if p["location"] == "Tucson" else "📍 Phoenix"
    st.sidebar.markdown(f"**{t}** ({loc_tag})")

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
        if pd.notna(t_clean) and t_clean in tech_metrics:
            tech_metrics[t_clean]["Revenue"] += row['Invoice Total']

if not jobs_filtered.empty:
    for _, row in jobs_filtered.iterrows():
        t_clean = row['Tech Clean']
        if pd.notna(t_clean) and t_clean in tech_metrics:
            tech_metrics[t_clean]["Jobs"] += 1

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
tab_exec, tab_parts, tab_minmax, tab_jobs, tab_inv, tab_ts, tab_test = st.tabs([
    "📈 Executive Summary",
    "⚙️ Parts Usage",
    "📦 Warehouse Min/Max Targets",
    "📋 Jobs Analysis",
    "💳 Invoices Analysis",
    "⏱️ Timesheets Analysis",
    "🧪 Test Section: BU & Item Efficiency"
])

# --- TAB 1: EXECUTIVE SUMMARY TABLE ---
with tab_exec:
    st.header("Technician Level Master Summary Table")
    st.markdown("Consolidated view for active technicians combining net parts cost, job counts, invoice revenue, regular/overtime hours, calculated pay, and gross pay % of revenue.")
    
    exec_rows = []
    for t in sorted(VALID_TECHS):
        m = tech_metrics[t]
        p_info = PAY_STRUCTURE[t]
        p_type = p_info["type"]
        
        # Calculate Pay with Dynamic Timeframe Scaling for Salaries
        if p_type == "Hourly":
            rate = p_info["rate"]
            pay = (m["RegHours"] * rate) + (m["OTHours"] * rate * 1.5)
        elif p_type == "Commission":
            pay = m["Revenue"] * p_info["rate"]
        elif p_type == "Salary":
            pay = p_info["annual"] * (total_days / 365.0)
            
        rev = m["Revenue"]
        pay_pct = (pay / rev * 100.0) if rev > 0 else 0.0

        exec_rows.append({
            "Technician": t,
            "Jobs Completed": m["Jobs"],
            "Reg Hours": m["RegHours"],
            "OT Hours": m["OTHours"],
            "Total Hours": m["Hours"],
            "Net Parts Cost": m["PartsCost"],
            "Attributed Revenue": rev,
            f"Gross Pay ({period_label})": pay,
            "Gross Pay % of Rev": pay_pct
        })

    master_df = pd.DataFrame(exec_rows)
    display_master = master_df.copy()
    display_master["Reg Hours"] = display_master["Reg Hours"].map('{:,.2f} hrs'.format)
    display_master["OT Hours"] = display_master["OT Hours"].map('{:,.2f} hrs'.format)
    display_master["Total Hours"] = display_master["Total Hours"].map('{:,.2f} hrs'.format)
    display_master["Net Parts Cost"] = display_master["Net Parts Cost"].map('${:,.2f}'.format)
    display_master["Attributed Revenue"] = display_master["Attributed Revenue"].map('${:,.2f}'.format)
    display_master[f"Gross Pay ({period_label})"] = display_master[f"Gross Pay ({period_label})"].map('${:,.2f}'.format)
    display_master["Gross Pay % of Rev"] = display_master["Gross Pay % of Rev"].map('{:.2f}%'.format)

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

# --- TAB 3: WAREHOUSE MIN/MAX COMPARISON ---
with tab_minmax:
    st.header("📦 Warehouse Min/Max Analysis & Comparison")
    st.markdown(f"""
    Comparison of **Current Warehouse Min/Max settings** against **Suggested 1.5-Week Inventory Targets** calculated from auto-detected historical demand (**{period_label}** / **{total_weeks:.2f} weeks**).
    
    *Only triggers an action if the suggested Min or Max deviates by **20% or more** from the current setting in Google Sheets. Organized with the highest priority items at the top.*
    """)
    
    if not df_parts.empty:
        item_usage = df_parts.groupby(['Business Unit', 'SKU', 'Item']).agg(
            Total_Net_Qty=('Qty', 'sum'),
            Total_Net_Cost=('Total Value', 'sum')
        ).reset_index()
        
        item_usage['SKU_clean'] = item_usage['SKU'].apply(clean_sku)
        item_usage['Weekly_Avg_Qty'] = item_usage['Total_Net_Qty'] / total_weeks
        item_usage['Min_Stock_Qty'] = np.ceil(item_usage['Weekly_Avg_Qty'] * 1.0).clip(lower=1).astype(int)
        item_usage['Target_Stock_Qty'] = np.ceil(item_usage['Weekly_Avg_Qty'] * 1.5).astype(int)
        item_usage['Max_Stock_Qty'] = np.maximum(np.ceil(item_usage['Weekly_Avg_Qty'] * 2.0), item_usage['Min_Stock_Qty'] + 1).astype(int)
        
        if not df_current_minmax.empty:
            merged_minmax = pd.merge(
                item_usage[['SKU_clean', 'Business Unit', 'Item', 'Total_Net_Qty', 'Weekly_Avg_Qty', 'Min_Stock_Qty', 'Target_Stock_Qty', 'Max_Stock_Qty']],
                df_current_minmax, 
                on=['SKU_clean'], 
                how='outer'
            )
            
            def resolve_wh_bu(row):
                bu = row.get('Business Unit')
                if pd.notna(bu) and str(bu).strip() != '' and bu != 'Unknown':
                    return bu
                return row.get('Business Unit_sheet', 'Unknown')

            merged_minmax['Business Unit'] = merged_minmax.apply(resolve_wh_bu, axis=1)
            
            item_name_col = merged_minmax['Item Name'] if 'Item Name' in merged_minmax.columns else pd.Series(index=merged_minmax.index)
            item_col = merged_minmax['Item'] if 'Item' in merged_minmax.columns else pd.Series(index=merged_minmax.index)
            
            merged_minmax['Item Description'] = item_name_col.fillna(item_col).fillna('')
            merged_minmax['Total_Net_Qty'] = merged_minmax['Total_Net_Qty'].fillna(0).astype(int)
            merged_minmax['Weekly_Avg_Qty'] = merged_minmax['Weekly_Avg_Qty'].fillna(0.0)
            merged_minmax['Current On Hand'] = merged_minmax['Current On Hand'].fillna(0).astype(int)
            merged_minmax['Current Min'] = merged_minmax['Current Min'].fillna(0).astype(int)
            merged_minmax['Current Max'] = merged_minmax['Current Max'].fillna(0).astype(int)
            merged_minmax['Min_Stock_Qty'] = merged_minmax['Min_Stock_Qty'].fillna(0).astype(int)
            merged_minmax['Target_Stock_Qty'] = merged_minmax['Target_Stock_Qty'].fillna(0).astype(int)
            merged_minmax['Max_Stock_Qty'] = merged_minmax['Max_Stock_Qty'].fillna(0).astype(int)
            merged_minmax['SKU'] = merged_minmax['SKU_clean']
            
        else:
            merged_minmax = item_usage.copy()
            merged_minmax.rename(columns={'Item': 'Item Description'}, inplace=True)
            merged_minmax['Current On Hand'] = 0
            merged_minmax['Current Min'] = 0
            merged_minmax['Current Max'] = 0

        def resolve_minmax_and_rec(row):
            c_min, c_max = int(row['Current Min']), int(row['Current Max'])
            s_min, s_max = int(row['Min_Stock_Qty']), int(row['Max_Stock_Qty'])
            target_qty = int(row['Target_Stock_Qty'])
            j_net = int(row['Total_Net_Qty'])
            
            if c_min == 0 and c_max == 0:
                if j_net == 0:
                    return 0, 0, 0, "⚪ Inactive / No Stock"
                return s_min, s_max, target_qty, "⚠️ Set Min/Max"
                
            if j_net == 0:
                s_min_adj = c_min
                s_max_adj = c_max
                target_adj = int(np.ceil(c_min * 1.5)) if c_min > 0 else c_max
                return s_min_adj, s_max_adj, target_adj, "⚪ Zero Demand in Period"
                
            d_min = s_min - c_min
            d_max = s_max - c_max
            
            min_flag = abs(d_min) >= (0.20 * c_min) if c_min > 0 else (s_min > 0)
            max_flag = abs(d_max) >= (0.20 * c_max) if c_max > 0 else (s_max > 0)
            
            if not min_flag and not max_flag:
                return s_min, s_max, target_qty, "🟢 On Target"
                
            rec = []
            if min_flag:
                rec.append(f"Inc Min (+{d_min})" if d_min > 0 else f"Dec Min ({d_min})")
            if max_flag:
                rec.append(f"Inc Max (+{d_max})" if d_max > 0 else f"Dec Max ({d_max})")
                
            if len(rec) == 2:
                if d_min > 0 and d_max > 0:
                    rec_str = f"⬆️ {rec[0]} & {rec[1].replace('Inc ', '')}"
                elif d_min < 0 and d_max < 0:
                    rec_str = f"⬇️ {rec[0]} & {rec[1].replace('Dec ', '')}"
                else:
                    rec_str = f"🔄 {rec[0]} & {rec[1]}"
            else:
                rec_str = f"⬆️ {rec[0]}" if "Inc" in rec[0] else f"⬇️ {rec[0]}"
                
            return s_min, s_max, target_qty, rec_str

        res = merged_minmax.apply(resolve_minmax_and_rec, axis=1)
        merged_minmax['Min_Stock_Qty_Adj'] = [r[0] for r in res]
        merged_minmax['Max_Stock_Qty_Adj'] = [r[1] for r in res]
        merged_minmax['Target_Stock_Qty_Adj'] = [r[2] for r in res]
        merged_minmax['Action / Rec'] = [r[3] for r in res]

        merged_minmax['Min (Curr ➔ Sug)'] = merged_minmax['Current Min'].astype(str) + " ➔ " + merged_minmax['Min_Stock_Qty_Adj'].astype(str)
        merged_minmax['Max (Curr ➔ Sug)'] = merged_minmax['Current Max'].astype(str) + " ➔ " + merged_minmax['Max_Stock_Qty_Adj'].astype(str)

        def get_sort_priority(action_str):
            if "⚠️" in action_str: return 1
            if "⬆️" in action_str or "⬇️" in action_str or "🔄" in action_str: return 2
            if "🟢" in action_str: return 3
            if "⚪ Zero Demand" in action_str: return 4
            return 5

        merged_minmax['Sort_Priority'] = merged_minmax['Action / Rec'].apply(get_sort_priority)

        def render_comparison_table(bu_name):
            bu_df = merged_minmax[merged_minmax['Business Unit'].str.contains(bu_name, case=False, na=False)].copy()
            if not bu_df.empty:
                bu_df.sort_values(by=['Sort_Priority', 'Target_Stock_Qty_Adj'], ascending=[True, False], inplace=True)
                
                bu_df.rename(columns={
                    'SKU': 'SKU',
                    'Total_Net_Qty': f'Net ({period_label})',
                    'Weekly_Avg_Qty': 'Wk Avg',
                    'Current On Hand': 'On Hand',
                    'Target_Stock_Qty_Adj': 'Target (1.5 Wk)'
                }, inplace=True)
                
                bu_df['Wk Avg'] = bu_df['Wk Avg'].map('{:.2f}'.format)
                
                show_cols = [
                    'SKU', 'Item Description', f'Net ({period_label})', 'Wk Avg', 
                    'On Hand', 'Min (Curr ➔ Sug)', 'Max (Curr ➔ Sug)', 'Target (1.5 Wk)', 
                    'Action / Rec'
                ]
                
                st.dataframe(
                    bu_df[show_cols], 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "Item Description": st.column_config.TextColumn("Item Description", width="medium"),
                        "Action / Rec": st.column_config.TextColumn("Action / Rec", width="medium")
                    }
                )
            else:
                st.info(f"No parts usage data available for {bu_name}.")

        st.subheader("1. Lowes - Simple Installs Min/Max Comparison")
        render_comparison_table('Simple Installs')
        
        st.subheader("2. Lowes - Water Heaters Min/Max Comparison")
        render_comparison_table('Water Heaters')

    else:
        st.info("Google Sheet parts data not loaded.")

# --- TAB 4: JOBS ANALYSIS ---
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

# --- TAB 5: INVOICES ANALYSIS ---
with tab_inv:
    st.header("Invoices Analysis")
    st.caption("Filtered exclusively for active technicians (Draft & Void invoices excluded). Full credit attributed to the first listed tech.")
    
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

# --- TAB 6: TIMESHEETS ANALYSIS ---
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

# --- TAB 7: TEST SECTION - BU & ITEM LEVEL EFFICIENCY ---
with tab_test:
    st.header("🧪 Test Section: Item-Level Usage, Revenue & Efficiency Analysis")
    st.markdown("""
    Test environment for evaluating item-level parts consumption linked directly to technician revenue generation, alongside business unit replenishment efficiency.
    """)

    st.markdown("---")
    # --- TEST TABLE 1: Side-by-Side Technician Comparison Drilldown ---
    st.subheader("1. Side-by-Side Technician Item Usage & Revenue Comparison (Excludes AO Smith Water Heaters)")
    
    selected_bu_label = st.selectbox(
        "Filter Category / Business Unit:",
        ["All (Simple Installs & WH Parts)", "Simple Installs", "Water Heater Parts"]
    )

    col_tech_a, col_tech_b = st.columns(2)

    def render_tech_drilldown(tech_name, bu_filter_label, col_container):
        with col_container:
            st.markdown(f"### 👤 {tech_name}")
            if not df_parts.empty:
                is_ao_smith = df_parts['Item'].astype(str).str.lower().str.contains(r'a\.?o\.?\s*smith', regex=True, na=False)
                df_parts_clean = df_parts[~is_ao_smith].copy()

                if bu_filter_label == "Simple Installs":
                    p_tech = df_parts_clean[(df_parts_clean['Tech'] == tech_name) & (df_parts_clean['Business Unit'].str.contains('Simple Installs', case=False, na=False))]
                    tech_rev = inv_filtered[(inv_filtered['Tech Clean'] == tech_name) & (inv_filtered['Business Unit Clean'].str.contains('Simple Installs', case=False, na=False))]['Invoice Total'].sum() if not inv_filtered.empty else 0.0
                    tech_jobs = len(jobs_filtered[(jobs_filtered['Tech Clean'] == tech_name) & (jobs_filtered['Business Unit'].str.contains('Simple Installs', case=False, na=False))]) if not jobs_filtered.empty else 0

                elif bu_filter_label == "Water Heater Parts":
                    p_tech = df_parts_clean[(df_parts_clean['Tech'] == tech_name) & (df_parts_clean['Business Unit'].str.contains('Water Heater', case=False, na=False)) & (~df_parts_clean['Business Unit'].str.contains('Units', case=False, na=False))]
                    tech_rev = inv_filtered[(inv_filtered['Tech Clean'] == tech_name) & (inv_filtered['Business Unit Clean'].str.contains('Water Heater', case=False, na=False))]['Invoice Total'].sum() if not inv_filtered.empty else 0.0
                    tech_jobs = len(jobs_filtered[(jobs_filtered['Tech Clean'] == tech_name) & (jobs_filtered['Business Unit'].str.contains('Water Heater', case=False, na=False))]) if not jobs_filtered.empty else 0

                else: # All (Simple Installs & WH Parts)
                    p_tech = df_parts_clean[(df_parts_clean['Tech'] == tech_name) & (~df_parts_clean['Business Unit'].str.contains('Units', case=False, na=False))]
                    tech_rev = tech_metrics[tech_name]['Revenue']
                    tech_jobs = tech_metrics[tech_name]['Jobs']

                tech_parts_total = p_tech['Total Value'].sum()
                mat_pct = (tech_parts_total / tech_rev * 100) if tech_rev > 0 else 0.0

                # Prominent 2x2 Metric Grid
                m1, m2 = st.columns(2)
                m1.metric("📋 Jobs Completed", f"{tech_jobs}")
                m2.metric("💰 Attributed Revenue", f"${tech_rev:,.2f}")
                
                m3, m4 = st.columns(2)
                m3.metric("⚙️ Net Parts Cost", f"${tech_parts_total:,.2f}")
                m4.metric("📊 Material Ratio", f"{mat_pct:.2f}%")

                if not p_tech.empty:
                    tech_item_summary = p_tech.groupby(['SKU', 'Item']).agg(
                        Qty_Used=('Qty', 'sum'),
                        Total_Cost=('Total Value', 'sum')
                    ).reset_index().sort_values(by='Total_Cost', ascending=False)

                    tech_item_summary['Cost % Rev'] = (
                        (tech_item_summary['Total_Cost'] / tech_rev * 100).map('{:.2f}%'.format)
                        if tech_rev > 0 else "0.00%"
                    )
                    tech_item_summary['Cost'] = tech_item_summary['Total_Cost'].map('${:,.2f}'.format)
                    tech_item_summary.rename(columns={'Qty_Used': 'Qty'}, inplace=True)

                    st.dataframe(
                        tech_item_summary[['SKU', 'Item', 'Qty', 'Cost', 'Cost % Rev']], 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "SKU": st.column_config.TextColumn("SKU", width="small"),
                            "Item": st.column_config.TextColumn("Item Description", width="medium"),
                            "Qty": st.column_config.NumberColumn("Qty", width="small"),
                            "Cost": st.column_config.TextColumn("Cost", width="small"),
                            "Cost % Rev": st.column_config.TextColumn("% Rev", width="small")
                        }
                    )
                else:
                    st.info(f"No parts usage logged for {tech_name} under {bu_filter_label}.")
            else:
                st.info("Parts usage data not loaded.")

    with col_tech_a:
        tech_a = st.selectbox("Select Primary Technician (Tech A):", sorted(VALID_TECHS), index=sorted(VALID_TECHS).index("Matt Schlosser") if "Matt Schlosser" in VALID_TECHS else 0)
        render_tech_drilldown(tech_a, selected_bu_label, col_tech_a)

    with col_tech_b:
        default_b_idx = sorted(VALID_TECHS).index("Erik Tange") if "Erik Tange" in VALID_TECHS else (1 if len(VALID_TECHS) > 1 else 0)
        tech_b = st.selectbox("Select Comparison Technician (Tech B):", sorted(VALID_TECHS), index=default_b_idx)
        render_tech_drilldown(tech_b, selected_bu_label, col_tech_b)

    st.markdown("---")
    # --- TEST TABLE 2: Top Consumed Items per Technician Summary (Excludes AO Smith Water Heaters) ---
    st.subheader("2. Top 3 Consumed Items per Technician (by Value, Excl. AO Smith)")
    
    bu_filter_sec2 = st.selectbox(
        "Filter Category / Business Unit:",
        ["All (Simple Installs & WH Parts)", "Simple Installs", "Water Heater Parts"],
        key="bu_filter_sec2"
    )

    if not df_parts.empty:
        is_ao_smith = df_parts['Item'].astype(str).str.lower().str.contains(r'a\.?o\.?\s*smith', regex=True, na=False)
        df_parts_clean = df_parts[~is_ao_smith]

        top_skus_list = []
        for t in sorted(VALID_TECHS):
            if bu_filter_sec2 == "Simple Installs":
                p_sub = df_parts_clean[(df_parts_clean['Tech'] == t) & (df_parts_clean['Business Unit'].str.contains('Simple Installs', case=False, na=False))]
                t_rev = inv_filtered[(inv_filtered['Tech Clean'] == t) & (inv_filtered['Business Unit Clean'].str.contains('Simple Installs', case=False, na=False))]['Invoice Total'].sum() if not inv_filtered.empty else 0.0
                t_jobs = len(jobs_filtered[(jobs_filtered['Tech Clean'] == t) & (jobs_filtered['Business Unit'].str.contains('Simple Installs', case=False, na=False))]) if not jobs_filtered.empty else 0

            elif bu_filter_sec2 == "Water Heater Parts":
                p_sub = df_parts_clean[(df_parts_clean['Tech'] == t) & (df_parts_clean['Business Unit'].str.contains('Water Heater', case=False, na=False)) & (~df_parts_clean['Business Unit'].str.contains('Units', case=False, na=False))]
                t_rev = inv_filtered[(inv_filtered['Tech Clean'] == t) & (inv_filtered['Business Unit Clean'].str.contains('Water Heater', case=False, na=False))]['Invoice Total'].sum() if not inv_filtered.empty else 0.0
                t_jobs = len(jobs_filtered[(jobs_filtered['Tech Clean'] == t) & (jobs_filtered['Business Unit'].str.contains('Water Heater', case=False, na=False))]) if not jobs_filtered.empty else 0

            else: # All (Simple Installs & WH Parts)
                p_sub = df_parts_clean[(df_parts_clean['Tech'] == t) & (~df_parts_clean['Business Unit'].str.contains('Units', case=False, na=False))]
                t_rev = tech_metrics[t]['Revenue']
                t_jobs = tech_metrics[t]['Jobs']

            if not p_sub.empty:
                top_items = (
                    p_sub.groupby('Item')['Total Value']
                    .sum()
                    .reset_index()
                    .sort_values(by='Total Value', ascending=False)
                    .head(3)
                )
                top_str = ", ".join([f"{r['Item']} (${r['Total Value']:,.2f})" for _, r in top_items.iterrows()])
            else:
                top_str = "None"

            t_cost = p_sub['Total Value'].sum()

            top_skus_list.append({
                "Technician": t,
                "Jobs Completed": t_jobs,
                "Attributed Revenue": t_rev,
                "Parts Cost (Excl. AO Smith)": t_cost,
                "Material % of Rev": (t_cost / t_rev * 100) if t_rev > 0 else 0.0,
                "Top 3 Consumed Items": top_str
            })

        df_top_summary = pd.DataFrame(top_skus_list)
        df_top_summary["Attributed Revenue"] = df_top_summary["Attributed Revenue"].map('${:,.2f}'.format)
        df_top_summary["Parts Cost (Excl. AO Smith)"] = df_top_summary["Parts Cost (Excl. AO Smith)"].map('${:,.2f}'.format)
        df_top_summary["Material % of Rev"] = df_top_summary["Material % of Rev"].map('{:.2f}%'.format)

        st.dataframe(df_top_summary, use_container_width=True, hide_index=True)

    st.markdown("---")
    # --- TEST TABLE 3: SKU Benchmarking Across Technicians (Excludes AO Smith Water Heaters) ---
    st.subheader("3. SKU Usage Benchmarking Matrix Across Technicians (Excl. AO Smith)")
    if not df_parts.empty:
        is_ao_smith = df_parts['Item'].astype(str).str.lower().str.contains(r'a\.?o\.?\s*smith', regex=True, na=False)
        df_parts_clean = df_parts[~is_ao_smith]

        sku_pivot = pd.pivot_table(
            df_parts_clean,
            index=['SKU', 'Item'],
            columns='Tech',
            values='Qty',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        cost_map = df_parts_clean.groupby('SKU')['Total Value'].sum().to_dict()
        sku_pivot['Total Units Used'] = sku_pivot.iloc[:, 2:].sum(axis=1)
        sku_pivot['Total Value ($)'] = sku_pivot['SKU'].map(cost_map).fillna(0)
        sku_pivot.sort_values(by='Total Value ($)', ascending=False, inplace=True)
        sku_pivot['Total Value ($)'] = sku_pivot['Total Value ($)'].map('${:,.2f}'.format)

        st.dataframe(sku_pivot, use_container_width=True, hide_index=True)

    st.markdown("---")
    # --- TEST TABLE 4: BU Level Replenishment Efficiency & Material Ratios ---
    st.subheader("4. BU-Level Replenishment Efficiency & Material Ratios")
    st.markdown("""
    Evaluates technician replenishment intensity against expected business unit material ratios.
    *Note: In Section B, only **Water Heater Parts** usage is counted towards Net Replenishment Cost so actual water heater unit tanks do not distort material ratios. Expected ratio range: **3.5% – 6.0%**. Mathew Hodges is based in Tucson and does not pull from the main warehouse.*
    """)

    def get_bu_efficiency_table(bu_name, min_material_ratio_threshold, max_material_ratio_threshold):
        bu_rows = []
        for t in sorted(VALID_TECHS):
            if not df_parts.empty:
                if 'water heater' in bu_name.lower():
                    p_sub = df_parts[
                        (df_parts['Tech'] == t) & 
                        (df_parts['Business Unit'].str.contains('Water Heater', case=False, na=False)) & 
                        (~df_parts['Business Unit'].str.contains('Units', case=False, na=False))
                    ]
                else:
                    p_sub = df_parts[
                        (df_parts['Tech'] == t) & 
                        (df_parts['Business Unit'].str.contains('Simple Installs', case=False, na=False))
                    ]
                parts_cost = p_sub['Total Value'].sum()
            else:
                parts_cost = 0.0

            j_count = 0
            if not jobs_filtered.empty:
                j_sub = jobs_filtered[(jobs_filtered['Business Unit'] == bu_name) & (jobs_filtered['Tech Clean'] == t)]
                j_count = len(j_sub)

            rev = 0.0
            if not inv_filtered.empty:
                i_sub = inv_filtered[(inv_filtered['Business Unit Clean'] == bu_name) & (inv_filtered['Tech Clean'] == t)]
                rev = i_sub['Invoice Total'].sum()

            cost_per_job = (parts_cost / j_count) if j_count > 0 else 0.0
            mat_pct = (parts_cost / rev * 100) if rev > 0 else 0.0

            if t == "Mathew Hodges":
                flag = "🌵 Tucson Tech (No Warehouse Restocks)"
            elif j_count > 0 and parts_cost == 0:
                flag = "⚠️ Zero Parts Restocked"
            elif mat_pct > max_material_ratio_threshold:
                flag = f"🔴 High Material % (>{max_material_ratio_threshold:.1f}%)"
            elif mat_pct > 0 and mat_pct < min_material_ratio_threshold and j_count > 5:
                flag = f"🟡 Low Material % (<{min_material_ratio_threshold:.1f}%)"
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

    st.markdown("##### A. Lowes - Simple Installs (Expected Material Ratio: 1.0% – 3.5%)")
    simple_eff_df = get_bu_efficiency_table('Lowes - Simple Installs', min_material_ratio_threshold=1.0, max_material_ratio_threshold=3.5)
    if not simple_eff_df.empty:
        st.dataframe(simple_eff_df, use_container_width=True, hide_index=True)
    else:
        st.info("No data available for Simple Installs.")

    st.markdown("##### B. Lowes - Water Heaters (Expected Material Ratio: 3.5% – 6.0%)")
    wh_eff_df = get_bu_efficiency_table('Lowes - Water Heaters', min_material_ratio_threshold=3.5, max_material_ratio_threshold=6.0)
    if not wh_eff_df.empty:
        st.dataframe(wh_eff_df, use_container_width=True, hide_index=True)
    else:
        st.info("No data available for Water Heaters.")

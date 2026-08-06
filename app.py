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

# Display active tech roster in sidebar without showing pay rates
st.sidebar.markdown("---")
st.sidebar.markdown("### 👷 Tech Roster")
for t, p in PAY_STRUCTURE.items():
    loc_tag = "🌵 Tucson" if p["location"] == "Tucson" else "📍 Phoenix"
    st.sidebar.markdown(f"**{t}** ({loc_tag})")

# --- PRE-PROCESS ALL DATA ONCE ---
# 1. Parts Data from Google Sheets
sheets_dict = fetch_live_google_sheet(sheet_url)
df_parts = pd.DataFrame()
df_current_minmax = pd.DataFrame()

if sheets_dict:
    sheet_names = list(sheets_dict.keys())
    
    simple_sheet, wh_parts_sheet, wh_units_sheet, minmax_sheet = None, None, None, None
    
    # Identify sheets safely
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
            
    # Fallbacks
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
    
    # Process Min/Max sheet
    if minmax_sheet and minmax_sheet in sheets_dict:
        raw_minmax = sheets_dict[minmax_sheet].copy()
        
        # Flatten all columns first to avoid weird hidden newlines
        raw_minmax.columns = raw_minmax.columns.astype(str).str.replace('\n', ' ').str.strip()
        
        # Fuzzy column locator
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
                
            if c_qty:
                raw_minmax['Current On Hand'] = pd.to_numeric(raw_minmax[c_qty], errors='coerce').fillna(0).astype(int)
            else:
                raw_minmax['Current On Hand'] = 0
                
            if c_dept:
                raw_minmax['Business Unit_sheet'] = raw_minmax[c_dept].apply(map_dept_to_bu)
            else:
                raw_minmax['Business Unit_sheet'] = 'Unknown'
                
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

# 3. Invoices Data (Excludes Draft & Void; attributing full credit to first listed tech)
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

# 4. Timesheets Data with Weekly Overtime Logic
ts_df = read_uploaded_csv(uploaded_timesheets)
if ts_df is not None and 'Clock In Date/Time' in ts_df.columns:
    ts_df['Tech Clean'] = ts_df['User'].apply(get_first_valid_tech)
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

# --- TAB 3: WAREHOUSE MIN/MAX COMPARISON ---
with tab_minmax:
    st.header("📦 Warehouse Min/Max Analysis & Comparison")
    st.markdown("""
    Comparison of **Current Warehouse Min/Max settings** against **Suggested 1.5-Week Inventory Targets** calculated from historical demand (July 2026 / 4.43 weeks).
    
    *Only triggers an action if the suggested Min or Max deviates by **20% or more** from the current setting in Google Sheets. Organized with the highest priority items at the top.*
    """)
    
    if not df_parts.empty:
        total_weeks = 31.0 / 7.0
        
        # Calculate item-level demand from df_parts
        item_usage = df_parts.groupby(['Business Unit', 'SKU', 'Item']).agg(
            Total_Net_Qty=('Qty', 'sum'),
            Total_Net_Cost=('Total Value', 'sum')
        ).reset_index()
        
        item_usage['SKU_clean'] = item_usage['SKU'].apply(clean_sku)
        item_usage['Weekly_Avg_Qty'] = item_usage['Total_Net_Qty'] / total_weeks
        item_usage['Min_Stock_Qty'] = np.ceil(item_usage['Weekly_Avg_Qty'] * 1.0).clip(lower=1).astype(int)
        item_usage['Target_Stock_Qty'] = np.ceil(item_usage['Weekly_Avg_Qty'] * 1.5).astype(int)
        item_usage['Max_Stock_Qty'] = np.maximum(np.ceil(item_usage['Weekly_Avg_Qty'] * 2.0), item_usage['Min_Stock_Qty'] + 1).astype(int)
        
        # Perform outer merge exclusively on SKU_clean
        if not df_current_minmax.empty:
            merged_minmax = pd.merge(
                item_usage[['SKU_clean', 'Business Unit', 'Item', 'Total_Net_Qty', 'Weekly_Avg_Qty', 'Min_Stock_Qty', 'Target_Stock_Qty', 'Max_Stock_Qty']],
                df_current_minmax, 
                on=['SKU_clean'], 
                how='outer'
            )
            
            # Safe Business Unit Resolver handling missing keys cleanly
            def resolve_wh_bu(row):
                bu = row.get('Business Unit')
                if pd.notna(bu) and str(bu).strip() != '' and bu != 'Unknown':
                    return bu
                
                bu_sheet = row.get('Business Unit_sheet', 'Unknown')
                if 'water heater' in str(bu_sheet).lower():
                    item_name_val = row.get('Item Name', '')
                    item_val = row.get('Item', '')
                    desc = (str(item_name_val) + " " + str(item_val)).lower()
                    
                    # Precise Unit heuristic: Must match full water heater unit descriptions
                    # and must NOT be an expansion tank or galvanized fitting
                    is_unit = (
                        'ao smith' in desc or 
                        'a.o. smith' in desc or 
                        re.search(r'\b(30|40|50|75|80)\s*gal\b', desc)
                    ) and not ('expansion tank' in desc or 'galv' in desc)

                    if is_unit:
                        return 'Lowes - Water Heaters (Units)'
                    else:
                        return 'Lowes - Water Heaters (Parts)'
                return bu_sheet

            merged_minmax['Business Unit'] = merged_minmax.apply(resolve_wh_bu, axis=1)
            
            # Fill Descriptions safely
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

        # Compact Min & Max Comparison formatting
        merged_minmax['Min (Curr ➔ Sug)'] = merged_minmax['Current Min'].astype(str) + " ➔ " + merged_minmax['Min_Stock_Qty'].astype(str)
        merged_minmax['Max (Curr ➔ Sug)'] = merged_minmax['Current Max'].astype(str) + " ➔ " + merged_minmax['Max_Stock_Qty'].astype(str)

        # Recommendation Logic with 20% threshold buffer
        def get_minmax_recommendation(row):
            c_min, c_max = row['Current Min'], row['Current Max']
            s_min, s_max = row['Min_Stock_Qty'], row['Max_Stock_Qty']
            
            if c_min == 0 and c_max == 0:
                if s_min == 0 and s_max == 0:
                    return "⚪ Zero Demand"
                return "⚠️ Set Min/Max"
            
            d_min = s_min - c_min
            d_max = s_max - c_max
            
            # Check if difference is >= 20%
            min_flag = abs(d_min) >= (0.20 * c_min) if c_min > 0 else (s_min > 0)
            max_flag = abs(d_max) >= (0.20 * c_max) if c_max > 0 else (s_max > 0)

            if not min_flag and not max_flag:
                return "🟢 On Target"

            rec = []
            if min_flag:
                if d_min > 0:
                    rec.append(f"Inc Min (+{d_min})")
                else:
                    rec.append(f"Dec Min ({d_min})")
                    
            if max_flag:
                if d_max > 0:
                    rec.append(f"Inc Max (+{d_max})")
                else:
                    rec.append(f"Dec Max ({d_max})")
                    
            if len(rec) == 2:
                if d_min > 0 and d_max > 0:
                    return f"⬆️ {rec[0]} & {rec[1].replace('Inc ', '')}"
                elif d_min < 0 and d_max < 0:
                    return f"⬇️ {rec[0]} & {rec[1].replace('Dec ', '')}"
                else:
                    return f"🔄 {rec[0]} & {rec[1]}"
            elif len(rec) == 1:
                if "Inc" in rec[0]:
                    return f"⬆️ {rec[0]}"
                else:
                    return f"⬇️ {rec[0]}"

            return "🟢 On Target"

        merged_minmax['Action / Rec'] = merged_minmax.apply(get_minmax_recommendation, axis=1)

        # Priority Sorter for the dataframe (so important items float to the top)
        def get_sort_priority(action_str):
            if "⚠️" in action_str: return 1
            if "⬆️" in action_str or "⬇️" in action_str or "🔄" in action_str: return 2
            if "🟢" in action_str: return 3
            return 4

        merged_minmax['Sort_Priority'] = merged_minmax['Action / Rec'].apply(get_sort_priority)

        def render_comparison_table(bu_name):
            bu_df = merged_minmax[merged_minmax['Business Unit'] == bu_name].copy()
            if not bu_df.empty:
                # Sort by action priority first, then by target stock volume
                bu_df.sort_values(by=['Sort_Priority', 'Target_Stock_Qty'], ascending=[True, False], inplace=True)
                
                bu_df.rename(columns={
                    'SKU': 'SKU',
                    'Total_Net_Qty': 'July Net',
                    'Weekly_Avg_Qty': 'Wk Avg',
                    'Current On Hand': 'On Hand',
                    'Target_Stock_Qty': 'Target (1.5 Wk)'
                }, inplace=True)
                
                bu_df['Wk Avg'] = bu_df['Wk Avg'].map('{:.2f}'.format)
                
                show_cols = [
                    'SKU', 'Item Description', 'July Net', 'Wk Avg', 
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
        render_comparison_table('Lowes - Simple Installs')
        
        st.subheader("2. Lowes - Water Heaters (Parts) Min/Max Comparison")
        render_comparison_table('Lowes - Water Heaters (Parts)')

        st.subheader("3. Lowes - Water Heaters (Units) Min/Max Comparison")
        render_comparison_table('Lowes - Water Heaters (Units)')

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

# --- TAB 7: TEST SECTION - BU LEVEL EFFICIENCY ---
with tab_test:
    st.header("🧪 Test Section: BU-Level Replenishment Efficiency & Material Ratios")
    st.markdown("""
    This section explicitly separates **Simple Installs** and **Water Heaters** to evaluate technician replenishment 
    intensity against expected business unit ratios. Full job and revenue credit is attributed to the first listed technician.
    *Note: Mathew Hodges is based in Tucson and does not pull from the main warehouse.*
    """)

    def get_bu_efficiency_table(bu_name, max_material_ratio_threshold):
        bu_rows = []
        for t in sorted(VALID_TECHS):
            if not df_parts.empty:
                p_sub = df_parts[(df_parts['Tech'] == t) & (df_parts['Business Unit'].str.contains(bu_name, case=False, na=False))]
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

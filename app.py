import streamlit as st
import pandas as pd
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="NexSys Data Analyzer", layout="wide")
st.title("NexSys Parts & Jobs Analysis")
st.markdown("Analyze simple installs and water heater jobs at the technician, business unit, and job title level.")

# --- HELPER FUNCTIONS ---
@st.cache_data
def parse_messy_parts_csv(text, business_unit):
    """Parses the raw text format from the Google Doc into a DataFrame."""
    rows = []
    if not text.strip():
        return pd.DataFrame()
        
    lines = text.strip().split('\n')
    for line in lines[1:]: # Skip header
        parts = line.split(',')
        if len(parts) >= 9:
            transferred_to = parts[-4]
            qty = parts[-3]
            unit_cost = parts[-2]
            total_value = parts[-1]
            from_loc = parts[-5]
            
            # Handle dates that contain commas (e.g., "Jul 10, 2026")
            if parts[0].startswith(('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')) and len(parts) > 1 and len(parts[1].strip()) == 4:
                date = parts[0] + "," + parts[1]
                direction = parts[2]
                sku = parts[3]
                item = ",".join(parts[4:-5])
            else:
                date = parts[0]
                direction = parts[1]
                sku = parts[2]
                item = ",".join(parts[3:-5])
                
            try:
                rows.append({
                    'Date': date, 
                    'Direction': direction, 
                    'SKU': sku, 
                    'Item': item.strip('"'),
                    'From': from_loc, 
                    'Transferred To': transferred_to,
                    'Qty': int(qty), 
                    'Unit Cost': float(unit_cost.replace('$', '').replace(',', '')),
                    'Total Value': float(total_value.replace('$', '').replace(',', ''))
                })
            except ValueError:
                continue # Skip malformed rows
                
    df = pd.DataFrame(rows)
    if not df.empty:
        df['Business Unit'] = business_unit
    return df

def map_tech_name(fleet_name):
    """Normalizes fleet names to match technician names in the jobs CSV."""
    fleet = str(fleet_name)
    if "Bill" in fleet: return "Bill"
    if "Bryan" in fleet: return "Bryan Pickett"
    if "Carmen" in fleet: return "Carmen Tripodi"
    if "Erik" in fleet: return "Erik Tange"
    if "Matt S" in fleet or "Matt's" in fleet: return "Matt Schlosser"
    if "Sean" in fleet: return "Sean Marble"
    if "Tanner" in fleet: return "Tanner LaForge"
    return fleet

# --- SIDEBAR & FILE UPLOADS ---
st.sidebar.header("1. Upload Data Files")
uploaded_jobs = st.sidebar.file_uploader("Upload 'all jobs.csv'", type=['csv'])

st.sidebar.header("2. Input Parts Data")
st.sidebar.markdown("Paste the raw parts text from your Google Doc below:")
simple_parts_text = st.sidebar.text_area("Simple Install Parts Data", height=150)
wh_parts_text = st.sidebar.text_area("Water Heater Parts Data", height=150)

# --- TABS ---
tab1, tab2 = st.tabs(["⚙️ Parts Usage", "📊 Jobs Analysis"])

# --- TAB 1: PARTS USAGE ---
with tab1:
    st.header("Parts Usage by Technician & Business Unit")
    
    if simple_parts_text or wh_parts_text:
        # Parse text
        df_simple = parse_messy_parts_csv(simple_parts_text, 'Lowes - Simple Installs')
        df_wh = parse_messy_parts_csv(wh_parts_text, 'Lowes - Water Heaters')
        
        # Combine
        df_parts = pd.concat([df_simple, df_wh], ignore_index=True)
        
        if not df_parts.empty:
            # Clean up Tech names
            df_parts['Transferred To'] = df_parts['Transferred To'].str.replace("Matt's TransitFleet", "Matt S")
            df_parts['Tech'] = df_parts['Transferred To'].apply(map_tech_name)
            
            # Aggregate
            parts_summary = df_parts.groupby(['Tech', 'Business Unit'])['Total Value'].sum().reset_index()
            parts_summary['Total Value'] = parts_summary['Total Value'].map('${:,.2f}'.format)
            
            col1, col2 = st.columns([2, 3])
            
            with col1:
                st.dataframe(parts_summary, use_container_width=True)
                
            with col2:
                # Chart
                chart_data = df_parts.groupby(['Tech', 'Business Unit'])['Total Value'].sum().unstack().fillna(0)
                st.bar_chart(chart_data)
        else:
            st.warning("Could not parse parts data. Ensure it matches the CSV comma format from the Google Doc.")
    else:
        st.info("👈 Paste your parts data in the sidebar to see the analysis.")

# --- TAB 2: JOBS ANALYSIS ---
with tab2:
    st.header("Job Analysis by Technician, Business Unit & Job Title")
    
    if uploaded_jobs is not None:
        # Load Jobs Data (skipping the first row if it's a double header based on typical ServiceTitan exports)
        try:
            jobs_df = pd.read_csv(uploaded_jobs, header=1)
            if 'Business Unit' not in jobs_df.columns:
                # Fallback in case header=1 was wrong for their specific export
                uploaded_jobs.seek(0)
                jobs_df = pd.read_csv(uploaded_jobs)
        except Exception as e:
            st.error(f"Error loading CSV: {e}")
            st.stop()
            
        # Filter for valid Business Units
        valid_bus = ['Lowes - Simple Installs', 'Lowes - Water Heaters']
        
        if 'Business Unit' in jobs_df.columns:
            jobs_filtered = jobs_df[jobs_df['Business Unit'].isin(valid_bus)].copy()
            
            # Clean financials
            jobs_filtered['Invoice Amount'] = pd.to_numeric(jobs_filtered['Total Invoice Amount'], errors='coerce').fillna(0)
            
            # Aggregate
            job_summary = jobs_filtered.groupby(['Assigned Team Members', 'Business Unit', 'Title']).agg(
                Job_Count=('Title', 'count'),
                Total_Invoice_Amount=('Invoice Amount', 'sum')
            ).reset_index()
            
            job_summary.rename(columns={'Assigned Team Members': 'Tech', 'Title': 'Job Title'}, inplace=True)
            
            # Display format
            display_df = job_summary.copy()
            display_df['Total_Invoice_Amount'] = display_df['Total_Invoice_Amount'].map('${:,.2f}'.format)
            
            st.dataframe(display_df.sort_values(by=['Tech', 'Business Unit']), use_container_width=True, hide_index=True)
            
            st.subheader("Invoice Revenue by Technician")
            revenue_chart = jobs_filtered.groupby(['Assigned Team Members', 'Business Unit'])['Invoice Amount'].sum().unstack().fillna(0)
            st.bar_chart(revenue_chart)
            
        else:
            st.error("The uploaded CSV does not contain a 'Business Unit' column. Please check your export format.")
    else:
        st.info("👈 Upload your 'all jobs.csv' file in the sidebar to view job volume and invoice totals.")
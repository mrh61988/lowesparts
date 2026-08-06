import streamlit as st
import pandas as pd

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="NexSys Data Analyzer", layout="wide")
st.title("NexSys Parts & Jobs Analysis")
st.markdown("Analyze simple installs and water heater jobs at the technician, business unit, and job title level.")

# --- HELPER FUNCTIONS ---
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

@st.cache_data
def load_google_sheet(url):
    """Reads a public Google Sheet URL into a dictionary of pandas DataFrames."""
    if not url:
        return None
        
    # Convert the Google Sheets link to a direct Excel export link
    base_url = url.split('/edit')[0]
    export_url = f"{base_url}/export?format=xlsx"
    
    try:
        # Read all sheets into a dictionary of DataFrames
        xls = pd.read_excel(export_url, sheet_name=None)
        return xls
    except Exception as e:
        st.error(f"Error reading Google Sheet. Ensure the link is public ('Anyone with the link can view'). Details: {e}")
        return None

def process_parts_df(df, business_unit):
    """Cleans the parts DataFrame loaded from Google Sheets."""
    df = df.copy()
    
    # Check if necessary columns exist
    if 'Transferred To' not in df.columns or 'Total Value' not in df.columns:
        st.warning(f"Could not find 'Transferred To' or 'Total Value' columns in the selected sheet for {business_unit}.")
        return pd.DataFrame()
        
    # Clean up currency strings if they are formatted as text
    if df['Total Value'].dtype == object:
        df['Total Value'] = df['Total Value'].astype(str).replace('[\$,]', '', regex=True)
    
    df['Total Value'] = pd.to_numeric(df['Total Value'], errors='coerce').fillna(0)
    
    # Clean up Tech names
    df['Transferred To'] = df['Transferred To'].astype(str).str.replace("Matt's TransitFleet", "Matt S")
    df['Tech'] = df['Transferred To'].apply(map_tech_name)
    df['Business Unit'] = business_unit
    
    return df

# --- SIDEBAR & FILE UPLOADS ---
st.sidebar.header("1. Upload Jobs Data")
uploaded_jobs = st.sidebar.file_uploader("Upload 'all jobs.csv'", type=['csv'])

st.sidebar.header("2. Link Parts Data")
sheet_url = st.sidebar.text_input(
    "Google Sheets URL", 
    value="https://docs.google.com/spreadsheets/d/1OR4mEgviGglKNLwinPLnc8NB3FrN7VH9rt5qAvo8RRs/edit?usp=sharing"
)

# Load the Google Sheet
sheets_dict = load_google_sheet(sheet_url)

simple_sheet_name = None
wh_sheet_name = None

if sheets_dict:
    sheet_names = list(sheets_dict.keys())
    st.sidebar.markdown("### Map your sheets:")
    simple_sheet_name = st.sidebar.selectbox("Select Simple Installs Parts Sheet", options=sheet_names, index=0)
    
    # Default to second sheet if available, otherwise first
    default_wh_index = 1 if len(sheet_names) > 1 else 0
    wh_sheet_name = st.sidebar.selectbox("Select Water Heaters Parts Sheet", options=sheet_names, index=default_wh_index)

# --- TABS ---
tab1, tab2 = st.tabs(["⚙️ Parts Usage", "📊 Jobs Analysis"])

# --- TAB 1: PARTS USAGE ---
with tab1:
    st.header("Parts Usage by Technician & Business Unit")
    
    if sheets_dict and simple_sheet_name and wh_sheet_name:
        # Process dataframes based on user selections
        df_simple = process_parts_df(sheets_dict[simple_sheet_name], 'Lowes - Simple Installs')
        df_wh = process_parts_df(sheets_dict[wh_sheet_name], 'Lowes - Water Heaters')
        
        # Combine them
        df_parts = pd.concat([df_simple, df_wh], ignore_index=True)
        
        if not df_parts.empty:
            # Aggregate
            parts_summary = df_parts.groupby(['Tech', 'Business Unit'])['Total Value'].sum().reset_index()
            
            # Formatting for display
            display_parts = parts_summary.copy()
            display_parts['Total Value'] = display_parts['Total Value'].map('${:,.2f}'.format)
            
            col1, col2 = st.columns([2, 3])
            
            with col1:
                st.dataframe(display_parts, use_container_width=True, hide_index=True)
                
            with col2:
                # Pivot for Chart
                chart_data = parts_summary.pivot(index='Tech', columns='Business Unit', values='Total Value').fillna(0)
                st.bar_chart(chart_data)
        else:
            st.warning("Processed parts data is empty. Check your Google Sheet formatting.")
    else:
        st.info("👈 Enter your Google Sheets URL in the sidebar.")

# --- TAB 2: JOBS ANALYSIS ---
with tab2:
    st.header("Job Analysis by Technician, Business Unit & Job Title")
    
    if uploaded_jobs is not None:
        # Load Jobs Data 
        try:
            # Typically ServiceTitan has two header rows, skip the first if necessary
            jobs_df = pd.read_csv(uploaded_jobs, header=1)
            if 'Business Unit' not in jobs_df.columns:
                uploaded_jobs.seek(0)
                jobs_df = pd.read_csv(uploaded_jobs)
        except Exception as e:
            st.error(f"Error loading CSV: {e}")
            st.stop()
            
        valid_bus = ['Lowes - Simple Installs', 'Lowes - Water Heaters']
        
        if 'Business Unit' in jobs_df.columns:
            jobs_filtered = jobs_df[jobs_df['Business Unit'].isin(valid_bus)].copy()
            
            # Clean financial amounts
            jobs_filtered['Invoice Amount'] = pd.to_numeric(jobs_filtered['Total Invoice Amount'], errors='coerce').fillna(0)
            
            # Aggregate
            job_summary = jobs_filtered.groupby(['Assigned Team Members', 'Business Unit', 'Title']).agg(
                Job_Count=('Title', 'count'),
                Total_Invoice_Amount=('Invoice Amount', 'sum')
            ).reset_index()
            
            job_summary.rename(columns={'Assigned Team Members': 'Tech', 'Title': 'Job Title'}, inplace=True)
            
            display_df = job_summary.copy()
            display_df['Total_Invoice_Amount'] = display_df['Total_Invoice_Amount'].map('${:,.2f}'.format)
            
            st.dataframe(display_df.sort_values(by=['Tech', 'Business Unit']), use_container_width=True, hide_index=True)
            
            st.subheader("Invoice Revenue by Technician")
            revenue_chart = jobs_filtered.groupby(['Assigned Team Members', 'Business Unit'])['Invoice Amount'].sum().unstack().fillna(0)
            st.bar_chart(revenue_chart)
            
        else:
            st.error("The uploaded CSV does not contain a 'Business Unit' column. Check your export format.")
    else:
        st.info("👈 Upload your 'all jobs.csv' file in the sidebar to view job volume and invoice totals.")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import requests
import time
from datetime import datetime, timedelta
import pytz
import re
import warnings

# Ignore pandas future warnings for clean logs
warnings.simplefilter(action='ignore', category=FutureWarning)

# =========================================================================
# PAGE CONFIGURATION
# =========================================================================
st.set_page_config(
    page_title="Data Quality Dashboard LDD4IG II",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Header
st.markdown("""
    <style>
    .main-header {
        background-color: #006EB6;
        padding: 15px 0;
        color: #FFFFFF;
        font-size: 26px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 20px;
    }
    </style>
    <div class="main-header">Data Quality Dashboard LDD4IG II</div>
""", unsafe_allow_html=True)


# =========================================================================
# 1. GLOBAL DATA: READ EXCEL
# =========================================================================
@st.cache_data
def load_excel_data(file_path="quality.xlsx"):
    try:
        df = pd.read_excel(file_path, dtype=str)
    except FileNotFoundError:
        df = pd.DataFrame()

    if df.empty:
        return df, pd.DataFrame(columns=["Upazila", "Known_District"])

    # Make columns unique
    cols = pd.Series(df.columns).str.strip()
    for dup in cols[cols.duplicated()].unique(): 
        cols[cols[cols == dup].index.values.tolist()] = [dup + '.' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols

    # Smart Rename Engine
    col_names = df.columns
    def rename_first_match(pattern, new_name, exclude_pattern=None):
        for col in col_names:
            if re.search(pattern, col) and (not exclude_pattern or not re.search(exclude_pattern, col)):
                df.rename(columns={col: new_name}, inplace=True)
                break

    rename_first_match(r"জেলা", "District", exclude_pattern=r"উপজেলা")
    rename_first_match(r"উপজেলা", "Upazila")
    rename_first_match(r"অফিস", "Office")
    rename_first_match(r"ইউজার", "User")
    rename_first_match(r"চেকার", "Checker", exclude_pattern=r"তারিখ")
    
    date_cols = [c for c in df.columns if "র‍্যান্ডম" in c and "তারিখ" in c]
    if date_cols:
        df.rename(columns={date_cols[0]: "Date_Raw"}, inplace=True)
    else:
        rename_first_match(r"তারিখ", "Date_Raw")
        
    rename_first_match(r"চেক করে যা|পাওয়া গেলো", "Check_Result")

    if "District" in df.columns and "Upazila" in df.columns:
        upazila_dict = df.dropna(subset=["District", "Upazila"])[["Upazila", "District"]].drop_duplicates(subset=["Upazila"])
        upazila_dict.rename(columns={"District": "Known_District"}, inplace=True)
    else:
        upazila_dict = pd.DataFrame(columns=["Upazila", "Known_District"])

    return df, upazila_dict


# =========================================================================
# 2. SERVER LOGIC & DIRECT API FETCH (Cached with TTL for 1 hour)
# =========================================================================
@st.cache_data(ttl=3600)  
def fetch_live_data():
    excel_df, upazila_dict = load_excel_data()
    
    api_url = "https://log.ldd4ig.org/api/data/random_checked_data"
    tz = pytz.timezone('Asia/Dhaka')
    api_start = datetime(2026, 6, 21).date()
    api_end = datetime.now(tz).date()
    
    date_seq = [api_start + timedelta(days=x) for x in range((api_end - api_start).days + 1)]
    api_data_list = []
    
    for d in date_seq:
        c_date_api = d.strftime("%Y/%m/%d")
        c_date_fallback = d.strftime("%Y-%m-%d")
        
        for attempt in range(3):
            try:
                response = requests.post(
                    api_url, 
                    json={"startDate": c_date_api, "endDate": c_date_api},
                    headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
                    verify=False,
                    timeout=45
                )
                if response.status_code == 200:
                    parsed_json = response.json()
                    temp_data = pd.DataFrame()
                    if isinstance(parsed_json, dict) and "data" in parsed_json:
                        temp_data = pd.DataFrame(parsed_json["data"])
                    elif isinstance(parsed_json, list) and len(parsed_json) > 0:
                        temp_data = pd.DataFrame(parsed_json[0])
                    
                    if not temp_data.empty:
                        temp_data.columns = [c.strip() for c in temp_data.columns]
                        rename_map = {
                            "random_check_comment": "Check_Result", "random_check_date": "Date_Raw",
                            "created_at": "Date_Raw", "date": "Date_Raw", "location_name": "Upazila",
                            "office_name": "Office", "user_name": "User", "random_check_user": "Checker"
                        }
                        temp_data.rename(columns=lambda x: rename_map.get(x, x), inplace=True)
                        temp_data = temp_data.astype(str)
                        temp_data["Fallback_Date"] = c_date_fallback
                        api_data_list.append(temp_data)
                    break 
            except Exception:
                time.sleep(2)
        time.sleep(0.5)

    api_df = pd.concat(api_data_list, ignore_index=True) if api_data_list else pd.DataFrame()
    raw_data = pd.concat([excel_df, api_df], ignore_index=True)
    
    required_cols = ["District", "Upazila", "Office", "User", "Checker", "Date_Raw", "Check_Result", "Fallback_Date"]
    for col in required_cols:
        if col not in raw_data.columns: raw_data[col] = np.nan

    if not raw_data.empty and not upazila_dict.empty:
        final_data = pd.merge(raw_data, upazila_dict, on="Upazila", how="left")
        final_data["District"] = final_data["District"].combine_first(final_data["Known_District"])
    else:
        final_data = raw_data

    final_data["Check_Result"] = final_data["Check_Result"].str.strip()
    
    # Bulletproof Date Parsing for Excel
    def parse_date(d_str):
        if pd.isna(d_str) or str(d_str).strip() in ["", "nan", "NaT", "None"]: return pd.NaT
        d_str = str(d_str).strip()
        # Handle Excel serial numbers (e.g. 45000)
        if re.match(r"^\d{5}(\.\d+)?$", d_str):
            try: return pd.to_datetime("1899-12-30") + pd.to_timedelta(float(d_str), unit="D")
            except: pass
        if "T" in d_str: d_str = d_str[:10]
        try: return pd.to_datetime(d_str, format='mixed', dayfirst=True, errors='coerce')
        except: return pd.to_datetime(d_str, errors='coerce')

    final_data["Date_Parsed"] = final_data["Date_Raw"].apply(parse_date)
    final_data["Date"] = final_data["Date_Parsed"].combine_first(pd.to_datetime(final_data["Fallback_Date"], errors='coerce'))
    
    # Status Mapping 
    final_data["Status"] = "ভুল"  
    is_correct = final_data["Check_Result"].str.contains("সব তথ্য ঠিক আছে", na=False)
    final_data.loc[is_correct, "Status"] = "সঠিক"
    is_blank = final_data["Check_Result"].isna() | (final_data["Check_Result"] == "")
    final_data.loc[is_blank, "Status"] = np.nan
    
    # Track drops for debugging UI
    total_before = len(final_data)
    final_data = final_data.dropna(subset=["Status", "Date"]).drop_duplicates()
    dropped_rows = total_before - len(final_data)
    
    return final_data, datetime.now(tz).strftime("%d-%b-%Y %I:%M:%S %p"), len(excel_df), dropped_rows


# Fetch the data
df, last_sync_time, excel_count, dropped_count = fetch_live_data()


# =========================================================================
# 3. USER INTERFACE (SIDEBAR & FILTERS)
# =========================================================================
st.sidebar.title("Filter")

# DYNAMIC DATE CALENDAR (Scans your data to find the true min/max dates)
if not df.empty and pd.notna(df["Date"].min()):
    min_date = df["Date"].min().date()
    max_date = df["Date"].max().date()
else:
    min_date = datetime(2023, 1, 1).date()
    max_date = datetime.now().date()

date_range = st.sidebar.date_input(
    "র‍্যান্ডম চেকের তারিখ (Date)", 
    [min_date, max_date], 
    min_value=min_date, 
    max_value=max_date
)

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start_date, end_date = date_range
    filtered_df = df[(df["Date"].dt.date >= start_date) & (df["Date"].dt.date <= end_date)].copy()
else:
    filtered_df = df.copy()

# District Filter
districts = ["All"] + sorted([str(x) for x in filtered_df["District"].dropna().unique()])
selected_district = st.sidebar.selectbox("জেলা (District)", districts)
if selected_district != "All":
    filtered_df = filtered_df[filtered_df["District"] == selected_district]

# Upazila Filter
upazilas = ["All"] + sorted([str(x) for x in filtered_df["Upazila"].dropna().unique()])
selected_upazila = st.sidebar.selectbox("Upazila/Revenue Circle", upazilas)
if selected_upazila != "All":
    filtered_df = filtered_df[filtered_df["Upazila"] == selected_upazila]

# Office Filter
offices = ["All"] + sorted([str(x) for x in filtered_df["Office"].dropna().unique()])
selected_office = st.sidebar.selectbox("ULO", offices)
if selected_office != "All":
    filtered_df = filtered_df[filtered_df["Office"] == selected_office]

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='color: #7f8c8d; font-size: 12px; text-align: center;'>🔄 Live API Sync Active: Auto-updates every 1 hour.</p>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div style='color: #27ae60; font-size: 11px; text-align: center; font-weight: bold; margin-top: -10px;'>Last Fetched: {last_sync_time}</div>", unsafe_allow_html=True)

# --- DIAGNOSTIC PANEL ---
st.sidebar.markdown("---")
st.sidebar.info(f"📁 **Excel Check:** {excel_count} rows found.\n🧹 **Data Dropped:** {dropped_count} rows missing Date/Status.")

# Helper indicator columns for simplified aggregations
filtered_df["Is_Error"] = (filtered_df["Status"] == "ভুল").astype(int)

# =========================================================================
# 4. KPIs
# =========================================================================
total_checked = len(filtered_df)
correct = len(filtered_df[filtered_df["Status"] == "সঠিক"])
errors = len(filtered_df[filtered_df["Status"] == "ভুল"])
error_pct = f"{round((errors / total_checked) * 100, 2)}%" if total_checked > 0 else "0%"

col1, col2, col3, col4 = st.columns(4)
col1.metric("মোট যাচাই", total_checked)
col2.metric("সব তথ্য ঠিক আছে", correct)
col3.metric("মোট ভুল", errors)
col4.metric("শতকরা ভুল", error_pct)

st.markdown("---")


# =========================================================================
# 5. TABS & VISUALIZATIONS
# =========================================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "ULO-Wise", "DEO-Wise", "Random Checker Wise", 
    "Day-wise Overall (%)", "District Day-wise", "DEO Error Quartiles"
])

def color_error_rows(val):
    if pd.isna(val): return ''
    color = '#990000' if val > 10 else 'inherit'
    bg_color = '#ffcccc' if val > 10 else 'transparent'
    return f'color: {color}; background-color: {bg_color}'

with tab1: 
    if not filtered_df.empty:
        df_ulo = filtered_df.groupby(["Upazila", "Office"]).agg(
            মোট_যাচাই=('Status', 'count'),
            মোট_ভুল=('Is_Error', 'sum')
        ).reset_index()
        
        df_ulo.rename(columns={"মোট_যাচাই": "মোট যাচাই", "মোট_ভুল": "মোট ভুল"}, inplace=True)
        df_ulo["সব তথ্য ঠিক আছে"] = df_ulo["মোট যাচাই"] - df_ulo["মোট ভুল"]
        df_ulo["শতকরা ভুল (%)"] = round((df_ulo["মোট ভুল"] / df_ulo["মোট যাচাই"]) * 100, 2)
        df_ulo.rename(columns={"Upazila": "উপজেলা (Upazila)", "Office": "ULO"}, inplace=True)
        st.dataframe(df_ulo.style.map(color_error_rows, subset=['শতকরা ভুল (%)']), use_container_width=True)

with tab2: 
    if not filtered_df.empty:
        df_deo = filtered_df.groupby(["Upazila", "Office", "User"]).agg(
            মোট_যাচাই=('Status', 'count'),
            মোট_ভুল=('Is_Error', 'sum')
        ).reset_index()
        
        df_deo.rename(columns={"মোট_যাচাই": "মোট যাচাই", "মোট_ভুল": "মোট ভুল"}, inplace=True)
        df_deo = df_deo[df_deo["মোট যাচাই"] >= 10]
        df_deo["সব তথ্য ঠিক আছে"] = df_deo["মোট যাচাই"] - df_deo["মোট ভুল"]
        df_deo["শতকরা ভুল (%)"] = round((df_deo["মোট ভুল"] / df_deo["মোট যাচাই"]) * 100, 2)
        df_deo.rename(columns={"Upazila": "উপজেলা (Upazila)", "Office": "ULO", "User": "DEO"}, inplace=True)
        st.dataframe(df_deo.style.map(color_error_rows, subset=['শতকরা ভুল (%)']), use_container_width=True)

with tab3: 
    if not filtered_df.empty:
        df_checker = filtered_df.groupby(["Upazila", "Office", "Checker"]).agg(
            মোট_যাচাই=('Status', 'count'),
            মোট_ভুল=('Is_Error', 'sum')
        ).reset_index()
        
        df_checker.rename(columns={"মোট_যাচাই": "মোট যাচাই", "মোট_ভুল": "মোট ভুল"}, inplace=True)
        df_checker["সব তথ্য ঠিক আছে"] = df_checker["মোট যাচাই"] - df_checker["মোট ভুল"]
        df_checker["শতকরা ভুল (%)"] = round((df_checker["মোট ভুল"] / df_checker["মোট যাচাই"]) * 100, 2)
        df_checker.rename(columns={"Upazila": "উপজেলা (Upazila)", "Office": "ULO", "Checker": "Random Checker"}, inplace=True)
        
        def color_check_rows(val):
            if pd.isna(val): return ''
            return f'color: {"#990000" if val < 30 else "inherit"}; background-color: {"#ffcccc" if val < 30 else "transparent"}'

        st.dataframe(df_checker.style.map(color_check_rows, subset=['মোট যাচাই']), use_container_width=True)

with tab4: 
    if not filtered_df.empty:
        df_comp = filtered_df.groupby(["Date", "Status"]).size().reset_index(name="Count")
        df_comp["Total_Day"] = df_comp.groupby("Date")["Count"].transform("sum")
        df_comp["Percentage"] = round((df_comp["Count"] / df_comp["Total_Day"]) * 100, 2)
        
        fig1 = px.line(df_comp, x="Date", y="Percentage", color="Status", markers=True,
                       color_discrete_map={"সঠিক": "#2ecc71", "ভুল": "#e74c3c"},
                       title="প্রতিদিনের সঠিক ও ভুল যাচাইয়ের শতকরা হার (%)",
                       labels={"Date": "তারিখ", "Percentage": "শতকরা হার (%)"})
        fig1.update_yaxes(range=[0, 100])
        st.plotly_chart(fig1, use_container_width=True)

with tab5: 
    if not filtered_df.empty:
        df_dist = filtered_df.groupby(["Date", "District"]).agg(
            Total=('Status', 'count'),
            Errors=('Is_Error', 'sum')
        ).reset_index()
        
        df_dist["Error_Pct"] = round((df_dist["Errors"] / df_dist["Total"]) * 100, 2)
        fig2 = px.line(df_dist, x="Date", y="Error_Pct", color="District", markers=True,
                       title="জেলা ভিত্তিক প্রতিদিনের ভুলের হার তুলনা (০-৫০%)",
                       labels={"Date": "তারিখ", "Error_Pct": "ভুলের হার (%)"})
        fig2.update_yaxes(range=[0, 50])
        st.plotly_chart(fig2, use_container_width=True)

with tab6: 
    if not filtered_df.empty:
        deo_stats = filtered_df.groupby(["District", "User"]).agg(
            Total=('Status', 'count'),
            Errors=('Is_Error', 'sum')
        ).reset_index()
        
        deo_stats = deo_stats[deo_stats["Total"] >= 10]
        if not deo_stats.empty:
            deo_stats["Error_Pct"] = round((deo_stats["Errors"] / deo_stats["Total"]) * 100, 2)
            deo_stats = deo_stats.sort_values(["District", "Error_Pct"], ascending=[True, False])
            deo_stats["DEO_Rank"] = deo_stats.groupby("District").cumcount() + 1
            q1, q2, q3, q4 = deo_stats["Error_Pct"].quantile([0.25, 0.50, 0.75, 1.00])
            
            fig3 = px.line(deo_stats, x="DEO_Rank", y="Error_Pct", color="District", markers=True,
                           hover_data=["User", "Total"],
                           title="জেলা ভিত্তিক DEO-দের ভুলের হার (৪টি কোয়ার্টাইলে বিভক্ত)",
                           labels={"DEO_Rank": "জেলার মধ্যে DEO ক্রমিক", "Error_Pct": "ভুলের শতকরা হার (%)"})
            
            fig3.add_hline(y=q1, line_dash="dash", line_color="#27ae60")
            fig3.add_hline(y=q2, line_dash="dash", line_color="#f39c12")
            fig3.add_hline(y=q3, line_dash="dash", line_color="#d35400")
            fig3.add_hline(y=q4, line_dash="dash", line_color="#c0392b")
            st.plotly_chart(fig3, use_container_width=True)

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
import requests
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="VMware Certification Dashboard 2026",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Auto-refresh configuration
REFRESH_INTERVAL = 300  # 5 minutes in seconds

# Custom CSS for professional look - Updated for taller uniform box heights
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2563EB;
        font-weight: 500;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        height: 160px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin-bottom: 1rem;
    }
    .metric-card h4 {
        margin: 0;
        font-size: 1rem;
        font-weight: 400;
        opacity: 0.9;
    }
    .metric-card h2 {
        margin: 0.7rem 0 0 0;
        font-size: 2.5rem;
        font-weight: 700;
    }
    .refresh-badge {
        position: fixed;
        top: 10px;
        right: 10px;
        background-color: #10B981;
        color: white;
        padding: 5px 15px;
        border-radius: 20px;
        font-size: 12px;
        z-index: 999;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    .status-completed {
        background-color: #10B981;
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
    }
    .status-progress {
        background-color: #FEF3C7;
        color: #92400E;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
    }
    .status-notstarted {
        background-color: #EF4444;
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
    }
    div[data-testid="column"] {
        display: flex;
        flex-direction: column;
    }
    .chart-container {
        background-color: white;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# Add auto-refresh meta tag
st.markdown(f"""
    <meta http-equiv="refresh" content="{REFRESH_INTERVAL}" />
    <div class="refresh-badge">
        🔄 Auto-refresh every {REFRESH_INTERVAL//60} minutes | Last: {datetime.now().strftime('%H:%M:%S')}
    </div>
""", unsafe_allow_html=True)

# OneDrive direct download link function
def get_direct_link(share_link):
    """Convert OneDrive sharing link to direct download link"""
    try:
        base_url = share_link.split('?')[0]
        if '/personal/' in base_url:
            return base_url + '?download=1'
        else:
            return base_url.replace('sharepoint.com/:x:', 'sharepoint.com/:x:/') + '?download=1'
    except Exception as e:
        st.warning(f"Error creating direct link: {str(e)}")
        return share_link

def parse_dates(date_str):
    """Parse dates in DD/MM/YY format"""
    if pd.isna(date_str) or date_str == '':
        return pd.NaT
    
    try:
        # Try parsing as DD/MM/YY
        if isinstance(date_str, str):
            # Remove any extra spaces
            date_str = date_str.strip()
            
            # Try different date formats
            for fmt in ['%d/%m/%y', '%d/%m/%Y', '%d-%m-%y', '%d-%m-%Y', '%Y-%m-%d']:
                try:
                    return pd.to_datetime(date_str, format=fmt)
                except:
                    continue
            
            # If specific formats fail, let pandas guess with dayfirst=True
            return pd.to_datetime(date_str, dayfirst=True)
        else:
            # If it's already a datetime or timestamp
            return pd.to_datetime(date_str)
    except:
        return pd.NaT

# Load data function (no cache for real-time updates)
def load_data_from_onedrive():
    """Load Excel data from OneDrive"""
    try:
        # Your OneDrive sharing link
        share_link = "https://jafferbrothers-my.sharepoint.com/:x:/g/personal/customercare_jbs_live/IQAb-s-HyehHTabXHhqu0FTWAVKn9D-CnZ0YH5kw1BZDOGA?e=0RVyMd"
        
        # Get direct download link
        direct_link = get_direct_link(share_link)
        
        # Add headers to mimic a browser request
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        # Download the file
        response = requests.get(direct_link, headers=headers, timeout=30)
        response.raise_for_status()
        
        # Read Excel from the response content
        df = pd.read_excel(BytesIO(response.content), sheet_name="for dashboard", engine='openpyxl')
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Rename the first column if it has the long name
        first_col = df.columns[0]
        if 'Sales / Pre-Sales / Post-Sales' in first_col:
            df.rename(columns={first_col: 'Category'}, inplace=True)
        
        # Rename status column
        status_col = None
        for col in df.columns:
            if 'Status' in col:
                status_col = col
                break
        
        if status_col:
            df.rename(columns={status_col: 'Status'}, inplace=True)
        
        # Convert Target Date to datetime with DD/MM/YY format handling
        if 'Target Date' in df.columns:
            # Apply date parsing function
            df['Target Date'] = df['Target Date'].apply(parse_dates)
            
        
        # Calculate days remaining
        if 'Target Date' in df.columns:
            df['Days Remaining'] = (df['Target Date'] - pd.Timestamp.now()).dt.days
        
        # Clean status values
        if 'Status' in df.columns:
            df['Status'] = df['Status'].fillna('Not Started').replace('', 'Not Started')
            df['Status'] = df['Status'].str.strip()
            # Standardize status values
            df['Status'] = df['Status'].replace({
                'In progress': 'In Progress',
                'in progress': 'In Progress',
                'completed': 'Completed',
                'Completed': 'Completed',
                'not started': 'Not Started'
            })
        else:
            df['Status'] = 'Not Started'
        
        return df
    
    except Exception as e:
        st.error(f"❌ Error loading from OneDrive: {str(e)}")
        return pd.DataFrame()

# Load data
with st.spinner("🔄 Loading latest data from OneDrive..."):
    df = load_data_from_onedrive()

# Check if data is loaded successfully
if df.empty:
    st.error("Could not load data. Please check your OneDrive link.")
    st.stop()

# Sidebar filters
st.sidebar.markdown("## 🔍 Filters")
st.sidebar.markdown("---")

# Get unique values for filters
categories = df['Category'].unique() if 'Category' in df.columns else []
enablement_areas = df['Enablement Area'].unique() if 'Enablement Area' in df.columns else []
cert_levels = df['Certification Level'].unique() if 'Certification Level' in df.columns else []
engineers = df['Engineer Name'].unique() if 'Engineer Name' in df.columns else []

# Multi-select filters
selected_categories = st.sidebar.multiselect(
    "Sales/Pre-Sales/Post-Sales",
    options=categories,
    default=categories.tolist() if len(categories) > 0 else []
)

selected_areas = st.sidebar.multiselect(
    "Enablement Area",
    options=enablement_areas,
    default=enablement_areas.tolist() if len(enablement_areas) > 0 else []
)

selected_levels = st.sidebar.multiselect(
    "Certification Level",
    options=cert_levels,
    default=cert_levels.tolist() if len(cert_levels) > 0 else []
)

selected_engineers = st.sidebar.multiselect(
    "Engineer Name",
    options=engineers,
    default=[]
)

# Status filter
status_options = ['Not Started', 'In Progress', 'Completed']
selected_status = st.sidebar.multiselect(
    "Status",
    options=status_options,
    default=['Not Started', 'In Progress', 'Completed']
)

# Date range filter - FIXED: Removed the format parameter
if 'Target Date' in df.columns and not df['Target Date'].isna().all():
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📅 Target Date Range")
    
    # Get min and max dates
    valid_dates = df['Target Date'].dropna()
    if not valid_dates.empty:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        
        # Simple date input without format parameter
        date_range = st.sidebar.date_input(
            "Select Range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
        
        # Show hint about expected format
        st.sidebar.caption("Dates are stored in DD/MM/YY format")
    else:
        date_range = None
else:
    date_range = None

# Manual refresh button
st.sidebar.markdown("---")
if st.sidebar.button("🔄 Manual Refresh Now"):
    st.rerun()

# Apply filters
filtered_df = df.copy()

if selected_categories and 'Category' in df.columns:
    filtered_df = filtered_df[filtered_df['Category'].isin(selected_categories)]

if selected_areas and 'Enablement Area' in df.columns:
    filtered_df = filtered_df[filtered_df['Enablement Area'].isin(selected_areas)]

if selected_levels and 'Certification Level' in df.columns:
    filtered_df = filtered_df[filtered_df['Certification Level'].isin(selected_levels)]

if selected_engineers and 'Engineer Name' in df.columns:
    filtered_df = filtered_df[filtered_df['Engineer Name'].isin(selected_engineers)]

if 'Status' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['Status'].isin(selected_status)]

if date_range and len(date_range) == 2 and 'Target Date' in filtered_df.columns:
    start_date, end_date = date_range
    # Convert to datetime for comparison
    start_datetime = pd.Timestamp(start_date)
    end_datetime = pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)  # End of day
    
    filtered_df = filtered_df[
        (filtered_df['Target Date'] >= start_datetime) &
        (filtered_df['Target Date'] <= end_datetime)
    ]

# Main dashboard
st.markdown('<p class="main-header">🎯 VMware Certification Dashboard 2026</p>', unsafe_allow_html=True)
st.markdown("### VMware Certification Status")

# Top KPI metrics - All 8 columns with equal width and increased uniform height
col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)

with col1:
    st.markdown("""
        <div class="metric-card">
            <h4>Resources</h4>
            <h2>{}</h2>
        </div>
    """.format(filtered_df['Engineer Name'].nunique() if 'Engineer Name' in filtered_df.columns else 0), unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div class="metric-card">
            <h4>Total Certs</h4>
            <h2>{}</h2>
        </div>
    """.format(len(filtered_df)), unsafe_allow_html=True)

with col3:
    sales_count = len(filtered_df[filtered_df['Category'] == 'Sales']) if 'Category' in filtered_df.columns else 0
    st.markdown("""
        <div class="metric-card">
            <h4>Sales</h4>
            <h2>{}</h2>
        </div>
    """.format(sales_count), unsafe_allow_html=True)

with col4:
    pre_sales_count = len(filtered_df[filtered_df['Category'] == 'Pre-Sales']) if 'Category' in filtered_df.columns else 0
    st.markdown("""
        <div class="metric-card">
            <h4>Pre-Sales</h4>
            <h2>{}</h2>
        </div>
    """.format(pre_sales_count), unsafe_allow_html=True)

with col5:
    post_sales_count = len(filtered_df[filtered_df['Category'] == 'Post-Sales']) if 'Category' in filtered_df.columns else 0
    st.markdown("""
        <div class="metric-card">
            <h4>Post-Sales</h4>
            <h2>{}</h2>
        </div>
    """.format(post_sales_count), unsafe_allow_html=True)

with col6:
    completed_count = len(filtered_df[filtered_df['Status'] == 'Completed']) if 'Status' in filtered_df.columns else 0
    st.markdown("""
        <div class="metric-card" style="background: linear-gradient(135deg, #10B981 0%, #059669 100%); height: 160px;">
            <h4>Completed</h4>
            <h2>{}</h2>
        </div>
    """.format(completed_count), unsafe_allow_html=True)

with col7:
    in_progress_count = len(filtered_df[filtered_df['Status'] == 'In Progress']) if 'Status' in filtered_df.columns else 0
    st.markdown("""
        <div class="metric-card" style="background: linear-gradient(135deg, #FFFF00 0%, #FDE68A 100%); height: 160px; color: #92400E;">
            <h4 style="color: #92400E;">In Progress</h4>
            <h2 style="color: #92400E;">{}</h2>
        </div>
    """.format(in_progress_count), unsafe_allow_html=True)

with col8:
    not_started_count = len(filtered_df[filtered_df['Status'] == 'Not Started']) if 'Status' in filtered_df.columns else 0
    st.markdown("""
        <div class="metric-card" style="background: linear-gradient(135deg, #EF4444 0%, #DC2626 100%); height: 160px;">
            <h4>Not Started</h4>
            <h2>{}</h2>
        </div>
    """.format(not_started_count), unsafe_allow_html=True)

st.markdown("---")

# Charts section with professional styling
col1, col2 = st.columns(2)

with col1:
    st.markdown('<p class="sub-header">📊 Certifications by Category</p>', unsafe_allow_html=True)
    if 'Category' in filtered_df.columns:
        category_counts = filtered_df['Category'].value_counts().reset_index()
        category_counts.columns = ['Category', 'Count']
        
        # Professional color palette for categories
        colors = {'Sales': '#2E4057',      # Dark blue-gray
                 'Pre-Sales': '#4A6FA5',   # Muted blue
                 'Post-Sales': '#6B4E71'}   # Muted purple
        
        # Create donut chart for more modern look
        fig = px.pie(category_counts, values='Count', names='Category', 
                     title='Distribution by Category',
                     color='Category', 
                     color_discrete_map=colors,
                     hole=0.4)  # Creates donut chart
        
        # Update layout for professional appearance
        fig.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            textfont=dict(size=12, color='white'),
            marker=dict(line=dict(color='white', width=2)),
            hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
        )
        
        fig.update_layout(
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5,
                font=dict(size=11)
            ),
            margin=dict(t=50, b=50, l=20, r=20),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            title=dict(
                text="Distribution by Category",
                font=dict(size=16, color='#1E3A8A'),
                x=0.5,
                xanchor='center'
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Category data not available")

with col2:
    st.markdown('<p class="sub-header">📈 Status Distribution</p>', unsafe_allow_html=True)
    if 'Status' in filtered_df.columns:
        status_counts = filtered_df['Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Count']
        
        # Professional color palette for status
        colors = {'Completed': '#2E7D32',      # Dark green
                 'In Progress': "#F1CF37",      # Warm amber
                 'Not Started': '#D32F2F'}      # Dark red
        
        # Create donut chart for more modern look
        fig = px.pie(status_counts, values='Count', names='Status', 
                     title='Overall Status Distribution',
                     color='Status', 
                     color_discrete_map=colors,
                     hole=0.4)  # Creates donut chart
        
        # Update layout for professional appearance
        fig.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            textfont=dict(size=12, color='white'),
            marker=dict(line=dict(color='white', width=2)),
            hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
        )
        
        fig.update_layout(
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5,
                font=dict(size=11)
            ),
            margin=dict(t=50, b=50, l=20, r=20),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            title=dict(
                text="Overall Status Distribution",
                font=dict(size=16, color='#1E3A8A'),
                x=0.5,
                xanchor='center'
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Status data not available")

# Second row
col1, col2 = st.columns(2)

with col1:
    st.markdown('<p class="sub-header">📊 Enablement Areas</p>', unsafe_allow_html=True)
    if 'Enablement Area' in filtered_df.columns:
        area_counts = filtered_df['Enablement Area'].value_counts().reset_index()
        area_counts.columns = ['Enablement Area', 'Count']
        
        # Professional bar chart styling
        custom_blues = ['#1E3A8A', '#2563EB', '#3B82F6', '#60A5FA', '#93C5FD', '#BFDBFE']
        fig = px.bar(area_counts, x='Enablement Area', y='Count', 
             color='Enablement Area',
             color_discrete_sequence=custom_blues,
             title='Certifications by Area')
        
        fig.update_layout(
            xaxis_title="Enablement Area",
            yaxis_title="Number of Certifications",
            showlegend=False,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            title=dict(
                text="Certifications by Area",
                font=dict(size=16, color='#1E3A8A'),
                x=0.5,
                xanchor='center'
            )
        )
        
        fig.update_traces(
            marker_line_color='white',
            marker_line_width=1,
            opacity=0.8,
            hovertemplate='<b>%{x}</b><br>Count: %{y}<extra></extra>'
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Enablement Area data not available")

with col2:
    st.markdown('<p class="sub-header">📊 Category-wise Status</p>', unsafe_allow_html=True)
    if 'Category' in filtered_df.columns and 'Status' in filtered_df.columns:
        category_status = pd.crosstab(filtered_df['Category'], filtered_df['Status'])
        
        # Professional color palette for status
        status_colors = {'Completed': '#2E7D32', 'In Progress': "#F1CF37", 'Not Started': '#D32F2F'}
        
        fig = px.bar(category_status, barmode='stack', 
                     title='Status Distribution by Category',
                     color_discrete_map=status_colors)
        
        fig.update_layout(
            xaxis_title="Category",
            yaxis_title="Number of Certifications",
            legend_title="Status",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            title=dict(
                text="Status Distribution by Category",
                font=dict(size=16, color='#1E3A8A'),
                x=0.5,
                xanchor='center'
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        
        fig.update_traces(
            marker_line_color='white',
            marker_line_width=1,
            opacity=0.8,
            hovertemplate='<b>%{x}</b><br>Status: %{legend}<br>Count: %{y}<extra></extra>'
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Category or Status data not available")

# Timeline
st.markdown('<p class="sub-header">📅 Certification Timeline</p>', unsafe_allow_html=True)
if 'Target Date' in filtered_df.columns and not filtered_df['Target Date'].isna().all():
    timeline_data = filtered_df.groupby([filtered_df['Target Date'].dt.date, 'Status']).size().reset_index()
    timeline_data.columns = ['Target Date', 'Status', 'Count']
    
    # Professional color palette for status
    status_colors = {'Completed': '#2E7D32', 'In Progress': "#F1CF37", 'Not Started': '#D32F2F'}
    
    fig = px.bar(timeline_data, x='Target Date', y='Count', color='Status',
                  title='Certifications by Target Date',
                  color_discrete_map=status_colors)
    
    fig.update_layout(
        xaxis_title="Target Date",
        yaxis_title="Number of Certifications",
        legend_title="Status",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        title=dict(
            text="Certifications by Target Date",
            font=dict(size=16, color='#1E3A8A'),
            x=0.5,
            xanchor='center'
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        )
    )
    
    fig.update_traces(
        marker_line_color='white',
        marker_line_width=1,
        opacity=0.8,
        hovertemplate='<b>%{x}</b><br>Status: %{legend}<br>Count: %{y}<extra></extra>'
    )
    
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Target Date data not available")

# Detailed table
st.markdown('<p class="sub-header">📋 Detailed Certification Plan</p>', unsafe_allow_html=True)

display_columns = ['Category', 'Enablement Area', 'Certification Level', 'Engineer Name', 
                   'Assigned Certification', 'Target Date', 'Exam Date','Status', 'Remarks']
available_columns = [col for col in display_columns if col in filtered_df.columns]

if available_columns:
    display_df = filtered_df[available_columns].copy()
    
    if 'Target Date' in display_df.columns:
        # Format dates in DD/MM/YY format for display
        display_df['Target Date'] = display_df['Target Date'].dt.strftime('%d/%m/%y')
    
    def color_status(val):
        if val == 'Completed':
            return 'background-color: #E8F5E9; color: #2E7D32'
        elif val == 'In Progress':
            return 'background-color: #FFF3E0; color: #FDB750'
        else:
            return 'background-color: #FFEBEE; color: #D32F2F'
    
    st.dataframe(
        display_df.style.map(color_status, subset=['Status']),
        width='stretch',
        height=500
    )

# Export
st.markdown("---")
col1, col2 = st.columns(2)
with col1:
    if not filtered_df.empty:
        csv = filtered_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Filtered Data (CSV)",
            data=csv,
            file_name=f"vmware_certifications_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

# Engineer Summary
st.markdown('<p class="sub-header">👥 Engineer Summary</p>', unsafe_allow_html=True)
if 'Engineer Name' in filtered_df.columns:
    engineer_summary = filtered_df.groupby('Engineer Name').agg({
        'Category': lambda x: ', '.join(x.unique()) if 'Category' in filtered_df.columns else 'N/A',
        'Assigned Certification': 'count' if 'Assigned Certification' in filtered_df.columns else 'size',
        'Status': [
            ('Completed', lambda x: (x == 'Completed').sum()),
            ('In Progress', lambda x: (x == 'In Progress').sum()),
            ('Not Started', lambda x: (x == 'Not Started').sum())
        ]
    }).reset_index()
    
    engineer_summary.columns = ['Engineer Name', 'Categories', 'Total Certs', 'Completed', 'In Progress', 'Not Started']
    engineer_summary['Completion Rate'] = (engineer_summary['Completed'] / engineer_summary['Total Certs'] * 100).round(1)
    
    st.dataframe(engineer_summary, width='stretch')

# Upcoming deadlines
st.markdown('<p class="sub-header">⏰ Upcoming Deadlines (Next 7 Days)</p>', unsafe_allow_html=True)
if 'Target Date' in filtered_df.columns and 'Status' in filtered_df.columns:
    today = pd.Timestamp.now().normalize()
    next_week = today + pd.Timedelta(days=7)
    
    upcoming = filtered_df[
        (filtered_df['Target Date'] >= today) & 
        (filtered_df['Target Date'] <= next_week) &
        (filtered_df['Status'] != 'Completed')
    ].copy()
    
    if not upcoming.empty:
        # Format dates in DD/MM/YY for display
        upcoming['Target Date'] = upcoming['Target Date'].dt.strftime('%d/%m/%y')
        st.dataframe(
            upcoming[['Engineer Name', 'Category', 'Enablement Area', 'Assigned Certification', 'Target Date', 'Status']],
            width='stretch'
        )
    else:
        st.info("No upcoming deadlines in the next 7 days")

# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: gray; padding: 1rem;'>
        Dashboard last updated: {} | Auto-refreshes every {} minutes | Manual refresh available<br>
        🟢 Completed | 🟡 In Progress | 🔴 Not Started<br>
        📅 Dates are stored in DD/MM/YY format and displayed accordingly
    </div>
""".format(
    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    REFRESH_INTERVAL//60
), unsafe_allow_html=True)

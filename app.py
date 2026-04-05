import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime
import pytz

# =============================================
# PAGE CONFIGURATION & LAYOUT FIXES
# =============================================
st.set_page_config(
    page_title="Demo Operations Analysis", 
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded",
    menu_items=None  # Remove default menu
)

# CSS and JavaScript fixes
st.markdown("""
    <style>
/* Nuke all margins/padding at the top */
html, body, .stApp, .block-container {
    padding-top: 0 !important;
    margin-top: 0 !important;
}

/* Force topmost container to start from top */
main[data-testid="stAppViewContainer"] > div:first-child {
    margin-top: 0 !important;
    padding-top: 0 !important;
}

/* Also force the container that wraps everything */
main .block-container {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}
            
    .custom-title {
    font-size: 2.5rem;
    font-weight: 700;
    color: #ffffff;
    border-bottom: 2px solid #FF8225;
    padding-bottom: 0.5rem;
    margin-top: 0 !important;
    margin-bottom: 0.5rem;
}


            
    </style>
""", unsafe_allow_html=True)

# JavaScript to handle fullscreen behavior
st.components.v1.html("""
<script>
    document.addEventListener('DOMContentLoaded', function() {
        // Fix for fullscreen exit
        window.addEventListener('resize', function() {
            if (!document.fullscreenElement) {
                setTimeout(function() {
                    window.dispatchEvent(new Event('resize'));
                }, 300);
            }
        });
    });
</script>
""", height=0)

# =============================================
# DATA PROCESSING FUNCTIONS
# =============================================
def load_data(uploaded_file):
    """Process the uploaded Excel file"""
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        df["Entry Date"] = pd.to_datetime(df["Entry Date"])
        df["Payment Due Date"] = pd.to_datetime(df["Payment Due Date"])
        
        if 'Currency' in df.columns:
            df['OPEX_USD'] = df.apply(lambda x: 
                x['OPEX'] * 1.17 if x['Currency'] == 'EURO' else 
                x['OPEX'] * 0.78 if x['Currency'] == 'SG Dollar' else 
                x['OPEX'], axis=1)
            
            df['PAID_USD'] = df.apply(lambda x: 
                x['PAID'] * 1.17 if x['Currency'] == 'EURO' else 
                x['PAID'] * 0.78 if x['Currency'] == 'SG Dollar' else 
                x['PAID'], axis=1)
            
            df['Non-Opex_USD'] = df.apply(lambda x: 
                x['Non-Opex'] * 1.17 if x['Currency'] == 'EURO' else 
                x['Non-Opex'] * 0.78 if x['Currency'] == 'SG Dollar' else 
                x['Non-Opex'], axis=1)
            
            df['Balance_USD'] = df.apply(lambda x: 
                x['Balance'] * 1.17 if x['Currency'] == 'EURO' else 
                x['Balance'] * 0.78 if x['Currency'] == 'SG Dollar' else 
                x['Balance'], axis=1)
        else:
            st.warning("No Currency column found - showing original values")
            df['OPEX_USD'] = df['OPEX']
            df['PAID_USD'] = df['PAID']
            df['Non-Opex_USD'] = df['Non-Opex']
            df['Balance_USD'] = df['Balance']
            
        return df
    except Exception as e:
        st.error(f"Failed to process file: {str(e)}")
        return pd.DataFrame()

def parse_date_range(date_str):
    try:
        start_str, end_str = [s.strip() for s in date_str.split("-")]
        start_date = datetime.strptime(start_str, "%Y/%m/%d").date()
        end_date = datetime.strptime(end_str, "%Y/%m/%d").date()
        return start_date, end_date
    except:
        return None

# =============================================
# MAIN APP LAYOUT
# =============================================
st.markdown('<h1 class="custom-title">Demo Vessel Operations Analysis</h1>', unsafe_allow_html=True)
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader(
    "📤 Upload Latest Excel Export", 
    type=["xlsx"],
    help="Export from Niriis Shipping and upload here (.xlsx format only)"
)

# Initialize session state for data persistence
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

# Process uploaded file
if uploaded_file:
    st.session_state.df = load_data(uploaded_file)

# Stop app if no data loaded
if st.session_state.df.empty:
    st.warning("Please upload an Excel file to begin")
    st.stop()

df = st.session_state.df

# =============================================
# SIDEBAR FILTERS
# =============================================
with st.sidebar:
    st.header("🔍 Filter Data")
    
    if st.button("Clear All Filters"):
        st.session_state["filter_vessels"] = []
        st.session_state["filter_suppliers"] = []
        st.session_state["filter_categories"] = []
        st.toast("All filters have been cleared", icon="✅")
        st.rerun()  # Use experimental_rerun to ensure fresh rerun
    
    vessels = st.multiselect(
        "Select Vessel (Consumer)",
        options=df["Consumer"].dropna().unique(),
        placeholder="All Vessels",
        key="filter_vessels"
    )

    suppliers = st.multiselect(
        "Select Supplier",
        options=df["Supplier"].dropna().unique(),
        placeholder="All Suppliers",
        key="filter_suppliers"
    )
    
    categories = st.multiselect(
        "Select Category",
        options=df["Category"].dropna().unique(),
        placeholder="All Categories",
        key="filter_categories"
    )
    
    # Date filters
    today = datetime.now(pytz.timezone('Asia/Dubai')).date()
    
    st.subheader("Entry Date Range")
    default_entry = f"{df['Entry Date'].min().date().strftime('%Y/%m/%d')} - {today.strftime('%Y/%m/%d')}"
    entry_date_text = st.text_input(
        "From - To (YYYY/MM/DD - YYYY/MM/DD)",
        value=default_entry,
        key="entry_date_text"
    )
    entry_dates = parse_date_range(entry_date_text)
    if not entry_dates:
        st.warning("Please use format: YYYY/MM/DD - YYYY/MM/DD")
        st.stop()
    entry_start, entry_end = entry_dates
    
    st.subheader("Payment Due Date Range")
    default_due = f"{df['Payment Due Date'].min().date().strftime('%Y/%m/%d')} - {today.strftime('%Y/%m/%d')}"
    due_date_text = st.text_input(
        "From - To (YYYY/MM/DD - YYYY/MM/DD)", 
        value=default_due,
        key="due_date_text"
    )
    due_dates = parse_date_range(due_date_text)
    if not due_dates:
        st.warning("Please use format: YYYY/MM/DD - YYYY/MM/DD")
        st.stop()
    due_start, due_end = due_dates

    st.markdown("---")

    st.subheader("💱 Conversion Rates")
    st.markdown("""
        - **1 EURO** = **1.17 USD**  
        - **1 SG Dollar** = **0.78 USD**
    """)


# =============================================
# FILTER DATA
# =============================================
filtered_df = df.copy()
if vessels:
    filtered_df = filtered_df[filtered_df["Consumer"].isin(vessels)]
if suppliers:
    filtered_df = filtered_df[filtered_df["Supplier"].isin(suppliers)]
if categories:
    filtered_df = filtered_df[filtered_df["Category"].isin(categories)]
    
filtered_df = filtered_df[
    (filtered_df["Entry Date"].dt.date >= entry_start) &
    (filtered_df["Entry Date"].dt.date <= entry_end) &
    (filtered_df["Payment Due Date"].dt.date >= due_start) &
    (filtered_df["Payment Due Date"].dt.date <= due_end)
]

# Optional Search Filter
search_term = st.text_input("🔍 Search Keyword in Table")

if search_term:
    search_term_lower = search_term.lower()
    filtered_df = filtered_df[
        filtered_df.apply(
            lambda row: row.astype(str).str.lower().str.contains(search_term_lower).any(),
            axis=1
        )
    ]


# =============================================
# SUMMARY METRICS
# =============================================
st.subheader("📌 Summary Metrics")
cols = st.columns(4)

try:
    metrics = [
        ("Total OPEX", filtered_df['OPEX_USD'].sum(), "#4CAF50"),
        ("Total Non-OPEX", filtered_df['Non-Opex_USD'].sum(), "#FF9800"),
        ("Total PAID", filtered_df['PAID_USD'].sum(), "#2196F3"),
        ("Total Balance", filtered_df['Balance_USD'].sum(), "#F44336")
    ]
except KeyError:
    st.warning("USD columns not found - showing original values")
    metrics = [
        ("Total OPEX", filtered_df['OPEX'].sum(), "#4CAF50"),
        ("Total Non-OPEX", filtered_df['Non-Opex'].sum(), "#FF9800"),
        ("Total PAID", filtered_df['PAID'].sum(), "#2196F3"),
        ("Total Balance", filtered_df['Balance'].sum(), "#F44336")
    ]

for col, (label, value, color) in zip(cols, metrics):
    with col:
        st.markdown(f"""
        <div style="background-color:{color}20; padding:15px; border-radius:10px;">
            <h4 style="margin:0; color:{color};">{label}</h4>
            <h2 style="margin:0; color:{color};">${value:,.2f}</h2>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# =============================================
# Vessel Fees Summary
# =============================================

st.subheader("🧾 Vessel Fees Summary")

# Define vessels list
vessels_list = ["OCEAN PRIME", "ATLAS BLUE", "NOVA STAR"]

summary_data = []

for vessel in vessels_list:
    vessel_df = filtered_df[filtered_df["Consumer"] == vessel]

    # Management Fees rows
    mgmt_fees_df = vessel_df[vessel_df["Category"] == "Vsl - Management Fees"]

    # Sum of balances for management fees (including negatives)
    mgmt_fees_balance = mgmt_fees_df["Balance_USD"].sum()

    # Total balance for vessel (including negatives)
    total_balance = vessel_df["Balance_USD"].sum()

    # Supplier fees = total balance - management fees balance
    supplier_fees_balance = total_balance - mgmt_fees_balance

    # Payments made (sum of PAID)
    payments_made = vessel_df["PAID_USD"].sum()

    summary_data.append({
        "Vessel": vessel,
        "Unpaid Management Fees": mgmt_fees_balance,
        "Supplier Fees": supplier_fees_balance,
        "Total Balance": total_balance,
        "Payments Made": payments_made
    })

# Convert to DataFrame for nice display
summary_df = pd.DataFrame(summary_data)

# Format values to 2 decimal places
summary_df_formatted = summary_df.style.format({
    "Unpaid Management Fees": "${:,.2f}",
    "Supplier Fees": "${:,.2f}",
    "Total Balance": "${:,.2f}",
    "Payments Made": "${:,.2f}"
})

summary_df.index = summary_df.index + 1
st.dataframe(summary_df_formatted, use_container_width=True)

# Check total balance sum to verify it matches your summary metrics total
total_balance_sum = summary_df["Total Balance"].sum()
total_payments_made_sum = summary_df["Payments Made"].sum()

col1, col2 = st.columns(2)

with col1:
    st.caption(f"Sum of Total Balances across vessels: ${total_balance_sum:,.2f}")

with col2:
    st.caption(f"Sum of Payments Made across vessels: ${total_payments_made_sum:,.2f}")

st.markdown("---")

# =============================================
# TODAY'S TRANSACTIONS SECTION
# =============================================

# Get today's date in UAE timezone
uae_today = datetime.now(pytz.timezone('Asia/Dubai')).date()

# Filter today's transactions
today_transactions = df[df['Entry Date'].dt.date == uae_today]

if not today_transactions.empty:
    st.subheader(f"📅 Today's Transactions ({uae_today.strftime('%d %b %Y')})")
    
    # Create a nice card display
    cols = st.columns(3)
    with cols[0]:
        st.metric("Total Transactions", len(today_transactions))
    with cols[1]:
        st.metric("Total Amount", f"${today_transactions['OPEX_USD'].sum():,.2f}")
    
    # Display the transactions in a nice table
    st.dataframe(
        today_transactions[[
            'ID','Consumer', 'Supplier', 'Category', 
            'OPEX_USD', 'Payment Due Date', 'Description'
        ]].rename(columns={
            'Consumer': 'Vessel',
            'OPEX_USD': 'Amount (USD)',
            'Payment Due Date': 'Due Date'
        }),
        column_config={
            "Amount (USD)": st.column_config.NumberColumn(
                format="$%.2f",
            ),
            "Due Date": st.column_config.DateColumn(
                format="DD MMM YYYY",
            )
        },
        use_container_width=True,
        height=min(300, 35 * len(today_transactions) + 3),  # Dynamic height
        hide_index=True
    )
else:
    st.subheader(f"📅 Today's Transactions ({uae_today.strftime('%d %b %Y')})")
    st.info("No transactions recorded for today")
    
st.markdown("---")

# =============================================
# OVERDUE SECTION
# =============================================

overdue_df = filtered_df[
    (filtered_df['Payment Due Date'].dt.date <= uae_today) &
    (filtered_df['Balance_USD'] > 0)
]

if not overdue_df.empty:
    st.subheader("⚠️ Overdue Payments")
    
    # Summary
    cols = st.columns(3)
    with cols[0]:
        st.metric("Overdue Invoices", len(overdue_df))
    with cols[1]:
        st.metric("Total Overdue", f"${overdue_df['Balance_USD'].sum():,.2f}")
    
    # Display Table
    st.dataframe(
        overdue_df[[
            'ID','Consumer', 'Supplier', 'Category',
            'Balance_USD', 'Payment Due Date', 'Description'
        ]].rename(columns={
            'Consumer': 'Vessel',
            'Balance_USD': 'Outstanding Amount (USD)',
            'Payment Due Date': 'Due Date'
        }),
        column_config={
            "Outstanding Amount (USD)": st.column_config.NumberColumn(format="$%.2f"),
            "Due Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        },
        use_container_width=True,
        hide_index=True,
        height=min(300, 35 * len(overdue_df) + 3)
    )
else:
    st.subheader("⚠️ Overdue Payments")
    st.info("No overdue payments found. ✅")

st.markdown("---")

# =============================================
# OVERPAYMENTS SECTION
# =============================================

overpaid_df = filtered_df[filtered_df['Balance_USD'] < 0]

if not overpaid_df.empty:
    st.subheader("💸 Overpaid Transactions")
    
    cols = st.columns(3)
    with cols[0]:
        st.metric("Count", len(overpaid_df))
    with cols[1]:
        st.metric("Total Overpaid", f"${overpaid_df['Balance_USD'].sum():,.2f}")
    
    st.dataframe(
        overpaid_df[[
            'ID','Consumer', 'Supplier', 'Category',
            'Balance_USD', 'Payment Due Date', 'Description'
        ]].rename(columns={
            'Consumer': 'Vessel',
            'Balance_USD': 'Overpaid Amount (USD)',
            'Payment Due Date': 'Due Date'
        }),
        column_config={
            "Overpaid Amount (USD)": st.column_config.NumberColumn(format="$%.2f"),
            "Due Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        },
        use_container_width=True,
        hide_index=True,
        height=min(300, 35 * len(overpaid_df) + 3)
    )
else:
    st.subheader("💸 Overpaid Transactions")
    st.info("No overpaid records found.")

st.markdown("---")

# =============================================
# CHARTS SECTION
# =============================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📑 Management Report",
    "📊 Category Analysis", 
    "📈 Monthly Trends", 
    "🏆 Top 10 Suppliers",
    "📂 Supplier Category Breakdown",
    "🍰 Expense Distribution", 
    "📋 Data Table"
])

with tab1:
    st.subheader("Management Report")

    # Format dates for display
    mgmt_df = filtered_df.copy()

    mgmt_df['Payment Due Date'] = mgmt_df['Payment Due Date'].dt.strftime('%d %b %Y')
    mgmt_df['Invoice Date'] = mgmt_df['Invoice Date'].dt.strftime('%d %b %Y') if 'Invoice Date' in mgmt_df.columns else None

    # Select only the requested columns
    mgmt_columns = [
        'ID', 'Consumer', 'Supplier', 'Description',
        'Grand Category', 'Category',
        'Invoice #', 'Invoice Date', 'Payment Due Date',
        'Currency', 'Balance', 'Balance_USD'
    ]

    # Some columns may be missing — filter out
    mgmt_columns = [col for col in mgmt_columns if col in mgmt_df.columns]

    st.dataframe(
        mgmt_df[mgmt_columns],
        use_container_width=True,
        height=500,
        hide_index=True
    )

    # Optional: Export the management report too
    mgmt_output = BytesIO()
    with pd.ExcelWriter(mgmt_output, engine='openpyxl') as writer:
        mgmt_df[mgmt_columns].to_excel(writer, index=False)
    mgmt_output.seek(0)

    st.download_button(
        "📥 Download Management Report as Excel",
        data=mgmt_output.getvalue(),
        file_name="management_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with tab2:
    st.subheader("Total Expenses (OPEX + Non-OPEX) by Vessel")
    try:
        filtered_df['TOTAL_EXPENSES'] = filtered_df['OPEX_USD'] + filtered_df['Non-Opex_USD']
        vessel_summary = filtered_df.groupby(['Grand Category', 'Consumer'])['TOTAL_EXPENSES'].sum().reset_index()
        
        fig1 = px.bar(
            vessel_summary,
            x='Grand Category',
            y='TOTAL_EXPENSES',
            color='Consumer',
            color_discrete_sequence=["#1f77b4", "#FF8225", "#F8EDED"],  # colorhunt.co/palette
            barmode='group',
            height=500,
            title='<b>Total Expenses by Vessel</b>',
            labels={'TOTAL_EXPENSES': 'Amount (USD)', 'Grand Category': 'Category Group'},
            category_orders={"Consumer": ["OCEAN PRIME", "ATLAS BLUE", "NOVA STAR"]}
        )
        
        fig1.update_layout(
            yaxis_tickprefix='$',
            yaxis_tickformat=',.0f',
            hovermode='x unified',
            xaxis={'categoryorder': 'total descending'},
            legend_title_text='Vessel'
        )
        fig1.update_xaxes(tickangle=45)
        st.plotly_chart(fig1, use_container_width=True)
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")

with tab3:
    st.subheader("Monthly Expenses by Vessel")
    
    try:
        # Prepare monthly data
        monthly_df = filtered_df.copy()
        monthly_df['Month'] = monthly_df['Entry Date'].dt.to_period('M').dt.to_timestamp()
        monthly_df['Month_Str'] = monthly_df['Entry Date'].dt.strftime('%b %Y')  # Format: "Jun 2023"
        monthly_df['Total_Expenses_USD'] = monthly_df['OPEX_USD'] + monthly_df['Non-Opex_USD']
        
        # Group by Month and Vessel
        vessel_monthly = monthly_df.groupby(['Month', 'Month_Str', 'Consumer'])['Total_Expenses_USD'].sum().reset_index()
        
        # Create scrollable container
        container = st.container()
        
        # Dynamic height based on number of months
        months_count = len(vessel_monthly['Month'].unique())
        chart_height = max(400, months_count * 25)  # 25px per month
        
        with container:
            # Grouped bar chart
            fig = px.bar(
                vessel_monthly,
                x='Month_Str',
                y='Total_Expenses_USD',
                color='Consumer',
                barmode='group',
                height=chart_height,
                title='<b>Monthly Expenses by Vessel</b>',
                labels={
                    'Total_Expenses_USD': 'Total Expenses (USD)',
                    'Month_Str': 'Month',
                    'Consumer': 'Vessel'
                },
                color_discrete_sequence=["#1f77b4", "#FF8225", "#F8EDED"]
            )
            
            fig.update_layout(
                yaxis_tickprefix='$',
                yaxis_tickformat=',.0f',
                hovermode='x unified',
                xaxis_title='Month',
                yaxis_title='Total Expenses (USD)',
                legend_title_text='Vessel',
                xaxis={'tickangle': 45}
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
        # Month range indicator
        min_month = vessel_monthly['Month'].min().strftime('%b %Y')
        max_month = vessel_monthly['Month'].max().strftime('%b %Y')
        st.caption(f"Showing data from {min_month} to {max_month} | Scroll → to see more months")
        
    except Exception as e:
        st.error(f"Error creating monthly chart: {str(e)}")

with tab4:
    st.subheader("🏆 Top 10 Suppliers by Total Expenses")
    
    try:
        # Ensure TOTAL_EXPENSES exists
        if 'TOTAL_EXPENSES' not in filtered_df.columns:
            filtered_df['TOTAL_EXPENSES'] = filtered_df['OPEX_USD'] + filtered_df['Non-Opex_USD']
        
        # Get top 10 suppliers
        top_10_suppliers = (
            filtered_df.groupby('Supplier')['TOTAL_EXPENSES']
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        
        fig_top_suppliers = px.bar(
            top_10_suppliers,
            x='Supplier',
            y='TOTAL_EXPENSES',
            text='TOTAL_EXPENSES',
            color='Supplier',
            title='Top 10 Suppliers by Total Expenses (USD)',
            color_discrete_sequence=px.colors.qualitative.Set2,
            labels={'TOTAL_EXPENSES': 'Total Expenses (USD)'}
        )
        
        fig_top_suppliers.update_traces(
            texttemplate='$%{text:,.0f}',
            textposition='outside',
            cliponaxis=False
        )

        
        fig_top_suppliers.update_layout(
            yaxis_tickprefix='$',
            yaxis_tickformat=',.0f',
            xaxis_title='Supplier',
            yaxis_title='Total Expenses (USD)',
            xaxis_tickangle=45,
            uniformtext_minsize=8,
            uniformtext_mode='hide'
        )
        
        st.plotly_chart(fig_top_suppliers, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error creating supplier chart: {str(e)}")

with tab5:
    st.subheader("Supplier Category Breakdown")

    # Make sure at least 1 supplier is selected
    if not suppliers:
        st.info("Please select a supplier from the sidebar to see the category breakdown.")
        pass
    else:

    # If multiple suppliers are selected, you can show all or force single supplier
    #if len(suppliers) > 1:
    #    st.info("Please select only **one supplier** to see a clear breakdown.")
    #    st.stop()
    
    # Filtered already by global sidebar filter
        supplier_df = filtered_df.copy()
            
        category_summary = (
            supplier_df.groupby('Category')['TOTAL_EXPENSES']
            .sum()
            .reset_index()
            .sort_values(by='TOTAL_EXPENSES', ascending=False)
        )

        # ✅ Smarter title if many suppliers
        if len(suppliers) == 1:
            title_supplier = suppliers[0]
        elif len(suppliers) <= 3:
            title_supplier = ", ".join(suppliers)
        else:
            title_supplier = f"{len(suppliers)} suppliers selected"

        title_text = f"Expenses Breakdown for {title_supplier}"

        
        fig = px.bar(
            category_summary,
            x='Category',
            y='TOTAL_EXPENSES',
            text='TOTAL_EXPENSES',
            title=title_text,  # Use dynamic title!
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        
        fig.update_traces(texttemplate='$%{text:,.2f}', textposition='outside',cliponaxis=False)
        fig.update_layout(yaxis_tickprefix='$', yaxis_tickformat=',.0f', xaxis_tickangle=45)
        
        st.plotly_chart(fig, use_container_width=True)
        #st.dataframe(category_summary)  

with tab6:
    st.subheader("OPEX vs Non-OPEX Distribution")
    try:
        opex_total = filtered_df["OPEX_USD"].sum()
        non_opex_total = filtered_df["Non-Opex_USD"].sum()
    except KeyError:
        opex_total = filtered_df["OPEX"].sum()
        non_opex_total = filtered_df["Non-Opex"].sum()
    
    fig2 = px.pie(
        names=["OPEX", "Non-OPEX"],
        values=[opex_total, non_opex_total],
        hole=0.3,
        height=500,
        color_discrete_sequence=["#4CAF50", "#FF9800"]  # Green and Orange
    )
    fig2.update_traces(
        textposition='inside',
        textinfo='percent+label+value',
        hovertemplate="<b>%{label}</b><br>Amount: $%{value:,.2f}<br>Percentage: %{percent}"
    )
    st.plotly_chart(fig2, use_container_width=True)

with tab7:
    st.subheader("Filtered Data")

    # Format Dates
    formatted_df = filtered_df.copy()
    formatted_df['Entry Date'] = formatted_df['Entry Date'].dt.strftime('%d %b %Y')
    formatted_df['Payment Due Date'] = formatted_df['Payment Due Date'].dt.strftime('%d %b %Y')

    if 'Invoice Date' in formatted_df.columns:
        formatted_df['Invoice Date'] = pd.to_datetime(formatted_df['Invoice Date'], errors='coerce').dt.strftime('%d %b %Y')

    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=500,
        hide_index=True
    )
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False)
    output.seek(0)
    
    st.download_button(
        "📥 Download Filtered Data as Excel",
        data=output.getvalue(),
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#st.markdown("---")

# =============================================
# FOOTER
# =============================================
st.markdown("---")
uae_time = datetime.now(pytz.timezone('Asia/Dubai'))
st.caption(f"UAE Time: {uae_time.strftime('%Y-%m-%d %H:%M')} | Developed by Abdelrahman Elshishiny")

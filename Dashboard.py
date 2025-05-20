import calendar
import smtplib
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Tumble Cup Admin", page_icon="📊", layout="wide")

conn = st.connection("gsheets", type=GSheetsConnection)

month_list = list(calendar.month_name)[1:]
current_month = datetime.today().month
current_month_name = calendar.month_name[current_month]
current_year = datetime.today().year

st.markdown("""
<style>
.st-emotion-cache-1weic72 {
display: none;
}
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Define product costs
PRODUCT_COSTS = {
    "Classic Tumbler": 1850,
    "Can Glass": 1250,
    "Coffee Mug": 1500
}


def get_orders(month_number=None):
    orders_df = conn.read(worksheet="Tumble_cup", ttl=0)

    if orders_df.empty:
        return pd.DataFrame()
    orders_df["Order Date"] = pd.to_datetime(orders_df["Order Date"], errors="coerce")

    if month_number:
        orders_df = orders_df[orders_df["Order Date"].dt.month == month_number]

    return orders_df


def update_order_status(order_id, new_status):
    orders_df = conn.read(worksheet="Tumble_cup", ttl=0)
    if order_id in orders_df["ID"].values:
        orders_df.loc[orders_df["ID"] == order_id, "Status"] = new_status
        conn.update(worksheet="Tumble_cup", data=orders_df)
        return True
    return False


def update_payment_status(order_id, new_status):
    orders_df = conn.read(worksheet="Tumble_cup", ttl=0)
    if order_id in orders_df["ID"].values:
        orders_df.loc[orders_df["ID"] == order_id, "Payment Status"] = new_status
        conn.update(worksheet="Tumble_cup", data=orders_df)
        return True
    return False


def delete_order(order_id):
    orders_df = conn.read(worksheet="Tumble_cup", ttl=0)
    if order_id in orders_df["ID"].values:
        orders_df = orders_df[orders_df["ID"] != order_id]
        conn.update(worksheet="Tumble_cup", data=orders_df)
        return True
    return False


def send_email_notification(to_email, subject, content):
    gmail_user = "teamtumblecup@gmail.com"
    try:
        app_password = st.secrets["Email"]["Password"]

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = gmail_user
        msg['To'] = to_email
        msg.set_content("This is a plain text version of the email")
        msg.add_alternative(content, subtype='html')

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(gmail_user, app_password)
            smtp.send_message(msg)
            return True
    except Exception as e:
        st.error(f"Failed to send email: {e}")
        return False


def calculate_sales_metrics(orders_df):
    """Calculate sales metrics including costs and profits"""
    if 'Item Name' not in orders_df.columns or 'Base Price' not in orders_df.columns:
        return {
            "total_sales": 0,
            "total_costs": 0,
            "total_profit": 0,
            "product_breakdown": pd.DataFrame()
        }

    orders_df['Base Price'] = pd.to_numeric(orders_df['Base Price'], errors='coerce')

    orders_df['Total Cost'] = orders_df['Item Name'].map(PRODUCT_COSTS) * orders_df['Item Quantity']

    orders_df['Profit'] = orders_df['Total'] - orders_df['Total Cost']

    total_sales = orders_df['Price'].sum()
    total_costs = orders_df['Total Cost'].sum()
    total_profit = orders_df['Profit'].sum()

    product_breakdown = orders_df.groupby('Item Name').agg({
        'Item Quantity': 'sum',
        'Base Price': 'sum',
        'Total Cost': 'sum',
        'Profit': 'sum'
    }).rename(columns={'Item Quantity': 'Total Quantity'}).reset_index()

    return {
        "total_sales": total_sales,
        "total_costs": total_costs,
        "total_profit": total_profit,
        "product_breakdown": product_breakdown
    }


# Main title for the app
st.title("Tumble Cup Admin Dashboard")

# Create tabs
tab1, tab2 = st.tabs(["Status Update", "Analytics"])

# Common month selection for both tabs
with st.sidebar:
    st.header("Filter Options")
    selected_month = st.selectbox(
        "Select Month",
        month_list,
        index=current_month - 1,
        placeholder="Select Month"
    )

    selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
    orders_df = get_orders(selected_month_number)

    if not orders_df.empty:
        orders_df["Order Date"] = orders_df["Order Date"].dt.strftime("%d-%B-%Y")

        search_term = st.text_input("Search by Name or Order Number", placeholder="Enter Search Term",
                                    key="search_term")

        if search_term:
            orders_df = orders_df[orders_df['Name'].str.contains(search_term, case=False) |
                                  orders_df['Order Number'].str.contains(search_term, case=False)]

            if orders_df.empty:
                st.warning("No such orders found!")

        status_filter = st.multiselect("Filter by Status", options=orders_df['Status'].unique().tolist(),
                                       default=orders_df['Status'].unique().tolist())
        payment_filter = st.multiselect("Filter by Payment Status",
                                        options=orders_df['Payment Status'].unique().tolist(),
                                        default=orders_df['Payment Status'].unique().tolist())

        filtered_df = orders_df[orders_df['Status'].isin(status_filter) &
                                orders_df['Payment Status'].isin(payment_filter)]

# Tab 1: Status Update
with tab1:
    st.header("Order Status Management")

    if not orders_df.empty:
        st.dataframe(filtered_df)

        st.subheader("Update Order & Payment Status")
        col1, col2, col3 = st.columns(3)

        with col1:
            order_id = st.number_input("Order ID", min_value=int(orders_df['ID'].min()),
                                       max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
                                       step=1, key="status_order_id")
            payment_order_id = st.number_input("Payment Order ID", min_value=int(orders_df['ID'].min()),
                                               max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
                                               step=1, key="payment_order_id")
            delete_order_id = st.number_input("Delete Order ID", min_value=int(orders_df['ID'].min()),
                                              max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
                                              step=1, key="delete_order_id")

        with col2:
            new_status = st.selectbox("New Status",
                                      ["Pending", "Processing", "Shipped", "Delivered", "Cancelled"],
                                      key="new_status")
            payment_new_status = st.selectbox("Payment New Status",
                                              ["Pending", "Processing", "Confirmed", "Cancelled"],
                                              key="payment_new_status")

        with col3:
            if st.button("Update Status", key="update_status_btn"):
                if update_order_status(order_id, new_status):
                    st.success(f"Order #{order_id} status updated to {new_status}")
                    customer_email = orders_df.loc[orders_df["ID"] == order_id, "Email"].values[0]
                    email_content = f"Your order #{order_id} status has been updated to {new_status}."
                    send_email_notification(customer_email, "Tumble Cup Order Status Update", email_content)
                    st.rerun()
                else:
                    st.error(f"Failed to update order #{order_id}")

            if st.button("Update Payment Status", key="update_payment_btn"):
                if update_payment_status(payment_order_id, payment_new_status):
                    st.success(f"Order #{payment_order_id} payment status updated to {payment_new_status}")
                    # Send email notification to customer
                    customer_email = orders_df.loc[orders_df["ID"] == payment_order_id, "Email"].values[0]
                    email_content = f"Your order #{payment_order_id} payment status has been updated to {payment_new_status}."
                    send_email_notification(customer_email, "Tumble Cup Payment Status Update", email_content)
                    st.rerun()
                else:
                    st.error(f"Failed to update payment status for order #{payment_order_id}")

            if st.button("Delete Order", key="delete_order_btn"):
                if delete_order(delete_order_id):
                    st.success(f"Order #{delete_order_id} has been deleted")
                    st.rerun()
                else:
                    st.error(f"Failed to delete order #{delete_order_id}")

        # Export to CSV button
        if st.button("Export Orders to CSV", key="export_csv_status"):
            csv = filtered_df.to_csv(index=False)
            st.download_button(
                label="Download CSV",
                data=csv,
                file_name=f"tumble_cup_orders_{datetime.today().strftime('%Y-%m-%d')}.csv",
                mime="text/csv"
            )
    else:
        st.info("No orders found for the selected month.")

# Tab 2: Analytics
with tab2:
    st.header("Sales Analytics")

    if not orders_df.empty:
        # Sales Metrics Section
        st.subheader("Sales Metrics")

        # Only calculate metrics for confirmed payments
        confirmed_orders = filtered_df[filtered_df['Payment Status'] == 'Confirmed'].copy()
        metrics = calculate_sales_metrics(confirmed_orders)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Sales (PKR)", f"{metrics['total_sales']:,.2f}")
        with col2:
            st.metric("Total Costs (PKR)", f"{metrics['total_costs']:,.2f}")
        with col3:
            st.metric("Total Profit (PKR)", f"{metrics['total_profit']:,.2f}")

        st.subheader("Product Sales Breakdown")
        if not metrics['product_breakdown'].empty:
            metrics_df = metrics['product_breakdown']
            metrics_df.index = range(1, len(metrics_df) + 1)
            st.dataframe(metrics_df)

            st.subheader("Product Sales Comparison")
            chart_data = metrics['product_breakdown'].set_index('Item Name')
            st.bar_chart(chart_data[['Total Quantity']])

            st.subheader("Profit by Product")
            st.bar_chart(chart_data[['Profit']])

            # Additional analytics
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Order Status Distribution")
                status_counts = filtered_df['Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                st.bar_chart(status_counts.set_index('Status'))

            with col2:
                st.subheader("Payment Status Distribution")
                payment_counts = filtered_df['Payment Status'].value_counts().reset_index()
                payment_counts.columns = ['Payment Status', 'Count']
                st.bar_chart(payment_counts.set_index('Payment Status'))

            # Cost vs Profit Ratio Pie Chart
            st.subheader("Cost vs Profit Ratio")
            if metrics['total_sales'] > 0:
                # Create pie chart data
                pie_data = {
                    'Category': ['Cost', 'Profit'],
                    'Value': [metrics['total_costs'], metrics['total_profit']]
                }
                pie_df = pd.DataFrame(pie_data)

                # Calculate percentages for display
                cost_percentage = (metrics['total_costs'] / metrics['total_sales']) * 100
                profit_percentage = (metrics['total_profit'] / metrics['total_sales']) * 100

                col1, col2 = st.columns([3, 2])

                with col1:
                    # Display pie chart
                    fig = {
                        'data': [{
                            'values': pie_df['Value'],
                            'labels': pie_df['Category'],
                            'type': 'pie',
                            'hole': 0.4,
                            'marker': {'colors': ['#FF6B6B', '#4CAF50']}
                        }],
                        'layout': {'title': 'Cost vs Profit Distribution'}
                    }
                    st.plotly_chart(fig, use_container_width=True)

                with col2:
                    # Display metrics with percentages
                    st.metric("Cost Percentage", f"{cost_percentage:.2f}%")
                    st.metric("Profit Percentage", f"{profit_percentage:.2f}%")
                    st.metric("Profit Margin", f"{(profit_percentage):.2f}%")
            else:
                st.info("No sales data available to calculate cost vs profit ratio.")

            # Export to CSV button for analytics data
            if st.button("Export Analytics to CSV", key="export_csv_analytics"):
                analytics_csv = metrics_df.to_csv(index=False)
                st.download_button(
                    label="Download Analytics CSV",
                    data=analytics_csv,
                    file_name=f"tumble_cup_analytics_{datetime.today().strftime('%Y-%m-%d')}.csv",
                    mime="text/csv"
                )
        else:
            st.info("No confirmed orders to show product breakdown.")
    else:
        st.info("No orders found for the selected month.")
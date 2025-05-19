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
    # #MainMenu {visibility: hidden;}
    # header {visibility: hidden;}
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
    try:
        msg = EmailMessage()
        msg.set_content(content)
        msg['Subject'] = subject
        msg['From'] = st.secrets["Email"]["sender"]
        msg['To'] = to_email

        server = smtplib.SMTP(st.secrets["Email"]["smtp_server"], st.secrets["Email"]["smtp_port"])
        server.starttls()
        server.login(st.secrets["Email"]["username"], st.secrets["Email"]["password"])
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Failed to send email: {str(e)}")
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

    # Ensure price is numeric
    orders_df['Base Price'] = pd.to_numeric(orders_df['Base Price'], errors='coerce')

    # Add cost column based on product type
    orders_df['Cost'] = orders_df['Item Name'].map(PRODUCT_COSTS)

    # Calculate profit for each item
    orders_df['Profit'] = orders_df['Base Price'] - orders_df['Cost']

    # Calculate totals
    total_sales = orders_df['Base Price'].sum()
    total_costs = orders_df['Cost'].sum()
    total_profit = orders_df['Profit'].sum()

    # Product breakdown
    product_breakdown = orders_df.groupby('Item Name').agg({
        'ID': 'count',
        'Base Price': 'sum',
        'Cost': 'sum',
        'Profit': 'sum'
    }).rename(columns={'ID': 'Count'}).reset_index()

    return {
        "total_sales": total_sales,
        "total_costs": total_costs,
        "total_profit": total_profit,
        "product_breakdown": product_breakdown
    }


with st.container():
    st.header("Admin Panel")
    admin_password = st.text_input("Enter Admin Password", type="password")
    pwd = st.secrets["Password"]["Password"]

    if admin_password == str(pwd):
        st.success("Admin authenticated!")

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
                else:

                    status_filter = st.multiselect("Filter by Status", options=orders_df['Status'].unique().tolist(),
                                                   default=orders_df['Status'].unique().tolist())
                    payment_filter = st.multiselect("Filter by Payment Status",
                                                    options=orders_df['Payment Status'].unique().tolist(),
                                                    default=orders_df['Payment Status'].unique().tolist())

                    filtered_df = orders_df[orders_df['Status'].isin(status_filter) &
                                            orders_df['Payment Status'].isin(payment_filter)]

                    st.dataframe(filtered_df)

                    st.subheader("Update Order Status")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        order_id = st.number_input("Order ID", min_value=1,
                                                   max_value=int(orders_df['ID'].max()),
                                                   step=1)
                    with col2:
                        new_status = st.selectbox("New Status",
                                                  ["Pending", "Processing", "Shipped", "Delivered", "Cancelled"])
                    with col3:
                        if st.button("Update Status"):
                            if update_order_status(order_id, new_status):
                                st.success(f"Order #{order_id} status updated to {new_status}")
                                # Send email notification to customer
                                customer_email = orders_df.loc[orders_df["ID"] == order_id, "Email"].values[0]
                                email_content = f"Your order #{order_id} status has been updated to {new_status}."
                                send_email_notification(customer_email, "Tumble Cup Order Status Update", email_content)
                                st.rerun()
                            else:
                                st.error(f"Failed to update order #{order_id}")

                    # Export to CSV button
                    if st.button("Export Orders to CSV"):
                        csv = filtered_df.to_csv(index=False)
                        st.download_button(
                            label="Download CSV",
                            data=csv,
                            file_name=f"tumble_cup_orders_{datetime.today().strftime('%Y-%m-%d')}.csv",
                            mime="text/csv"
                        )
            else:
                # Show all orders with filters
                status_filter = st.multiselect("Filter by Status", options=orders_df['Status'].unique().tolist(),
                                               default=orders_df['Status'].unique().tolist())
                payment_filter = st.multiselect("Filter by Payment Status",
                                                options=orders_df['Payment Status'].unique().tolist(),
                                                default=orders_df['Payment Status'].unique().tolist())

                filtered_df = orders_df[orders_df['Status'].isin(status_filter) &
                                        orders_df['Payment Status'].isin(payment_filter)]

                st.dataframe(filtered_df)

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
                    st.bar_chart(chart_data[['Count']])

                    st.subheader("Profit by Product")
                    st.bar_chart(chart_data[['Profit']])
                else:
                    st.info("No confirmed orders to show product breakdown.")

                st.subheader("Update Order & Payment Status")
                col1, col2, col3 = st.columns(3)

                with col1:
                    order_id = st.number_input("Order ID", min_value=1,
                                               max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
                                               step=1)
                    payment_order_id = st.number_input("Payment Order ID", min_value=1,
                                                       max_value=int(
                                                           orders_df['ID'].max()) if not orders_df.empty else 1,
                                                       step=1)
                    delete_order_id = st.number_input("Delete Order ID", min_value=1,
                                                      max_value=int(
                                                          orders_df['ID'].max()) if not orders_df.empty else 1,
                                                      step=1)

                with col2:
                    new_status = st.selectbox("New Status",
                                              ["Pending", "Processing", "Shipped", "Delivered", "Cancelled"])
                    payment_new_status = st.selectbox("Payment New Status",
                                                      ["Pending", "Processing", "Confirmed", "Cancelled"])

                with col3:
                    if st.button("Update Status"):
                        if update_order_status(order_id, new_status):
                            st.success(f"Order #{order_id} status updated to {new_status}")
                            # customer_email = orders_df.loc[orders_df["ID"] == order_id, "Emai"].values[0]
                            email_content = f"Your order #{order_id} status has been updated to {new_status}."
                            # send_email_notification(customer_email, "Tumble Cup Order Status Update", email_content)
                            st.rerun()
                        else:
                            st.error(f"Failed to update order #{order_id}")

                    if st.button("Update Payment Status"):
                        if update_payment_status(payment_order_id, payment_new_status):
                            st.success(f"Order #{payment_order_id} payment status updated to {payment_new_status}")
                            # Send email notification to customer
                            customer_email = orders_df.loc[orders_df["ID"] == payment_order_id, "Email"].values[0]
                            email_content = f"Your order #{payment_order_id} payment status has been updated to {payment_new_status}."
                            # send_email_notification(customer_email, "Tumble Cup Payment Status Update", email_content)
                            st.rerun()
                        else:
                            st.error(f"Failed to update payment status for order #{payment_order_id}")

                    if st.button("Delete Order"):
                        if delete_order(delete_order_id):
                            st.success(f"Order #{delete_order_id} has been deleted")
                            st.rerun()
                        else:
                            st.error(f"Failed to delete order #{delete_order_id}")

                # Export to CSV button
                if st.button("Export Orders to CSV"):
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(
                        label="Download CSV",
                        data=csv,
                        file_name=f"tumble_cup_orders_{datetime.today().strftime('%Y-%m-%d')}.csv",
                        mime="text/csv"
                    )
        else:
            st.info("No orders found for the selected month.")
    elif admin_password:
        st.error("Incorrect password")
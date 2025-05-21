import calendar
import smtplib
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import streamlit as st
from PIL import Image
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


def update_by_order_number(order_number, new_status, status_field="Status", tracking_id=None, partner=None):
    """
    Update all orders with matching order number

    Args:
        order_number: Order number to match (e.g., "TC00001")
        new_status: New status to set
        status_field: Column to update ("Status" or "Payment Status")
        tracking_id: Optional tracking ID to add when shipping

    Returns:
        tuple: (success_count, list of updated order IDs)
    """
    orders_df = conn.read(worksheet="Tumble_cup", ttl=0)

    matching_orders = orders_df[orders_df["Order Number"].astype(str).str.contains(order_number, case=False)]

    if matching_orders.empty:
        return 0, []
    matching_ids = matching_orders["ID"].tolist()
    orders_df.loc[orders_df["ID"].isin(matching_ids), status_field] = new_status

    if tracking_id and status_field == "Status" and new_status == "Shipped" and partner:
        if 'Tracking ID' not in orders_df.columns:
            orders_df['Tracking ID'] = ""
        orders_df.loc[orders_df["ID"].isin(matching_ids), "Tracking ID"] = tracking_id
        orders_df.loc[orders_df["ID"].isin(matching_ids), "Tracking Partner"] = partner

    conn.update(worksheet="Tumble_cup", data=orders_df)

    return len(matching_ids), matching_ids


def delete_order(order_id):
    """
    Delete all orders with matching order number
    :param order_id:
    :return: None

    """
    orders_df = conn.read(worksheet="Tumble_cup", ttl=0)
    if order_id in orders_df["ID"].values:
        orders_df = orders_df[orders_df["ID"] != order_id]
        conn.update(worksheet="Tumble_cup", data=orders_df)
        return True
    return False


def send_email_notification(to_email, subject, content):
    """
    Send email notification
    :param to_email: Clients email address
    :param subject: Subject
    :param content: Email content
    :return: None
    """
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
            "product_breakdown": pd.DataFrame(),
            "style_breakdown": pd.DataFrame()
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

    style_breakdown = None
    if 'Item Style' in orders_df.columns:

        style_df = orders_df.copy()
        if style_df['Item Style'].isna().any():
            style_df['Item Style'] = style_df['Item Style'].fillna('Regular')

        custom_styles = style_df[style_df['Item Style'].str.contains('Custom|Hand painted|Handpainted',
                                                                     case=False, na=False)]

        if not custom_styles.empty:
            style_breakdown = custom_styles.groupby(['Item Style', 'Item Name']).agg({
                'Item Quantity': 'sum',
                'Base Price': 'sum',
                'Total Cost': 'sum',
                'Total': 'sum',
                'Profit': 'sum'
            }).rename(columns={'Item Quantity': 'Total Quantity'}).reset_index()

    return {
        "total_sales": total_sales,
        "total_costs": total_costs,
        "total_profit": total_profit,
        "product_breakdown": product_breakdown,
        "style_breakdown": style_breakdown if style_breakdown is not None else pd.DataFrame()
    }


st.markdown("<h1 style='text-align: center;'>Tumble Cup Dashboard</h1>", unsafe_allow_html=True)
image = Image.open("Tumblecup.jpeg")
left_co, cent_co, last_co = st.columns(3)
with cent_co:
    st.image(image, width=500)

tab1, tab2 = st.tabs(["Status Update", "Analytics"])
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
with tab1:
    st.header("Order Status Management")

    if not orders_df.empty:
        st.dataframe(filtered_df)

        st.subheader("Update by ID")
        col1, col2, col3 = st.columns(3)

        with col1:
            # order_id = st.number_input("Order ID", min_value=int(orders_df['ID'].min()),
            #                            max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
            #                            step=1, key="status_order_id")
            # payment_order_id = st.number_input("Payment Order ID", min_value=int(orders_df['ID'].min()),
            #                                    max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
            #                                    step=1, key="payment_order_id")
            delete_order_id = st.number_input("Delete Order ID", min_value=int(orders_df['ID'].min()),
                                              max_value=int(orders_df['ID'].max()) if not orders_df.empty else 1,
                                              step=1, key="delete_order_id")

        # with col2:
        #     new_status = st.selectbox("New Status",
        #                               ["Pending", "Processing", "Shipped", "Delivered", "Cancelled"],
        #                               key="new_status")
        #     payment_new_status = st.selectbox("Payment New Status",
        #                                       ["Pending", "Processing", "Confirmed", "Cancelled"],
        #                                       key="payment_new_status")

        with col2:
            # if st.button("Update Status", key="update_status_btn"):
            #     if update_order_status(order_id, new_status):
            #         st.success(f"Order #{order_id} status updated to {new_status}")
            #         customer_email = orders_df.loc[orders_df["ID"] == order_id, "Email"].values[0]
            #         email_content = f"Your order #{order_id} status has been updated to {new_status}."
            #         send_email_notification(customer_email, "Tumble Cup Order Status Update", email_content)
            #         st.rerun()
            #     else:
            #         st.error(f"Failed to update order #{order_id}")
            #
            # if st.button("Update Payment Status", key="update_payment_btn"):
            #     if update_payment_status(payment_order_id, payment_new_status):
            #         st.success(f"Order #{payment_order_id} payment status updated to {payment_new_status}")
            #         customer_email = orders_df.loc[orders_df["ID"] == payment_order_id, "Email"].values[0]
            #         email_content = f"Your order #{payment_order_id} payment status has been updated to {payment_new_status}."
            #         send_email_notification(customer_email, "Tumble Cup Payment Status Update", email_content)
            #         st.rerun()
            #     else:
            #         st.error(f"Failed to update payment status for order #{payment_order_id}")

            if st.button("Delete Order", key="delete_order_btn"):
                if delete_order(delete_order_id):
                    st.success(f"Order #{delete_order_id} has been deleted")
                    st.rerun()
                else:
                    st.error(f"Failed to delete order #{delete_order_id}")

        st.divider()
        st.subheader("Update by Order Number")
        order_num_col1, order_num_col2, order_num_col3, order_num_col4 = st.columns(4)

        with order_num_col1:
            order_number = st.text_input("Order Number (e.g., TC00001)",
                                         placeholder="Enter full or partial order number",
                                         key="order_number_input")

            if order_number:
                matches = filtered_df[filtered_df["Order Number"].astype(str).str.contains(order_number, case=False)]
                match_count = len(matches)

                if match_count > 0:
                    st.info(f"Found {match_count} matching orders")
                    if st.checkbox("Show matching orders", key="show_matches"):
                        st.dataframe(matches[["Order Number", "Name", "Status", "Payment Status"]])
                else:
                    st.warning("No matching orders found")

        with order_num_col2:
            update_type = st.radio("What to update",
                                   ["Order Status", "Payment Status", "Both"],
                                   key="order_num_update_type")

        with order_num_col3:
            if "Order Status" in update_type:
                order_num_status = st.selectbox("New Order Status",
                                                ["Pending", "Processing", "Shipped", "Delivered", "Cancelled"],
                                                key="order_num_order_status")
                batch_tracking_id = None
                partner = None
                if order_num_status == "Shipped":
                    batch_tracking_id = st.text_input("Shipping/Tracking ID",
                                                      placeholder="Enter tracking number for all matching orders",
                                                      key="batch_tracking_id_input")
                    partner = st.text_input("Shipping Partner",
                                            placeholder="Enter shipping partner",
                                            key="shipping_partner_input")
            if "Payment Status" in update_type:
                order_num_payment_status = st.selectbox("New Payment Status",
                                                        ["Pending", "Processing", "Confirmed", "Cancelled"],
                                                        key="order_num_payment_status")

        with order_num_col4:
            if st.button("Update All Matching Orders", key="update_by_order_num_btn"):
                if not order_number:
                    st.error("Please enter an Order Number")
                else:
                    updates_made = False
                    if update_type in ["Order Status", "Both"]:
                        success_count, updated_ids = update_by_order_number(
                            order_number,
                            order_num_status,
                            "Status",
                            batch_tracking_id if order_num_status == "Shipped" else None,
                            partner if order_num_status == "Shipped" else None
                        )

                        if success_count > 0:
                            success_msg = f"Updated order status to '{order_num_status}' for {success_count} orders"
                            if batch_tracking_id and order_num_status == "Shipped":
                                success_msg += f" with tracking ID: {batch_tracking_id} and Shipping Partner: {partner}"
                            st.success(success_msg)

                            updates_made = True
                            for order_id in updated_ids:
                                try:
                                    customer_email = orders_df.loc[orders_df["ID"] == order_id, "Email"].values[0]
                                    order_num = orders_df.loc[orders_df["ID"] == order_id, "Order Number"].values[0]
                                    email_content = f"""
                                            <p>Dear Customer,</p>
                                            <p>Your order <strong>{order_num}</strong> status has been updated to <strong>{order_num_status}</strong>.</p>
                                            """

                                    if batch_tracking_id and order_num_status == "Shipped":
                                        email_content += f"""
                                                <p>Your shipment is on its way! You can track your package using the tracking number: 
                                                <strong>{batch_tracking_id} via {partner}</strong></p>
                                                """

                                    email_content += "<p>Thank you for shopping with Tumble Cup!</p>"

                                    send_email_notification(customer_email, "Tumble Cup Order Status Update",
                                                            email_content)

                                    break
                                except Exception as e:
                                    st.warning(f"Could not send email for order #{order_id}: {str(e)}")

                    if update_type in ["Payment Status", "Both"]:
                        success_count, updated_ids = update_by_order_number(
                            order_number,
                            order_num_payment_status,
                            "Payment Status"
                        )

                        if success_count > 0:
                            st.success(
                                f"Updated payment status to '{order_num_payment_status}' for {success_count} orders")
                            updates_made = True

                            for order_id in updated_ids:
                                try:
                                    customer_email = orders_df.loc[orders_df["ID"] == order_id, "Email"].values[0]
                                    order_num = orders_df.loc[orders_df["ID"] == order_id, "Order Number"].values[0]
                                    email_content = f"""
                                            <p>Dear Customer,</p>
                                            <p>Your order <strong>{order_num}</strong> payment status has been updated to <strong>{order_num_payment_status}</strong>.</p>
                                            <p>Thank you for shopping with Tumble Cup!</p>
                                            """
                                    send_email_notification(customer_email, "Tumble Cup Payment Status Update",
                                                            email_content)
                                    break
                                except Exception as e:
                                    st.warning(f"Could not send email for order #{order_id}: {str(e)}")

                    if updates_made:
                        st.rerun()
                    else:
                        st.error(f"No orders found matching '{order_number}'")
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
with tab2:
    st.header("Sales Analytics")

    if not orders_df.empty:
        st.subheader("Sales Metrics")

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

            if 'style_breakdown' in metrics and not metrics['style_breakdown'].empty:
                st.subheader("Custom & Handpainted Items Analysis")
                metrics_df_style = metrics['style_breakdown']

                # metrics_df_style.index = range(1, len(metrics_df_style) + 1)

                st.dataframe(metrics_df_style)

                st.subheader("Custom & Handpainted Items Comparison")
                style_chart = metrics['style_breakdown'].groupby('Item Style').agg({
                    'Total Quantity': 'sum',
                    'Profit': 'sum'
                }).reset_index()

                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Quantity by Style")
                    st.bar_chart(style_chart.set_index('Item Style')['Total Quantity'])

                with col2:
                    st.subheader("Profit by Style")
                    st.bar_chart(style_chart.set_index('Item Style')['Profit'])
                if len(style_chart) > 1:
                    style_chart['Profit per Item'] = style_chart['Profit'] / style_chart['Total Quantity']
                    st.subheader("Profitability Analysis")
                    st.bar_chart(style_chart.set_index('Item Style')['Profit per Item'])
                    style_chart['Profit %'] = (style_chart['Profit'] / style_chart['Profit'].sum()) * 100
                    style_chart['Quantity %'] = (style_chart['Total Quantity'] / style_chart[
                        'Total Quantity'].sum()) * 100

                    st.subheader("Contribution Analysis")
                    contribution_df = style_chart[['Item Style', 'Profit %', 'Quantity %']]
                    st.dataframe(contribution_df)

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

            st.subheader("Cost vs Profit Ratio")
            if metrics['total_sales'] > 0:
                pie_data = {
                    'Category': ['Cost', 'Profit'],
                    'Value': [metrics['total_costs'], metrics['total_profit']]
                }
                pie_df = pd.DataFrame(pie_data)

                cost_percentage = (metrics['total_costs'] / metrics['total_sales']) * 100
                profit_percentage = (metrics['total_profit'] / metrics['total_sales']) * 100

                col1, col2 = st.columns([3, 2])

                with col1:
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
                    st.metric("Cost Percentage", f"{cost_percentage:.2f}%")
                    st.metric("Profit Percentage", f"{profit_percentage:.2f}%")
                    st.metric("Profit Margin", f"{(profit_percentage):.2f}%")
            else:
                st.info("No sales data available to calculate cost vs profit ratio.")
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

import calendar
import smtplib
from datetime import datetime
from email.message import EmailMessage
import pandas as pd
import streamlit as st
from PIL import Image
from notion_client import Client

st.set_page_config(
    page_title="Tumble Cup Admin",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Initialize Notion client
notion = Client(auth=st.secrets["Notion"]["NOTION_TOKEN"])
DATASOURCE_ID = st.secrets["Notion"]["DATASOURCE_ID"]

month_list = list(calendar.month_name)[1:]
current_month = datetime.today().month
current_month_name = calendar.month_name[current_month]
current_year = datetime.today().year

st.markdown("""
<style>
/* Add your custom CSS here */
</style>
""", unsafe_allow_html=True)

PRODUCT_COSTS = {
    "Classic Tumbler": 1850,
    "Can Glass": 1250,
    "Coffee Mug": 1500
}


def notion_date_to_datetime(notion_date):
    """Convert Notion date format to datetime"""
    if notion_date and 'start' in notion_date:
        return pd.to_datetime(notion_date['start'])
    return None


def get_property_value(properties, prop_name):
    """Extract value from Notion property based on its type"""
    if prop_name not in properties:
        return None

    prop = properties[prop_name]
    prop_type = prop['type']

    if prop_type == 'title':
        return prop['title'][0]['plain_text'] if prop['title'] else ""
    elif prop_type == 'rich_text':
        return prop['rich_text'][0]['plain_text'] if prop['rich_text'] else ""
    elif prop_type == 'number':
        return prop['number']
    elif prop_type == 'select':
        return prop['select']['name'] if prop['select'] else ""
    elif prop_type == 'status':
        return prop['status']['name'] if prop['status'] else ""
    elif prop_type == 'date':
        return notion_date_to_datetime(prop['date'])
    elif prop_type == 'email':
        return prop['email']
    elif prop_type == 'phone_number':
        return prop['phone_number']
    elif prop_type == 'checkbox':
        return prop['checkbox']
    else:
        return None


def notion_to_dataframe(results):
    """Convert Notion query results to pandas DataFrame"""
    data = []
    for page in results:
        props = page['properties']
        row = {
            'ID': page['id'],
            'Order Number': get_property_value(props, 'Order Number'),
            'Name': get_property_value(props, 'Customer Name'),
            'Email': get_property_value(props, 'Email'),
            'Phone': get_property_value(props, 'Phone'),
            'Address': get_property_value(props, 'Address'),
            'City': get_property_value(props, 'City'),
            'Item Name': get_property_value(props, 'Item'),
            'Item Quantity': get_property_value(props, 'Quantity'),
            'Item Style': get_property_value(props, 'Item Style'),
            'Base Price': get_property_value(props, 'Base Price'),
            'Price': get_property_value(props, 'Price'),
            'Total': get_property_value(props, 'Total'),
            'Order Date': get_property_value(props, 'Date'),
            'Status': get_property_value(props, 'Status'),
            'Payment Status': get_property_value(props, 'Payment Status'),
            'Payment Method': get_property_value(props, 'Payment Method'),
            'Tracking ID': get_property_value(props, 'Tracking ID'),
            'Tracking Partner': get_property_value(props, 'Tracking Partner'),
        }
        data.append(row)

    return pd.DataFrame(data)

@st.cache_data(ttl=600)
def get_orders(month_number=None):
    """Fetch orders from Notion data source"""
    try:
        # Query all pages from the data source
        all_results = []
        has_more = True
        start_cursor = None

        while has_more:
            if start_cursor:
                response = notion.data_sources.query(
                    data_source_id=DATASOURCE_ID,
                    start_cursor=start_cursor,
                    sorts=[{"timestamp": "created_time", "direction": "ascending"}]
                )
            else:
                response = notion.data_sources.query(
                    data_source_id=DATASOURCE_ID,
                    sorts=[{"timestamp": "created_time", "direction": "ascending"}]
                )

            all_results.extend(response['results'])
            has_more = response['has_more']
            start_cursor = response.get('next_cursor')

        orders_df = notion_to_dataframe(all_results)

        if orders_df.empty:
            return pd.DataFrame()

        # Filter by month if specified
        if month_number:
            orders_df = orders_df[orders_df["Order Date"].dt.month == month_number]

        return orders_df
    except Exception as e:
        st.error(f"Error fetching orders: {e}")
        return pd.DataFrame()


def update_notion_property(page_id, property_name, value, property_type='select'):
    """Update a specific property in Notion"""
    try:
        properties = {}

        if property_type == 'select':
            properties[property_name] = {"select": {"name": value}}
        elif property_type == 'status':
            properties[property_name] = {"status": {"name": value}}
        elif property_type == 'rich_text':
            properties[property_name] = {"rich_text": [{"text": {"content": value}}]}
        elif property_type == 'number':
            properties[property_name] = {"number": value}

        notion.pages.update(page_id=page_id, properties=properties)
        return True
    except Exception as e:
        st.error(f"Error updating property: {e}")
        return False


def update_order_status(order_id, new_status):
    """Update order status in Notion"""
    return update_notion_property(order_id, 'Status', new_status, 'select')


def update_payment_status(order_id, new_status):
    """Update payment status in Notion"""
    return update_notion_property(order_id, 'Payment Status', new_status, 'status')


def update_by_order_number(order_number, new_status, status_field="Status", tracking_id=None, partner=None):
    """
    Update all orders with matching order number
    """
    orders_df = get_orders()
    matching_orders = orders_df[orders_df["Order Number"].astype(str).str.contains(order_number, case=False, na=False)]

    if matching_orders.empty:
        return 0, []

    matching_ids = matching_orders["ID"].tolist()
    success_count = 0

    for page_id in matching_ids:
        property_name = 'Status' if status_field == 'Status' else 'Payment Status'
        if update_notion_property(page_id, property_name, new_status, 'select'):
            success_count += 1

            # Update tracking info if provided
            if tracking_id and status_field == "Status" and new_status == "Shipped" and partner:
                update_notion_property(page_id, 'Tracking ID', tracking_id, 'rich_text')
                update_notion_property(page_id, 'Tracking Partner', partner, 'rich_text')

    return success_count, matching_ids


def delete_order(order_id):
    """Archive (delete) an order in Notion"""
    try:
        notion.pages.update(page_id=order_id, archived=True)
        return True
    except Exception as e:
        st.error(f"Error deleting order: {e}")
        return False


def send_email_notification(to_email, subject, content):
    """Send email notification"""
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

        custom_styles = style_df[
            style_df['Item Style'].str.contains('Custom|Hand painted|Handpainted', case=False, na=False)]
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


def get_date_orders(date, month, year):
    """Get orders for a specific date"""
    orders_df = get_orders()
    if orders_df.empty:
        return pd.DataFrame()

    orders_df["Order Date"] = pd.to_datetime(orders_df["Order Date"], format="ISO8601", errors='coerce')

    orders_df = orders_df[orders_df["Order Date"].notna()]

    if orders_df.empty:
        return pd.DataFrame()

    orders_df = orders_df[
        (orders_df["Order Date"].dt.day == date) &
        (orders_df["Order Date"].dt.month == month) &
        (orders_df["Order Date"].dt.year == year)
        ]
    return orders_df



st.markdown("<h1 style='text-align: center;'>Tumble Cup Dashboard</h1>", unsafe_allow_html=True)

image = Image.open("Tumblecup.jpeg")
left_co, cent_co, last_co = st.columns(3)
with cent_co:
    st.image(image, width=500)

tab1, tab2, tab3, tab4 = st.tabs(["Admin", "Status Update", "Filter", "Analytics"])


with st.sidebar:
    st.header("Filter Options")
    selected_month = st.selectbox(
        "Select Month",
        month_list,
        index=current_month - 1,
        placeholder="Select Month"
    )
    if st.button("🔃 Fetch Latest"):
        st.cache_data.clear()

    selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
    orders_df = get_orders(selected_month_number)

    if not orders_df.empty:
        orders_df["Order Date"] = orders_df["Order Date"].dt.strftime("%d-%B-%Y")

        search_term = st.text_input("Search by Name or Order Number",
                                    placeholder="Enter Search Term",
                                    key="search_term")

        if search_term:
            orders_df = orders_df[
                orders_df['Name'].str.contains(search_term, case=False, na=False) |
                orders_df['Order Number'].str.contains(search_term, case=False, na=False)
                ]
            if orders_df.empty:
                st.warning("No such orders found!")

        status_filter = st.multiselect("Filter by Status",
                                       options=orders_df['Status'].unique().tolist(),
                                       default=orders_df['Status'].unique().tolist())

        payment_filter = st.multiselect("Filter by Payment Status",
                                        options=orders_df['Payment Status'].unique().tolist(),
                                        default=orders_df['Payment Status'].unique().tolist())

        filtered_df = orders_df[
            orders_df['Status'].isin(status_filter) &
            orders_df['Payment Status'].isin(payment_filter)
            ]
    else:
        st.warning("No orders found!")
        filtered_df = pd.DataFrame()

# Session state for password
if 'user_entered' not in st.session_state:
    st.session_state.user_entered = {}

# Tab 1: Admin Login
with tab1:
    st.header("Admin Login 🔑")
    pwd = st.text_input(label="Enter password", type="password",
                        placeholder="Enter Password", key="password")

    if st.button("Login"):
        if str(pwd).strip():
            st.session_state.user_entered["password"] = pwd

            if st.session_state.user_entered.get("password") == st.secrets["Password"]["Password"]:
                st.success("Access granted! ✅")
            else:
                st.warning("Access denied! ⛔")
        else:
            st.info("Enter Password!")

# Tab 2: Status Update
with tab2:
    if st.session_state.user_entered.get("password") == st.secrets["Password"]["Password"]:
        st.header("Order Status Management")

        if not orders_df.empty:
            st.dataframe(filtered_df)

            st.subheader("Update by ID")
            col1, col2 = st.columns(2)

            with col1:
                delete_order_id = st.text_input("Delete Order ID (Notion Page ID)",
                                                placeholder="Enter Notion page ID",
                                                key="delete_order_id")

            with col2:
                if st.button("Delete Order", key="delete_order_btn"):
                    if delete_order_id:
                        if delete_order(delete_order_id):
                            st.success(f"Order #{delete_order_id} has been deleted")
                            st.rerun()
                        else:
                            st.error(f"Failed to delete order #{delete_order_id}")
                    else:
                        st.warning("Please enter an Order ID")

            st.divider()

            st.subheader("Update by Order Number")
            order_num_col1, order_num_col2, order_num_col3, order_num_col4 = st.columns(4)

            with order_num_col1:
                order_number = st.text_input("Order Number (e.g., TC00001)",
                                             placeholder="Enter full or partial order number",
                                             key="order_number_input")

                if order_number:
                    matches = filtered_df[
                        filtered_df["Order Number"].astype(str).str.contains(order_number, case=False, na=False)
                    ]
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
                                                          placeholder="Enter tracking number",
                                                          key="batch_tracking_id_input")
                        partner = st.text_input("Shipping Partner",
                                                placeholder="Enter shipping partner",
                                                key="shipping_partner_input")

                if "Payment Status" in update_type:
                    order_num_payment_status = st.selectbox("New Payment Status",
                                                            ["Pending", "Processing", "Confirmed", "Cancelled"],
                                                            key="order_num_payment_status")

            with order_num_col4:
                show_button = True

                if not order_number:
                    show_button = False
                    st.warning("Please enter an Order Number")

                if update_type in ["Order Status", "Both"] and order_num_status == "Shipped":
                    if not batch_tracking_id:
                        show_button = False
                        st.warning("Tracking ID is required when status is 'Shipped'")
                    if not partner:
                        show_button = False
                        st.warning("Shipping Partner is required when status is 'Shipped'")

                if show_button:
                    if st.button("Update All Matching Orders", key="update_by_order_num_btn"):
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

                                # Send email notifications
                                for order_id in updated_ids:
                                    try:
                                        order_row = orders_df[orders_df["ID"] == order_id]
                                        if not order_row.empty:
                                            customer_email = order_row["Email"].values[0]
                                            order_num = order_row["Order Number"].values[0]

                                            if order_num_status != "Shipped":
                                                email_content = f"""
                                                <html>
                                                <body>
                                                <p>Dear Customer,</p>
                                                <p>Your order {order_num} status has been updated to {order_num_status}.</p>
                                                </body>
                                                </html>
                                                """
                                            else:
                                                email_content = f"""
                                                <html>
                                                <body>
                                                <p>Dear Customer,</p>
                                                <p>Your order {order_num} status has been updated to {order_num_status}.</p>
                                                <p>Your shipment is on its way! You can track your package using the tracking number: {batch_tracking_id} via {partner}</p>
                                                </body>
                                                </html>
                                                """

                                            email_content += "<p>Thank you for shopping with Tumble Cup!</p></body></html>"
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

                                # Send email notifications
                                for order_id in updated_ids:
                                    try:
                                        order_row = orders_df[orders_df["ID"] == order_id]
                                        if not order_row.empty:
                                            customer_email = order_row["Email"].values[0]
                                            order_num = order_row["Order Number"].values[0]

                                            email_content = f"""
                                            <html>
                                            <body>
                                            <p>Dear Customer,</p>
                                            <p>Your order {order_num} payment status has been updated to {order_num_payment_status}.</p>
                                            <p>Thank you for shopping with Tumble Cup!</p>
                                            </body>
                                            </html>
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
    else:
        st.warning("Access denied! ⛔")

# Tab 3: Filter
with tab3:
    if st.session_state.user_entered.get("password") == st.secrets["Password"]["Password"]:
        st.header("Order Filtering 🥅")
        date = st.date_input("Select a date", value="today")

        if date:
            dd = date.day
            month = date.month
            year = date.year
            df = get_date_orders(dd, month, year)

            if df.empty:
                st.error(f"No orders found for {dd}-{month}-{year}")
            else:
                st.dataframe(df)
    else:
        st.warning("Access denied! ⛔")


with tab4:
    if st.session_state.user_entered.get("password") == st.secrets["Password"]["Password"]:
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
                metrics_df_style.index = range(1, len(metrics_df_style) + 1)
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

            if st.button("Export Analytics to CSV", key="export_csv_analytics"):
                if not metrics['product_breakdown'].empty:
                    analytics_csv = metrics['product_breakdown'].to_csv(index=False)
                    st.download_button(
                        label="Download Analytics CSV",
                        data=analytics_csv,
                        file_name=f"tumble_cup_analytics_{datetime.today().strftime('%Y-%m-%d')}.csv",
                        mime="text/csv"
                    )
        else:
            st.info("No orders found for the selected month.")
    else:
        st.warning("Access denied! ⛔")
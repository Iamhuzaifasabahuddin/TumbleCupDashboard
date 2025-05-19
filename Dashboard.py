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

def get_orders(month_number=None):
    orders_df = conn.read(worksheet="Tumble_cup")

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


                    # display_df = filtered_df.copy()
                    # display_df["Order Date"] = display_df["display_date"]
                    # display_df = display_df.drop(columns=["display_date"])
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

                # Order management section
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
                #
                # # Backup and analytics section
                # st.subheader("Data Management Tools")
                #
                # if st.button("Backup to Google Drive"):
                #     try:
                #         # Create a timestamp for the backup
                #         timestamp = datetime.today().strftime('%Y%m%d_%H%M%S')
                #         backup_sheet_name = f"Orders_Backup_{timestamp}"
                #
                #         # Create a new worksheet for the backup
                #         conn.update(worksheet=backup_sheet_name, data=orders_df)
                #
                #         st.success(f"Data backed up successfully to sheet '{backup_sheet_name}'")
                #     except Exception as e:
                #         st.error(f"Backup failed: {str(e)}")
                #
                # # Add monthly sales analytics
                # st.subheader("Monthly Sales Analytics")
                #
                # # Group by day and count orders
                # if not filtered_df.empty:
                #     filtered_df["day"] = filtered_df["Order Date"].dt.day
                #     daily_orders = filtered_df.groupby("day").size().reset_index(name="orders")
                #
                #     # Create chart
                #     st.bar_chart(daily_orders.set_index("day"))
                #
                #     # Total revenue
                #     if "order_total" in filtered_df.columns:
                #         total_revenue = filtered_df["order_total"].sum()
                #         st.metric("Total Revenue", f"${total_revenue:.2f}")
        else:
            st.info("No orders found for the selected month.")
    elif admin_password:
        st.error("Incorrect password")
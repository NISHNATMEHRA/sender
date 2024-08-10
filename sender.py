import pandas as pd
import win32com.client as win32
import pythoncom
import streamlit as st

# Function to send an email
def send_email(to_email, subject, body, cc_mail=None):
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        # Convert CC email addresses to strings and handle NaNs
        if cc_mail is None:
            cc_mail = []
        cc_emails = [str(email).strip() for email in cc_mail if pd.notna(email)]
        cc_emails = [email for email in cc_emails if email]
        cc_str = ";".join(cc_emails)

        # Initialize Outlook
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # Configure email
        mail.To = to_email
        mail.CC = cc_str
        mail.Subject = subject
        mail.HTMLBody = body

        try:
            # Send the email
            mail.Send()
            return "Email sent successfully."
        except Exception as e:
            return f"Failed to send email. Error: {e}"
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

# Streamlit app
def main():
    st.title("Email Sender")

    city = st.text_input("Enter the City:")
    order_id = st.text_input("Enter the Order ID:")
    cr_name = st.text_input("Enter the CR Name:")
    concern = st.text_input("Enter the Concern:")
    status = st.text_input("Enter the Status:")
    status_time = st.text_input("Enter the Status Time:")
    remarks = st.text_area("Enter the Remarks:")

    to_email = st.text_input("Recipient Email Address:")
    cc_mail = st.text_area("CC Email Addresses (comma-separated):")
    cc_mail = [email.strip() for email in cc_mail.split(',') if email.strip()]

    if st.button("Send Email"):
        if not to_email:
            st.error("Recipient email address is required.")
        else:
            subject = f"! High Important || Warning Mail For {status} || {cr_name}"
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>{subject}</title>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        line-height: 1.6;
                        margin: 20px;
                    }}
                    .container {{
                        width: 80%;
                        margin: auto;
                        padding: 20px;
                        border: 1px solid #ddd;
                        border-radius: 5px;
                        background-color: #f9f9f9;
                    }}
                    h1 {{
                        color: #e74c3c;
                    }}
                    table {{
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                    }}
                    table, th, td {{
                        border: 1px solid #ddd;
                    }}
                    th, td {{
                        padding: 8px;
                        text-align: left;
                    }}
                    th {{
                        background-color: #f2f2f2;
                    }}
                    .remarks {{
                        color: #e74c3c;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>{subject}</h1>
                    <table>
                        <tr>
                            <th>City</th>
                            <td>{city}</td>
                        </tr>
                        <tr>
                            <th>Order ID</th>
                            <td>{order_id}</td>
                        </tr>
                        <tr>
                            <th>CR Name</th>
                            <td>{cr_name}</td>
                        </tr>
                        <tr>
                            <th>Concern</th>
                            <td>{concern}</td>
                        </tr>
                        <tr>
                            <th>Status</th>
                            <td>{status}</td>
                        </tr>
                        <tr>
                            <th>Status Time</th>
                            <td>{status_time}</td>
                        </tr>
                    </table>
                    <p class="remarks">{remarks}</p>
                </div>
            </body>
            </html>
            """
            result = send_email(to_email, subject, html_content, cc_mail)
            st.write(result)

if __name__ == "__main__":
    main()

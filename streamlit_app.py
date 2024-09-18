import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import base64
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.generator import BytesGenerator


def split_excel_and_zip(df, column_name):
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for value, group in df.groupby(column_name):
            excel_buffer = BytesIO()

            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                group.to_excel(writer, index=False, sheet_name='Sheet1')

            excel_buffer.seek(0)
            zipf.writestr(f"{value}.xlsx", excel_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


def create_outlook_emails(df, column_name, email_list):
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for value, group in df.groupby(column_name):
            # Create Excel file in memory
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                group.to_excel(writer, index=False, sheet_name='Sheet1')
            excel_buffer.seek(0)

            # Create email message
            msg = MIMEMultipart()
            msg['Subject'] = f'Split Excel File - {value}'
            msg['To'] = '; '.join(email_list)
            msg['From'] = 'your_email@example.com'  # Add a From address
            msg.attach(MIMEText(f'Please find attached the Excel file for {value}.'))

            # Attach Excel file
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            part.set_payload(excel_buffer.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{value}.xlsx"')
            msg.attach(part)

            # Save as .eml file
            eml_file = BytesIO()
            generator = BytesGenerator(eml_file)
            generator.flatten(msg)
            eml_file.seek(0)

            zipf.writestr(f"{value}.eml", eml_file.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


st.title('Excel File Splitter')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("File uploaded successfully. Preview:")
    st.write(df.head())

    column_name = st.selectbox("Select the column to split by:", df.columns)

    output_format = st.radio("Choose output format:", ("ZIP", "Email"))

    if output_format == "Email":
        email_list = st.text_area("Enter email addresses (one per line):")
        email_list = [email.strip() for email in email_list.split('\n') if email.strip()]

    if st.button("Process"):
        if output_format == "ZIP":
            zip_buffer = split_excel_and_zip(df, column_name)

            st.download_button(
                label="Download ZIP file",
                data=zip_buffer,
                file_name="suppliers_files.zip",
                mime="application/zip"
            )
        else:  # Email format
            if email_list:
                zip_buffer = create_outlook_emails(df, column_name, email_list)

                st.download_button(
                    label="Download Email Files (ZIP)",
                    data=zip_buffer,
                    file_name="email_files.zip",
                    mime="application/zip"
                )
            else:
                st.error("Please enter at least one email address.")
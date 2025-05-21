import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from PyPDF2 import PdfReader, PdfWriter

# App settings
st.set_page_config(page_title="Payslip Generator", layout="centered")
st.title("📩 ENA Coach Payslip Generator & Email Sender")

# Upload
uploaded_file = st.file_uploader("Upload Payroll Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        if 'Month' not in df.columns or 'Email' not in df.columns:
            st.error("Missing required 'Month' or 'Email' columns in your file.")
        else:
            months = sorted(df['Month'].dropna().unique())
            selected_month = st.selectbox("Select Month to Send Payslips", months)
            filtered_df = df[df['Month'] == selected_month]

            st.write(f"✅ Found {len(filtered_df)} employees for month: **{selected_month}**")
            st.dataframe(filtered_df)

            # Email setup
            st.subheader("📧 Email Settings")
            sender_email = st.text_input("Sender Email", value="your_email@example.com")
            sender_password = st.text_input("App Password (Gmail App Password)", type="password")

            if st.button("📨 Send Payslips"):
                if not sender_email or not sender_password:
                    st.error("Please provide both sender email and password.")
                else:
                    for _, row in filtered_df.iterrows():
                        try:
                            # Generate unencrypted PDF
                            buffer = BytesIO()
                            c = canvas.Canvas(buffer, pagesize=A4)
                            width, height = A4

                            # Left-aligned Company Info
                            c.setFont("Helvetica-Bold", 16)
                            c.drawString(30, height - 40, "ENA COACH LTD")
                            c.setFont("Helvetica", 10)
                            c.drawString(30, height - 60, "KPCU, Nairobi Kenya")
                            c.drawString(30, height - 75, "Phone: +254 709 832 000")
                            c.drawString(30, height - 90, "Email: info@enacoach.co.ke")

                            # Right-aligned Pay Info
                            c.setFont("Helvetica-Bold", 10)
                            c.drawRightString(width - 30, height - 40, f"PAY DATE: 20 {selected_month} 2025")
                            c.drawRightString(width - 30, height - 60, "PAY TYPE: Bank Transfer")
                            c.drawRightString(width - 30, height - 80, f"PERIOD: {selected_month}")

                            # Payslip Title and Employee Info
                            c.setFont("Helvetica-Bold", 14)
                            c.drawString(30, height - 120, "EMPLOYEE PAYSLIP")
                            c.setFont("Helvetica", 11)
                            c.drawString(30, height - 140, f"Month: {selected_month}")
                            c.drawString(30, height - 160, f"Employee Name: {row['Name']}")
                            c.drawString(30, height - 180, f"Email: {row['Email']}")

                            # Salary Table
                            c.setFont("Helvetica-Bold", 12)
                            c.drawString(30, height - 210, "SALARY BREAKDOWN")

                            table_data = [
                                ['Description', 'Amount (KES)'],
                                ['Basic Salary', f"{row['Basic Salary']:.2f}"],
                                ['Overtime Pay', f"{row['Overtime']:.2f}"],
                                ['Allowance', f"{row['Allowance']:.2f}"],
                                ['PAYE Tax', f"{row['PAYE Tax']:.2f}"],
                                ['SHA', f"{row['SHA']:.2f}"],
                                ['NSSF', f"{row['NSSF']:.2f}"],
                                ['Penalties', f"{row['Penalties']:.2f}"],
                                ['Deductions', f"{row['Deductions']:.2f}"],
                                ['Net Pay', f"{row['Net Salary']:.2f}"],
                            ]

                            t = Table(table_data, colWidths=[140, 100])
                            t.setStyle(TableStyle([
                                ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                ('FONTSIZE', (0, 0), (-1, -1), 10),
                                ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.black),
                                ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
                            ]))

                            table_x = 30
                            table_y = height - 460
                            t.wrapOn(c, width, height)
                            t.drawOn(c, table_x, table_y)

                            # Footer
                            c.setFont("Helvetica", 10)
                            c.drawString(30, table_y - 40, "Payment Method: Bank Transfer")
                            c.drawString(30, table_y - 60, f"Sent on: 20 {selected_month} 2025")

                            c.showPage()
                            c.save()
                            buffer.seek(0)

                            # Encrypt PDF
                            reader = PdfReader(buffer)
                            writer = PdfWriter()
                            for page in reader.pages:
                                writer.add_page(page)

                            pin = str(row['pin']) if 'pin' in row and pd.notna(row['pin']) else "1234"
                            writer.encrypt(user_password=pin, owner_password=pin)

                            encrypted_pdf = BytesIO()
                            writer.write(encrypted_pdf)
                            encrypted_pdf.seek(0)

                            # Email
                            msg = MIMEMultipart()
                            msg['From'] = sender_email
                            msg['To'] = row['Email']
                            msg['Subject'] = f"{selected_month} Payslip - ENA COACH LTD"

                            body = f"Dear {row['Name']},\n\nPlease find attached your payslip for {selected_month}.\n\nRegards,\nENA COACH LTD"
                            msg.attach(MIMEText(body, 'plain'))

                            attachment = MIMEApplication(encrypted_pdf.read(), _subtype="pdf")
                            attachment.add_header('Content-Disposition', 'attachment', filename=f"Payslip_{row['Name'].replace(' ', '_')}.pdf")
                            msg.attach(attachment)

                            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                                server.login(sender_email, sender_password)
                                server.send_message(msg)

                            st.success(f"Payslip sent to {row['Email']}")

                        except Exception as e:
                            st.error(f"Error sending to {row['Email']}: {e}")

    except Exception as e:
        st.error(f"Error processing file: {e}")

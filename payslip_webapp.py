import streamlit as st
import pandas as pd
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
import yagmail
import os
import tempfile
import time
from datetime import datetime
# Test secret access - REMOVE THIS AFTER TESTING
email = st.secrets["PAYSLIP_EMAIL"]  # ← Use the KEY, not the value
password = st.secrets["PAYSLIP_APP_PASSWORD"]

# Email credentials (use Gmail app password)
EMAIL = "bradenkipkirui@gmail.com"
APP_PASSWORD = "vxgi ozkz czer gcdp"

# Configure Streamlit page
st.set_page_config(page_title="Encrypted Payslip Generator", layout="centered")
st.title("🔐 Encrypted Payslip Generator & Sender")

# Helper function for safe file cleanup
def safe_file_cleanup(*files):
    """
    Safely remove files with retries and error handling
    """
    for filepath in files:
        try:
            if filepath and os.path.exists(filepath):
                os.remove(filepath)
        except PermissionError:
            time.sleep(1)  # Wait a second and try again
            try:
                os.remove(filepath)
            except:
                pass  # Give up if still locked
        except:
            pass  # Ignore other errors

# File uploader
uploaded_file = st.file_uploader(r"C:\Users\Ena Data\OneDrive\Desktop\b\EncryptedPayslipApp\Payslip.csv", type=['csv', 'xlsx'])

if uploaded_file:
    # Read the uploaded file
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith("csv") else pd.read_excel(uploaded_file)
        st.write("Preview of uploaded data:")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    if st.button("📤 Generate and Send Payslips"):
        with st.spinner("Processing and sending payslips..."):
            # Initialize email client
            try:
                yag = yagmail.SMTP(EMAIL, APP_PASSWORD)
            except Exception as e:
                st.error(f"Failed to login to email server: {e}")
                st.stop()

            # Process each employee
            success_count = 0
            error_count = 0
            temp_files = []  # Track all temp files for cleanup

            for _, row in df.iterrows():
                temp_pdf_path = None
                final_pdf_path = None
                
                try:
                    # Get values with fallbacks for column name variations
                    name = row.get("Name") or row.get("Employee Name") or "Employee"
                    email = row.get("Email") or ""
                    basic = float(str(row.get("Basic Salary", 0)).replace(",", ""))
                    overtime = float(str(row.get("Overtime Pay", row.get("Overtime", 0))).replace(",", ""))
                    bonus = float(str(row.get("Bonus", 0)).replace(",", ""))
                    paye = float(str(row.get("PAYE Tax", 0)).replace(",", ""))
                    sha = float(str(row.get("SHA", 0)).replace(",", ""))
                    nhif = float(str(row.get("NHIF", 0)).replace(",", ""))
                    net = float(str(row.get("Net Pay", row.get("Net Salary", 0))).replace(",", ""))
                    month = row.get("Month", datetime.now().strftime("%B %Y"))
                    pin = str(row.get("pin", "0000"))  # Default PIN if not provided

                    # Generate PDF
                    pdf = FPDF()
                    pdf.add_page()
                    
                    # Company header
                    pdf.set_font("Arial", 'B', 14)
                    pdf.cell(200, 10, txt="ENA COACH LTD", ln=1, align="C")
                    pdf.set_font("Arial", size=10)
                    pdf.cell(200, 6, txt="KPCU, Nairobi Kenya", ln=1, align="C")
                    pdf.cell(200, 6, txt="Phone: +254 709 832 000", ln=1, align="C")
                    pdf.cell(200, 6, txt="Email: info@enacoach.co.ke", ln=1, align="C")
                    pdf.ln(10)
                    
                    # Pay details
                    pdf.set_font("Arial", size=10)
                    pdf.cell(60, 6, txt="PAY DATE", border=1)
                    pdf.cell(60, 6, txt="PAY TYPE", border=1)
                    pdf.cell(60, 6, txt="PERIOD", border=1, ln=1)
                    pdf.cell(60, 8, txt=datetime.now().strftime("%d %B %Y"), border=1)
                    pdf.cell(60, 8, txt=month.split()[0], border=1)  # Extract month name
                    pdf.cell(60, 8, txt="", border=1, ln=1)
                    pdf.ln(10)
                    
                    # Employee details
                    pdf.set_font("Arial", 'B', 14)
                    pdf.cell(200, 10, txt="EMPLOYEE PAYSLIP", ln=1, align="C")
                    pdf.ln(5)
                    
                    pdf.set_font("Arial", size=10)
                    pdf.cell(60, 8, txt="Month:", border=0)
                    pdf.cell(60, 8, txt=month, border=0, ln=1)
                    pdf.cell(60, 8, txt="Employee Name:", border=0)
                    pdf.cell(60, 8, txt=name, border=0, ln=1)
                    pdf.cell(60, 8, txt="Email:", border=0)
                    pdf.cell(60, 8, txt=email, border=0, ln=1)
                    pdf.ln(10)
                    
                    # Salary breakdown
                    pdf.set_font("Arial", 'B', 12)
                    pdf.cell(200, 10, txt="SALARY BREAKDOWN", ln=1, align="L")
                    pdf.ln(5)
                    
                    col_width = 90
                    row_height = 8
                    
                    pdf.set_font("Arial", size=10)
                    pdf.cell(col_width, row_height, txt="Description", border=1)
                    pdf.cell(col_width, row_height, txt="Amount (KES)", border=1, ln=1)
                    
                    pdf.cell(col_width, row_height, txt="Basic Salary", border=1)
                    pdf.cell(col_width, row_height, txt=f"{basic:,.2f}", border=1, ln=1)
                    
                    pdf.cell(col_width, row_height, txt="Overtime Pay", border=1)
                    pdf.cell(col_width, row_height, txt=f"{overtime:,.2f}", border=1, ln=1)
                    
                    pdf.cell(col_width, row_height, txt="Bonus", border=1)
                    pdf.cell(col_width, row_height, txt=f"{bonus:,.2f}", border=1, ln=1)
                    
                    pdf.cell(col_width, row_height, txt="PAYE Tax", border=1)
                    pdf.cell(col_width, row_height, txt=f"{paye:,.2f}", border=1, ln=1)
                    
                    pdf.cell(col_width, row_height, txt="SHA", border=1)
                    pdf.cell(col_width, row_height, txt=f"{sha:,.2f}", border=1, ln=1)
                    
                    pdf.cell(col_width, row_height, txt="NHIF", border=1)
                    pdf.cell(col_width, row_height, txt=f"{nhif:,.2f}", border=1, ln=1)
                    
                    pdf.set_font("Arial", 'B', 10)
                    pdf.cell(col_width, row_height, txt="Net Pay", border=1)
                    pdf.cell(col_width, row_height, txt=f"{net:,.2f}", border=1, ln=1)
                    
                    pdf.ln(10)
                    pdf.set_font("Arial", size=10)
                    pdf.cell(200, 8, txt="Payment Method: Bank Transfer", ln=1)
                    pdf.cell(200, 8, txt=f"Sent on: {datetime.now().strftime('%d %B %Y')}", ln=1)

                    # Save to temp file
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                        pdf.output(temp_pdf.name)
                        temp_pdf_path = temp_pdf.name
                        temp_files.append(temp_pdf_path)

                    # Encrypt PDF
                    with open(temp_pdf_path, "rb") as f:
                        reader = PdfReader(f)
                        writer = PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)
                        writer.encrypt(pin)

                        final_pdf_path = os.path.join(tempfile.gettempdir(), f"Payslip_{name.replace(' ', '_')}_{month.replace(' ', '_')}.pdf")
                        with open(final_pdf_path, "wb") as f_out:
                            writer.write(f_out)
                        temp_files.append(final_pdf_path)

                    # Send email
                    yag.send(
                        to=email,
                        subject=f"Encrypted Payslip for {name} - {month}",
                        contents=f"""
Hi {name},

Attached is your encrypted payslip for {month}.

🔐 PIN to open: {pin}

Please keep this code secure.

Regards,
HR Department
ENA COACH LTD
                        """,
                        attachments=final_pdf_path
                    )

                    success_count += 1
                    time.sleep(1)  # Small delay between emails

                except Exception as e:
                    error_count += 1
                    st.error(f"Failed processing {name}: {str(e)}")
                    safe_file_cleanup(temp_pdf_path, final_pdf_path)
                    continue

            # Final cleanup
            for filepath in temp_files:
                safe_file_cleanup(filepath)

            # Show results
            if error_count == 0:
                st.success(f"✅ Successfully processed all {success_count} payslips!")
            else:
                st.warning(f"⚠️ Processed {success_count} payslips successfully, {error_count} failed")
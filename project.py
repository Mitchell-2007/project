import pandas as pd
from fpdf import FPDF
import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv

# --- Load environment variables ---
load_dotenv()

SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')

# --- Load and clean data ---
employee_data = pd.read_excel('employees.xlsx', engine='openpyxl')
employee_data.columns = employee_data.columns.str.strip()

# Fix common header typos
employee_data.rename(columns={
    'Allawances': 'Allowance',
    'Deduction': 'Deductions'
}, inplace=True)

# Ensure required columns exist
required_columns = ['Employee ID', 'Name', 'Basic Salary', 'Allowance', 'Deductions', 'Email']
for col in required_columns:
    if col not in employee_data.columns:
        employee_data[col] = 0 if col != 'Email' else ''

# Convert salary columns to numeric
employee_data['Basic Salary'] = pd.to_numeric(employee_data['Basic Salary'], errors='coerce').fillna(0)
employee_data['Allowance'] = pd.to_numeric(employee_data['Allowance'], errors='coerce').fillna(0)
employee_data['Deductions'] = pd.to_numeric(employee_data['Deductions'], errors='coerce').fillna(0)

# --- PDF Generator ---
class PayslipPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 18)
        self.set_text_color(0, 51, 102)
        self.cell(200, 10, "Mitchell Mukwaruwa", ln=True, align='C')
        self.set_font("Arial", "I", 12)
        self.cell(200, 10, "Payslip for the Month", ln=True, align='C')
        self.ln(5)
        self.set_text_color(255, 0, 0)
        self.cell(200, 2, "=" * 100, ln=True)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 10)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Page {self.page_no()}", align='C')

# --- Email Sender ---
def send_email(to_address, subject, body, attachment_path):
    msg = EmailMessage()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)

    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

# --- Create output directory ---
output_dir = "payslips"
os.makedirs(output_dir, exist_ok=True)

# --- Process Each Employee ---
for index, row in employee_data.iterrows():
    try:
        print(f"Processing: {row['Name']} ({row['Email']})")

        pdf = PayslipPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Employee Info
        pdf.set_font("Arial", "B", 12)
        pdf.cell(40, 10, "Employee Details:", ln=True)
        pdf.set_font("Arial", "", 12)
        pdf.cell(200, 8, f"Employee ID   : {row['Employee ID']}", ln=True)
        pdf.cell(200, 8, f"Name          : {row['Name']}", ln=True)
        pdf.cell(200, 8, "-" * 50, ln=True)

        # Salary Info
        pdf.set_font("Arial", "B", 12)
        pdf.cell(40, 10, "Salary Details:", ln=True)
        pdf.set_font("Arial", "", 12)
        pdf.cell(200, 8, f"Basic Salary  : $ {row['Basic Salary']:.2f}", ln=True)
        pdf.cell(200, 8, f"Allowance     : $ {row['Allowance']:.2f}", ln=True)
        pdf.cell(200, 8, f"Deductions    : $ {row['Deductions']:.2f}", ln=True)
        pdf.set_text_color(255, 0, 0)
        pdf.cell(200, 2, "=" * 80, ln=True)

        # Net Salary
        net_salary = row['Basic Salary'] + row['Allowance'] - row['Deductions']
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(200, 10, f"Net Salary    : $ {net_salary:.2f}", ln=True, align='R')

        # Save PDF
        filename = f"{row['Name'].replace(' ', '_')}_Payslip.pdf"
        filepath = os.path.join(output_dir, filename)
        pdf.output(filepath)

        # Send email if valid
        if pd.notna(row['Email']) and row['Email'].strip():
            email_body = f"""
Dear {row['Name']},

Please find attached your payslip for this month.

Regards,
Mitchell Mukwaruwa
"""
            send_email(
                to_address=row['Email'],
                subject="Your Monthly Payslip",
                body=email_body,
                attachment_path=filepath
            )

    except Exception as e:
        print(f"❌ Error processing {row['Name']} (ID: {row['Employee ID']}): {e}")

print("✅ All payslips processed and emailed successfully.")

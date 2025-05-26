from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from datetime import datetime

def format_currency(amount):
    """Format amount as currency with thousands separator"""
    try:
        return f"{float(amount):,.2f}"
    except (ValueError, TypeError):
        return "0.00"

def generate_and_send_payslip(row, sender_email, sender_password, selected_month):
    # Create PDF buffer
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # Define colors
    header_color = colors.Color(0.2, 0.3, 0.7)  # Professional blue
    text_color = colors.Color(0.2, 0.2, 0.2)    # Dark gray
    
    # Header section
    c.setFillColor(header_color)
    c.rect(0, height - 120, width, 120, fill=1)
    c.setFillColor(colors.white)
    
    # Company logo placeholder (you can add actual logo later)
    c.rect(30, height - 90, 60, 60, fill=0)
    
    # Company details
    c.setFont("Helvetica-Bold", 20)
    c.drawString(100, height - 45, "ENA COACH LTD")
    c.setFont("Helvetica", 10)
    c.drawString(100, height - 60, "KPCU, Nairobi Kenya")
    c.drawString(100, height - 75, "Phone: +254 709 832 000")
    c.drawString(100, height - 90, "Email: info@enacoach.co.ke")
    
    # Payslip details on right
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(width - 30, height - 45, f"PAY DATE: 20 {selected_month}")
    c.drawRightString(width - 30, height - 60, "PAY TYPE: Bank Transfer")
    c.drawRightString(width - 30, height - 75, f"PERIOD: {selected_month}")
    
    # Reset color for main content
    c.setFillColor(text_color)
    
    # Employee details section
    c.roundRect(30, height - 200, width - 60, 60, 5, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, height - 170, "EMPLOYEE DETAILS")
    c.setFont("Helvetica", 10)
    c.drawString(40, height - 185, f"Name: {row['Name']}")
    c.drawString(40, height - 200, f"Email: {row['Email']}")
    
    if 'Employee ID' in row:
        c.drawString(width/2, height - 185, f"Employee ID: {row['Employee ID']}")
    if 'Department' in row:
        c.drawString(width/2, height - 200, f"Department: {row['Department']}")
    
    # Salary breakdown table
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30, height - 240, "SALARY BREAKDOWN")
    
    # Create two columns for earnings and deductions
    earnings_data = [
        ['EARNINGS', 'AMOUNT (KES)'],
        ['Basic Salary', format_currency(row['Basic Salary'])],
        ['Overtime Pay', format_currency(row['Overtime'])],
        ['Allowance', format_currency(row['Allowance'])],
    ]
    
    deductions_data = [
        ['DEDUCTIONS', 'AMOUNT (KES)'],
        ['PAYE Tax', format_currency(row['PAYE Tax'])],
        ['SHA', format_currency(row['SHA'])],
        ['NSSF', format_currency(row['NSSF'])],
        ['Penalties', format_currency(row['Penalties'])],
        ['Other Deductions', format_currency(row['Deductions'])],
    ]
    
    # Style for tables
    table_style = TableStyle([
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONT', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.9, 0.9, 0.9)),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
    ])
    
    # Draw earnings table
    earnings = Table(earnings_data, colWidths=[140, 100])
    earnings.setStyle(table_style)
    earnings.wrapOn(c, width, height)
    earnings.drawOn(c, 30, height - 380)
    
    # Draw deductions table
    deductions = Table(deductions_data, colWidths=[140, 100])
    deductions.setStyle(table_style)
    deductions.wrapOn(c, width, height)
    deductions.drawOn(c, width/2 + 30, height - 380)
    
    # Net pay section with highlight
    net_pay = float(row['Net Salary'])
    c.setFillColor(colors.Color(0.95, 0.95, 1.0))  # Light blue background
    c.roundRect(30, height - 440, width - 60, 40, 5, stroke=1, fill=1)
    c.setFillColor(header_color)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, height - 420, "NET PAY:")
    c.setFont("Helvetica-Bold", 14)
    c.drawRightString(width - 40, height - 420, f"KES {format_currency(net_pay)}")
    
    # Footer
    c.setFillColor(text_color)
    c.setFont("Helvetica", 8)
    footer_text = [
        "This is a computer generated payslip and does not require signature.",
        f"Generated on: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}",
        "For any queries, please contact HR department at hr@enacoach.co.ke"
    ]
    
    y_position = 50
    for text in footer_text:
        c.drawCentredString(width/2, y_position, text)
        y_position -= 15
    
    # Save PDF
    c.showPage()
    c.save()
    buffer.seek(0)
    
    # Encrypt PDF
    reader = PdfReader(buffer)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    # Use last 4 digits of employee ID or default PIN
    pin = str(row['pin']) if 'pin' in row and pd.notna(row['pin']) else "1234"
    writer.encrypt(user_password=pin, owner_password=pin)
    
    encrypted_pdf = BytesIO()
    writer.write(encrypted_pdf)
    encrypted_pdf.seek(0)
    
    # Prepare email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = row['Email']
    msg['Subject'] = f"Payslip for {selected_month} - ENA COACH LTD"
    
    # Email body
    email_body = f"""Dear {row['Name']},

Please find attached your payslip for {selected_month}.

To open the PDF, use your PIN: {pin}

Note: This is an automated email. Please do not reply to this email address.
For any queries regarding your payslip, please contact the HR department.

Best regards,
HR Department
ENA COACH LTD
"""
    
    msg.attach(MIMEText(email_body, 'plain'))
    
    # Attach PDF
    attachment = MIMEApplication(encrypted_pdf.read(), _subtype="pdf")
    filename = f"Payslip_{row['Name'].replace(' ', '_')}_{selected_month}.pdf"
    attachment.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(attachment)
    
    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)

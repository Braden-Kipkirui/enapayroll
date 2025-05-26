from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from datetime import datetime

def generate_and_send_payslip(row, sender_email, sender_password, selected_month):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # Set up styles
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        'Header',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#2c3e50'),
        spaceAfter=14
    )
    subheader_style = ParagraphStyle(
        'Subheader',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=colors.HexColor('#34495e'),
        spaceAfter=6
    )
    normal_style = styles['Normal']
    
    # Company header with logo placeholder
    c.setFillColor(colors.HexColor('#2c3e50'))
    c.setFont("Helvetica-Bold", 16)
    c.drawString(30, height - 40, "ENA COACH LTD")
    
    c.setFillColor(colors.HexColor('#7f8c8d'))
    c.setFont("Helvetica", 9)
    c.drawString(30, height - 60, "KPCU, Nairobi Kenya")
    c.drawString(30, height - 75, "Phone: +254 709 832 000")
    c.drawString(30, height - 90, "Email: info@enacoach.co.ke")
    
    # Payslip info box
    c.setFillColor(colors.HexColor('#3498db'))
    c.rect(width - 180, height - 100, 150, 60, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(width - 35, height - 50, f"PAY DATE: {datetime.now().strftime('%d %b %Y')}")
    c.drawRightString(width - 35, height - 70, "PAY TYPE: Bank Transfer")
    c.drawRightString(width - 35, height - 90, f"PERIOD: {selected_month}")
    
    # Employee information section
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 130, "EMPLOYEE PAYSLIP")
    
    c.setFont("Helvetica", 11)
    c.drawString(30, height - 155, f"Employee ID: {row.get('Employee ID', 'N/A')}")
    c.drawString(30, height - 175, f"Name: {row['Name']}")
    c.drawString(30, height - 195, f"Email: {row['Email']}")
    c.drawString(width - 200, height - 155, f"Department: {row.get('Department', 'N/A')}")
    c.drawString(width - 200, height - 175, f"Position: {row.get('Position', 'N/A')}")
    
    # Salary breakdown section
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30, height - 225, "SALARY BREAKDOWN")
    
    # Table data
    table_data = [
        ['Description', 'Amount (KES)'],
        ['Basic Salary', f"{row.get('Basic Salary', 0):,.2f}"],
        ['Overtime Pay', f"{row.get('Overtime', 0):,.2f}"],
        ['Allowance', f"{row.get('Allowance', 0):,.2f}"],
        ['Bonus', f"{row.get('Bonus', 0):,.2f}"],
        ['', ''],
        ['PAYE Tax', f"-{row.get('PAYE Tax', 0):,.2f}"],
        ['SHA', f"-{row.get('SHA', 0):,.2f}"],
        ['NSSF', f"-{row.get('NSSF', 0):,.2f}"],
        ['Penalties', f"-{row.get('Penalties', 0):,.2f}"],
        ['Other Deductions', f"-{row.get('Deductions', 0):,.2f}"],
        ['', ''],
        ['NET PAY', f"{row.get('Net Salary', 0):,.2f}"],
    ]
    
    # Create table
    t = Table(table_data, colWidths=[300, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#f8f9fa')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e0e0e0')),
        ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor('#3498db')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
    ]))
    
    # Draw table
    table_x = 30
    table_y = height - 450
    t.wrapOn(c, width, height)
    t.drawOn(c, table_x, table_y)
    
    # Footer notes
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.HexColor('#7f8c8d'))
    c.drawString(30, 50, "Payment Method: Bank Transfer")
    c.drawString(30, 35, f"Generated on: {datetime.now().strftime('%d %b %Y %H:%M')}")
    c.drawString(30, 20, "This is a computer-generated document and does not require a signature.")
    
    c.showPage()
    c.save()
    buffer.seek(0)

    # PDF encryption
    reader = PdfReader(buffer)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    pin = str(row['pin']) if 'pin' in row and pd.notna(row['pin']) else "1234"
    writer.encrypt(user_password=pin, owner_password=pin)

    encrypted_pdf = BytesIO()
    writer.write(encrypted_pdf)
    encrypted_pdf.seek(0)

    # Email composition
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = row['Email']
    msg['Subject'] = f"{selected_month} Payslip - ENA COACH LTD"

    body = f"""Dear {row['Name']},

Please find attached your payslip for {selected_month}.

Your payslip is password protected. The password is your PIN (default: 1234 if not set).

If you have any questions about your payslip, please contact HR.

Regards,
ENA COACH LTD
"""
    msg.attach(MIMEText(body, 'plain'))

    # Attach PDF
    attachment = MIMEApplication(encrypted_pdf.read(), _subtype="pdf")
    attachment.add_header('Content-Disposition', 'attachment', 
                        filename=f"Payslip_{selected_month}_{row['Name'].replace(' ', '_')}.pdf")
    msg.attach(attachment)

    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)

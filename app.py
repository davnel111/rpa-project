# Insert your python code here
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# --- Question 1: Folder Selection ---
while True:
    directory = single_input("Please select the folder containing the information needed")
    directory = directory.strip('"').strip("'")
    if directory[-1] != '/':
        directory = directory + '/'

    files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']
    valid = True
    for file in files:
        if not os.path.exists(directory + file):
            show_message('Error', f'The given file does not contain the {file} file. Please re-enter file information.')
            valid = False
            break

    if valid:
        break

# --- Question 2: Load Product Data ---
descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for index, row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

# --- Question 3: Extract Order Data Function ---
def get_order_data(invoice_num, filename, directory):
    full_path = directory + filename
    template = openpyxl.load_workbook(full_path)
    selected_sheet = template['OrderForm']
    
    date = selected_sheet['C9'].value
    date_str = f'{date.month}/{date.day}/{date.year}'
    email = selected_sheet['C13'].value
    name = selected_sheet['C14'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date_str
    order_data['Total'] = ''
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f'{city}, {state} {str(zipcode)}'

    i = 0
    rownum = 24
    total = 0
    while True:
        if selected_sheet[f'B{rownum}'].value is None:
            break
        unitnum = selected_sheet[f'B{rownum}'].value
        price = product_prices[unitnum]
        quantity = selected_sheet[f'C{rownum}'].value
        amount = price * quantity
        description = product_descriptions[unitnum]

        order_data[f'Price{i}'] = str(f'${price:,.2f}')
        order_data[f'Quantity{i}'] = str(quantity)
        order_data[f'Amount{i}'] = str(f'${amount:,.2f}')
        order_data[f'Description{i}'] = str(description)
        total += amount
        i += 1
        rownum += 1

    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date_str

# --- Question 4: Create Invoice Function ---
def create_invoice(invoice_num, order_data, directory):
    input_pdf = directory + 'MasterInvoice.pdf'
    output_pdf = directory + f'Invoice {invoice_num}.pdf'
    create_pdf(input_pdf, output_pdf, order_data)

# --- Instructor Note: Prompt for Password ONCE before the email loop ---
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read().strip()
gmail_user = 'davidnelson2002@gmail.com'

# --- Question 5: Send Email Function ---
def send_email(invoice_num, name, email, date, directory):
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""
    message.set_content(body)

    path_to_invoice = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_invoice, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()

# --- Question 6: Automation Loop ---
filenames = os.listdir(directory)
filenames.sort()

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1
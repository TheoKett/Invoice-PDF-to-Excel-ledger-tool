#---------------------------------------------------------------------------------------------------------------
#- Project 1
#---------------------------------------------------------------------------------------------------------------


#---------------------------------------------------------------------------------------------------------------

#- 1. Import any relevant modules

import pdfplumber
from openpyxl import load_workbook
import re

#---------------------------------------------------------------------------------------------------------------

#- 2. Store the pathing of the .pdf

pdf_path = r'C:\Users\theok\Desktop\Python projects\CV Automated script\Unpaid Invoice.pdf'

#---------------------------------------------------------------------------------------------------------------

#- 3. Extract all the text and copy it into Python

with pdfplumber.open(pdf_path) as pdf:
    full_text = ''
    for page in pdf.pages:
        full_text += page.extract_text() + '\n'

print(full_text)

#---------------------------------------------------------------------------------------------------------------
#- This removes the table and all unnecessary details so Python can read it
#- I use the print function throughout this module to make sure everything is running smoothly
#- If something doesn't print then I know that its something in that section that isn't functioning
#---------------------------------------------------------------------------------------------------------------

#- 4. Search the text for the relevant data and store it in Python

D= {}

date_match = re.search(r'Date:\s*(.+)', full_text)

if date_match:
    D= date_match.group(1)

print(D)

#---------------------------------------------------------------------------------------------------------------
#- Very repetitive process.
#- Program searches for words such as Date, Name, INVOICE NUMBER etc.
#- Using this '\s*(.+)'' the program looks for the text following the colon.
#- From there it stores that data inside of the letter, in this case 'D'.
#---------------------------------------------------------------------------------------------------------------

N= {}

name_match = re.search(r'Name:\s*(.+)', full_text)

if name_match:
    N= name_match.group(1)

print(N)



IN= {}

invoice_number_match = re.search(r'INVOICE NUMBER:\s*(.+)', full_text)

if invoice_number_match:
    IN= invoice_number_match.group(1)
    
print(IN)



St= {}

subtotal_match = re.search(r'Subtotal:\s*(.+)', full_text)

if subtotal_match:
    St= subtotal_match.group(1)

print (St)

STR= {}

sales_tax_rate_match = re.search(r'Sales Tax:\s*(.+)', full_text)

if sales_tax_rate_match:
    STR= sales_tax_rate_match.group(1)

print(STR)



T= {}

total_match = re.search(r'Total:\s*(.+)', full_text)

if total_match:
    T= total_match.group(1)

print(T)

#---------------------------------------------------------------------------------------------------------------

#-5. Store all the data from the previous step into Python
transactions=[
    (D, N, IN, St, STR, T)
]

#---------------------------------------------------------------------------------------------------------------
#- Transactions is where Python stores all the data that I want the program to enter into excel, in a tuple list.
#- Tuple= can not be changed- its a fixed sequence.
#---------------------------------------------------------------------------------------------------------------

#-6. Give Python the location of the .xlsx file
workbook=load_workbook('C:\Users\theok\Desktop\Python Invoice Project\General Ledger.xlsx')

#---------------------------------------------------------------------------------------------------------------
#- This string of code just tells the program where it can find the excel file I want to copy the data into.
#---------------------------------------------------------------------------------------------------------------

#-7. Paste the data into the excel file
sheet=workbook['Restock Expenses']

for transaction in transactions:
    sheet.append(transaction)

#---------------------------------------------------------------------------------------------------------------
#- Here i'm telling the program to open the excel sheet and add the data row by row using append.
#- Append function enters data by row- to enter information in columns you must use a different command.
#---------------------------------------------------------------------------------------------------------------

#-8. Save the edited file
workbook.save('General_Ledger.xlsx')

print('Ledger updated successfully.')

#---------------------------------------------------------------------------------------------------------------
#- These final lines of code save the changes made to the excel file.
#- After that action is completed, the program prints a message stating the task was completed successfully. 
#---------------------------------------------------------------------------------------------------------------

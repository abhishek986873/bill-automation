from docx import Document
from docx.shared import Pt

import docx

# ------------------------------------------------------------------------------------------------
mintu_bill = docx.Document()

# ---------------------------------------------------------------------------------------------

# ------------------------------------------------------------------------------------------------

# taking input of rows and column of BILL

bill_row = int(input("• ENTER NUMBER OF ROW = "))
bill_coloumn = 5
new_bill_row = bill_row + 1

table = mintu_bill.add_table(rows=new_bill_row + 1, cols=5)

print(
    "-------------------------------------------------------------------------------------------------------------------")

# ------------------------------------------------------------------------------------------------


# storing work amount

# ------------------------------------------------------------------------------------------------


dmls = 600
metal_free = 1800
pfm = 350
flexible = 0
rpd = 600
ortho = 700
repair = 150
dental_repair = 150
cd = 600
zirconia = 1800

# ------------------------------------------------------------------------------------------------


# ADDING HEADING IN THE FIRST ROW OF MINTU TABLE
heading_row = table.rows[0].cells

heading_row[0].text = "S.NO"
heading_row[1].text = "DATE"
heading_row[2].text = "UNIT"
heading_row[3].text = "WORK"
heading_row[4].text = "AMOUNT"

# ------------------------------------------------------------------------------------------------


# filling the SERIAL NUMBER COLUMN


# ------------------------------------------------------------------------------------------------


i = 1
k = 1
while i < new_bill_row:
    data_row = table.rows[i].cells
    data_row[0].text = str(k)
    k = k + 1
    i = i + 1

# ------------------------------------------------------------------------------------------------


# entering unit in "UNIT" cell


# ------------------------------------------------------------------------------------------------

i = 1

while i < new_bill_row:
    data_row = table.rows[i].cells
    data_row[2].text = str(input("• ENTER UNIT = "))
    i = i + 1

print(
    "-------------------------------------------------------------------------------------------------------------------")

# ------------------------------------------------------------------------------------------------


# entering work in WORK column


# ------------------------------------------------------------------------------------------------


i = 1

while i < new_bill_row:
    data_row = table.rows[i].cells
    data_row[3].text = str(input("• ENTER WORK = "))
    i = i + 1

print("------------------------------------------------------------------------------")

# ------------------------------------------------------------------------------------------------


# entering amount in DATE column


# ------------------------------------------------------------------------------------------------


i = 1

while i < new_bill_row:
    data_row = table.rows[i].cells
    data_row[1].text = str(input("• ENTER DATE = ") + "-02-2023")
    i = i + 1

print(
    "-------------------------------------------------------------------------------------------------------------------")

# Add a row to the end of the table
new_row = table.add_row()

mintu_bill.save("mint_bill.docx")

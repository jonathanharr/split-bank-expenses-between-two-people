import pandas as pd
import math
import openpyxl
import os
from expense import Expense

class FileNotFoundError(Exception):
    pass

script_directory = os.path.dirname(os.path.abspath(__file__))

input_filename = None
old_input_filename = 'old_input.csv'

for file in os.listdir(script_directory):
    if file.startswith("PERSONKONTO"):
        input_filename = file
        break

if input_filename is None:
    raise FileNotFoundError("No file starting with 'PERSONKONTO' found in the directory.")
else:
    print(f"Found file: {input_filename}")

df = pd.read_csv(old_input_filename, delimiter=';')
past_expenses = []

index = 1
for index, row in df.iterrows():
    expense = Expense()
    expense.sender = row['Avsändare']
    expense.amount = float(row['Belopp'].replace(',', '.'))
    saldo_value = row['Saldo']
    if isinstance(saldo_value, str):
        expense.current_balance = float(saldo_value.replace(',', '.'))
    else:
        expense.current_balance = 0.0

    expense.date = row['Bokföringsdag']
    expense.title = row['Rubrik']

    past_expenses.append(expense)
    index+=1

df = pd.read_csv(input_filename, delimiter=';')

all_expenses = []

index = 1
for index, row in df.iterrows():
    expense = Expense()
    expense.sender = row['Avsändare']
    expense.amount = float(row['Belopp'].replace(',', '.'))
    saldo_value = row['Saldo']
    if isinstance(saldo_value, str):
        expense.current_balance = float(saldo_value.replace(',', '.'))
    else:
        expense.current_balance = 0.0

    expense.date = row['Bokföringsdag']
    expense.title = row['Rubrik']

    all_expenses.append(expense)
    index+=1


print(f'Past expenses count: {len(past_expenses)}')
print(f'All expenses count: {len(all_expenses)}')

unique_expenses = []
for expense in all_expenses:
    is_unique = True
    expense_repr = expense.__title_amount__()
    for past_expense in past_expenses:
        past_expense_repr = past_expense.__title_amount__()
        if expense_repr == past_expense_repr:
            is_unique = False
            break

    if is_unique:
        unique_expenses.append(expense)
        print(expense)

print(f'Expenses not in past expenses: {len(unique_expenses)}')


wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Utgifter'
sheet['A1'] = 'Checkbox'
sheet['B1'] = 'Saldo'
sheet['C1'] = 'Avsändare'
sheet['D1'] = 'Bokföringsdag'
sheet['E1'] = 'Rubrik'
sheet['F1'] = 'Belopp'
sheet['G1'] = 'Total summa'
sheet['H1'] = 'Delat i 2'

for row_idx, expense in enumerate(unique_expenses, start=2):
    sheet[f'A{row_idx}'] = 'X'
    sheet[f'B{row_idx}'] = expense.current_balance
    sheet[f'C{row_idx}'] = expense.sender
    sheet[f'D{row_idx}'] = expense.date
    sheet[f'E{row_idx}'] = expense.title
    sheet[f'F{row_idx}'] = expense.amount

formula = '=SUMIF(A:A, "X", F:F)'
sheet['G2'] = formula

formula = '=G2/2'
sheet['H2'] = formula

wb.save('generated.xlsx')

print("Excel file created successfully.")

os.remove(old_input_filename)
print(f'{old_input_filename} removed')

os.rename(input_filename, old_input_filename)
print(f'{input_filename} renamed to {old_input_filename}')


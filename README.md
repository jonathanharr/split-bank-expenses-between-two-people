# split-bank-expenses-between-two-people
A simple Python script to create an excel that two people can use to split expenses over time. The script can be used for anyone who uses Nordea's banking services.

## Input

You need to have a previously generated old_input file. If you don't have one, just create one called 'old_input.xlsv'.

This way, the script can compare itself towards your previous expenses and output an Excel file where you can X-mark the expense that should be split.

## Excel file

The output file will contain all new expenses from the previous calculation. You can alter the values, and if you don't want the expense to be added to the calculation, you can remove the X-symbol from the A[index]-cell.

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import datetime


def is_valid_date(year, month, day):
    """
    This function tell us if the date inserted by the user is valid.

    :param year: year inserted.
    :param month: month inserted.
    :param day: day inserted.
    :return: True if date is valid, False if it is not.
    """
    day_count_for_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if year%4==0 and (year%100 != 0 or year%400==0):
        day_count_for_month[2] = 29
    return (1 <= month <= 12 and 1 <= day <= day_count_for_month[month])


wb = load_workbook("Trading_statistics.xlsx")

# Select the proper sheets.
# 'wo' is the sheet were are saved the operations
# 'ws' is the sheet were we will save the Statistics
wo = wb['Operazioni']
ws = wb['Statistiche']

# METRICS OF EXCEL SHEET 'Operazioni'
"""
"ROW START" is the row we actually start to work in excel sheet.
We cannot specify ROW END because we will add new row as we have new operation results.
The column that are meaningful are the ones between 1 and 4 (Operation, Openind date, Close date, result)
Since the first is kept for the heading we start from position two.
"""
ROW_START_OPERATIONS = 2
COLUMN_START_OPERATIONS = 1
COLUMN_END_OPERATIONS = 4

# METRICS OF EXCEL SHEET 'Statistiche'
"""
ROW_START_STATISTICS è la cella da cui partono le statistiche sul foglio excel.
Row distance beetween total statistics, last day statistics and selected day statistics.
It is important to not move those cells on excel sheet.
COLUMN_SHIFT_STATISTICS indicates the cells were we have to insert the statistics of a salected period.
"""
ROW_START_STATISTICS = 5
NUMBER_OF_STATISTICS = 10
ROW_SHIFT_STATISTICS = 15
COLUMN_SHIFT_STATISTICS = 7
COLUMN_STATISTIC = 5

# The currency we are working on excel sheets.
CURRENCY = "$"

# Then we iterate through the rows of the operations done.
rows = wo.iter_rows(min_row=ROW_START_OPERATIONS, min_col=COLUMN_START_OPERATIONS, max_col=COLUMN_END_OPERATIONS)

# We declare 4 lists, to divide and store the information.
Operations = []
Openings = []
Closings = []
Results = []


# We iterate through the tuples of rows and we add the data to a new element of the list
for stocks, open, close, res in rows:
    if stocks.value != None and open.value != None and close.value != None and res.value != None:
        Operations.append(stocks.value)
        Openings.append(open.value)
        Closings.append(close.value)
        Results.append(res.value)
    elif stocks.value != None or open.value != None or close.value != None or res.value != None:
        # we enter in this conditional block if we have a partially filled row (for example if we missed one cell).
        Operations.append(None)
        Openings.append(None)
        Closings.append(None)
        Results.append(None)
        print("Erroneus data at excel rows: {0}".format(len(Operations + 1)))
    else:
        # All the elements ot the rows are empty, no operation needed.
        pass

# We check if all the rows have the same length:
if len(Operations) == len(Openings) and len(Operations) == len(Closings) and len(Operations) == len(Results):
    # All it should be, no need to perform operations or raising any warning.
    pass
else:
    print("WARNING, not all the rows have the same length! Please check 'Operazioni' sheet in excel file")

# Timing data, will be necessary to calculate the statistics in a certain period of time.
# start_day and end_day will be needed to calculate the statistics in a certain period of the time.
last_day = max(Openings)
first_day = min(Openings)
# TODO: inserire da temrminale la data. Altrimenti si deve aprire excel, cambiare data, chiudere e lanciare lo script.
"""
selected_day = input("Enter selected data, press ENTER to select data on excel sheet: ")
day = int(selected_day[0:2])
month = int(selected_day[3:5])
year = int(selected_day[6:10])
selected_day = datetime.date(year, month, day)
if is_valid_date(year, month, day):
    pass
else:
    selected_day = ws['E34'].value
"""
selected_day = ws['E34'].value
start_day = ws['K19'].value
end_day = ws['L19'].value

# TODO: implementare i controlli sulle date. la data di fine deve essere minore della data di inizio ecc

total_operations = len(Operations)

# We calculate the number of operation with gain.
# This rule will be the same for all the variable metrics.
# _l suffix means last day
# _s suffix means selected day
# _p suffix means selected period

total_operations = len(Operations)
total_operations_l = 0
total_operations_s = 0
total_operations_p = 0

gain_operations = 0
gain_operations_l = 0
gain_operations_s = 0
gain_operations_p = 0
win = 0
win_l = 0
win_s = 0
win_p = 0
lose = 0
lose_l = 0
lose_s = 0
lose_p = 0
lose_operations = 0
lose_operations_l = 0
lose_operations_s = 0
lose_operations_p = 0

for index, result in enumerate(Results):
    if result >= 0.0:
        gain_operations += 1
        win += result
        if Openings[index] == last_day:
            total_operations_l += 1
            gain_operations_l += 1
            win_l += result
        if Openings[index] == selected_day:
            total_operations_s += 1
            gain_operations_s += 1
            win_s += result
        if Openings[index] >= start_day and Openings[index] <= end_day:
            total_operations_p += 1
            gain_operations_p += 1
            win_p += result
    else:
        # We calculate the number of operation with loses.
        lose_operations += 1
        lose += result
        if Openings[index] == last_day:
            total_operations_l += 1
            lose_operations_l += 1
            lose_l += result
        if Openings[index] == selected_day:
            total_operations_s += 1
            lose_operations_s += 1
            lose_s += result
        if Openings[index] >= start_day and Openings[index] <= end_day:
            total_operations_p += 1
            lose_operations_p += 1
            lose_p += result

# NOTE: lose is already negative number.
net_income = win + lose
net_income_l = win_l + lose_l
net_income_s = win_s + lose_s
net_income_p = win_p + lose_p

Batting_Average = round((100 * gain_operations/total_operations), 2)
Batting_Average_l = round((100 * gain_operations_l/total_operations_l), 2)
Batting_Average_s = round((100 * gain_operations_s/total_operations_s), 2)
Batting_Average_p = round((100 * gain_operations_p/total_operations_p), 2)

Average_Win = win/gain_operations
Average_Win_l = win_l/gain_operations_l
Average_Win_s = win_s/gain_operations_s
Average_Win_p = win_p/gain_operations_p

Average_Lose = lose/lose_operations
Average_Lose_l = lose_l/lose_operations_l
Average_Lose_s = lose_s/lose_operations_s
Average_Lose_p = lose_p/lose_operations_p

# N.B. Average_Lose is a negative number
Win_loss = Average_Win/(-Average_Lose)
Win_loss_l = Average_Win/(-Average_Lose_l)
Win_loss_s = Average_Win/(-Average_Lose_s)
Win_loss_p = Average_Win/(-Average_Lose_p)

# Write all data in the sheets.
ws.cell(5, COLUMN_STATISTIC).value = total_operations
ws.cell(6, COLUMN_STATISTIC).value = str(round(net_income, 2)) + CURRENCY
ws.cell(7, COLUMN_STATISTIC).value = gain_operations
ws.cell(8, COLUMN_STATISTIC).value = lose_operations
ws.cell(9, COLUMN_STATISTIC).value = str(Batting_Average) + "%"
ws.cell(10, COLUMN_STATISTIC).value = str(round(win, 2)) + CURRENCY
ws.cell(11, COLUMN_STATISTIC).value = str(round(lose, 2)) + CURRENCY
ws.cell(12, COLUMN_STATISTIC).value = str(round(Average_Win, 2)) + CURRENCY
ws.cell(13, COLUMN_STATISTIC).value = str(round(Average_Lose, 2)) + CURRENCY
ws.cell(14, COLUMN_STATISTIC).value = str(round(Win_loss, 2)) + CURRENCY

# Last day statistics
ws.cell(5 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = total_operations_l
ws.cell(6 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(net_income_l, 2)) + CURRENCY
ws.cell(7 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = gain_operations_l
ws.cell(8 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = lose_operations_l
ws.cell(9 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(Batting_Average_l) + "%"
ws.cell(10 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(win_l, 2)) + CURRENCY
ws.cell(11 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(lose_l, 2)) + CURRENCY
ws.cell(12 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(Average_Win_l, 2)) + CURRENCY
ws.cell(13 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(Average_Lose_l, 2)) + CURRENCY
ws.cell(14 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(Win_loss_l, 2)) + CURRENCY


# Selected day statistics
ws.cell(5 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = total_operations_s
ws.cell(6 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(net_income_s, 2)) + CURRENCY
ws.cell(7 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = gain_operations_s
ws.cell(8 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = lose_operations_s
ws.cell(9 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(Batting_Average_s) + "%"
ws.cell(10 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(win_s, 2)) + CURRENCY
ws.cell(11 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(lose_s, 2)) + CURRENCY
ws.cell(12 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(Average_Win_s, 2)) + CURRENCY
ws.cell(13 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(Average_Lose_s, 2)) + CURRENCY
ws.cell(14 + 2 * ROW_SHIFT_STATISTICS, COLUMN_STATISTIC).value = str(round(Win_loss_s, 2)) + CURRENCY

# Selected period statistics
ws.cell(5 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = total_operations_p
ws.cell(6 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(round(net_income_p, 2)) + CURRENCY
ws.cell(7 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = gain_operations_p
ws.cell(8 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = lose_operations_p
ws.cell(9 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(Batting_Average_p) + "%"
ws.cell(10 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(round(win_p, 2)) + CURRENCY
ws.cell(11 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(round(lose_p, 2)) + CURRENCY
ws.cell(12 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(round(Average_Win_p, 2)) + CURRENCY
ws.cell(13 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(round(Average_Lose_p, 2)) + CURRENCY
ws.cell(14 + ROW_SHIFT_STATISTICS, COLUMN_STATISTIC + COLUMN_SHIFT_STATISTICS).value = str(round(Win_loss_p, 2)) + CURRENCY

# Save the work
wb.save("Trading_statistics.xlsx")
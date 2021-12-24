import datetime
# STATISTICS TIMING
"""
selected_day select a particular day. Statistics will be calculated for that day.
start_day - end_day. Statistics will be calculated for a particular period of time.
USE THIS FORMAT:
selected_day = 14/12/2021
start_day = 08/12/2021
end_day = 14/12/2021
"""
selected_day = datetime.date(2021, 12, 14)
start_day = datetime.date(2021, 12, 6)
end_day = datetime.date(2021, 12, 24)

# MAXIMUM DAYS LENGTH
"""
If you want to consider only short term operations you have to fix a days limit: 
if closing date - opening date > OPERATION_LENGTH the operation statistics will not be calculated.
Worst case allowed: friday (opening) - saturday- sunday - monday - tuesday (closing)
"""
MAXIMUM_DAYS_OPERATION_LENGTH = 2

# if COPYTRADER_ENABLE = False copytrading operation will not be taken in account.
COPYTRADER_ENABLE = False

# METRICS OF EXCEL SHEET 'Operazioni'
"""
"ROW START" is the row we actually start to work in excel sheet.
We cannot specify ROW END because we will add new row as we have new operation results.
The column that are meaningful are the ones between 1 and 4 (Operation, Openind date, Close date, result)
Since the first is kept for the heading we start from position two.
"""
ROW_START_OPERATIONS = 2
COLUMN_START_OPERATIONS = 1
COLUMN_END_OPERATIONS = 18

# METRICS OF EXCEL SHEET 'Statistiche'
"""
ROW_START_STATISTICS è la cella da cui partono le statistiche sul foglio excel.
Row distance beetween total statistics, last day statistics and selected day statistics.
It is important to not move those cells on excel sheet.
COLUMN_SHIFT_STATISTICS indicates the cells were we have to insert the statistics of a salected period.
"""
ROW_START_STATISTICS = 2
NUMBER_OF_STATISTICS = 10
ROW_SHIFT_STATISTICS = 15
COLUMN_SHIFT_STATISTICS = 7
COLUMN_STATISTIC = 2

# METRICS OF EXCEL SHEET 'Statistiche_per_azione'
ROW_START_STAT_PER_SHARE = 2
COLUMN_SHARES_STAT_PER_SHARE = 1
COLUMN_NUM_OF_SHARES_STAT_PER_SHARE = 2
COLUMN_TOTAL_RES_STAT_PER_SHARE = 3
COLUMN_AVERAGE_RES_STAT_PER_SHARE = 4

# The currency we are working on excel sheets.
CURRENCY = "$"

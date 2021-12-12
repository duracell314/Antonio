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

# METRICS OF EXCEL SHEET 'Statistiche_per_azione'
ROW_START_STAT_PER_SHARE = 2
COLUMN_SHARES_STAT_PER_SHARE = 1
COLUMN_NUM_OF_SHARES_STAT_PER_SHARE = 2
COLUMN_TOTAL_RES_STAT_PER_SHARE = 3
COLUMN_AVERAGE_RES_STAT_PER_SHARE = 4

# The currency we are working on excel sheets.
CURRENCY = "$"

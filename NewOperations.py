import datetime
"""
 Data necessary to compile by the user.
 Insert the data of the operation then insert financial results.
"""
Open_Date = datetime.date(2021, 12, 11)
Close_Date = datetime.date(2021, 12, 11)
Open_Date_2 = datetime.date(2021, 12, 12)
Close_Date_2 = datetime.date(2021, 12, 12)
New_Operations = [
    # Stock trade- Date open - Date close - Result
    ["STC", Open_Date, Close_Date, 100],
    ["STOC2", Open_Date, Close_Date, 100],
    ["STOC3", Open_Date, Close_Date, 100],
]

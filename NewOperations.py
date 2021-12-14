import datetime
"""
 Data necessary to compile by the user.
 Insert the data of the operation then insert financial results.
"""
Open_Date = datetime.date(2021, 12, 14)
Close_Date = datetime.date(2021, 12, 14)
Open_Date_2 = datetime.date(2021, 12, 12)
Close_Date_2 = datetime.date(2021, 12, 12)
New_Operations = [
    # Stock trade- Date open - Date close - Result
    ["PHG", Open_Date, Close_Date, 0.58],
    ["DIS", Open_Date, Close_Date, 0.03],
    ["INTC", Open_Date, Close_Date, 0.10],
    ["NVDA", Open_Date, Close_Date, 3.69],
    ["NVAX", Open_Date, Close_Date, 4.04],
    ["VIR", Open_Date, Close_Date, -6.44],
    ["GOOG", Open_Date, Close_Date, 1.31],
    ["FB", Open_Date, Close_Date, -2.07],
    ["PFE", Open_Date, Close_Date, -10.06],
    ["BNTX", Open_Date, Close_Date, -13.24],
]

"""
Template = [
    # Stock trade- Date open - Date close - Result
    ["STOC", Open_Date, Close_Date, 100],
    ["STOC", Open_Date, Close_Date, 100],
    ["STOC", Open_Date, Close_Date, 100],
    ["STOC", Open_Date, Close_Date, 100],
]
"""



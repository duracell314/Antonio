from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import datetime


def check_if_duplicated() -> bool:
    """
    Thi function check if what we want to insert is already in.
    :return: True is already inserted.
    """
    result = True
    for i in range(ROW_OFFSET, new_row + ROW_OFFSET):
        # We cannot check the data now, for this reason we skip position 2 and 3.
        for j in range(1, 5, 3):
            operation = New_Operations[i - 2]
            data = operation[j - 1]
            # Convert the column number into the letter
            char = get_column_letter(j)
            dummy = ws[char + str(i)].value
            if dummy == data:
                pass
            else:
                result = False
                break
    return result


"""
 As first thing insert financial results here:
"""
Open_Date = datetime.date(2021, 12, 11)
Close_Date = datetime.date(2021, 12, 11)

New_Operations = [
    # Stock traded- Date open - Date close - Result
    ["STC", Open_Date, Close_Date, -1.32],
    ["STOC2", Open_Date, Close_Date, 1.74],
    ["STOC3", Open_Date, Close_Date, 1.94],
]

ROW_OFFSET = 2
wb = load_workbook("Trading_statistics.xlsx")
ws = wb.active

# Calculate necessary rows
new_row = len(New_Operations)


# check is data is already inserted.
duplicated = check_if_duplicated()
# Now we insert financial data in excel blank rows
force_duplication = "N"
if duplicated:
    print("Data seems already inserted, if you want to force insertion type: Y")
    force_duplication = input(": ")
if (duplicated is False) or (force_duplication == "Y"):
    # Insert blank rows.
    ws.insert_rows(2, new_row)
    # ws.delete_rows(2, new_row)
    for row in range(ROW_OFFSET, new_row + ROW_OFFSET):
        for col in range(1, 5):
            operation = New_Operations[row - 2]
            data = operation[col - 1]
            # Convert the column number into the letter
            char = get_column_letter(col)
            ws[char + str(row)].value = data
    print("Data inserted")
else:
    print("Data not inserted")

# Save the result
wb.save("Trading_statistics.xlsx")

from openpyxl import load_workbook
import pandas as pd
import xlwings as xw


workbook = load_workbook("ceonenew.xlsx", data_only=True)
sheet = workbook["Sheet1"]
cell_value = sheet["g57"].value
print(cell_value)

# ---------------write extracted cell value to text file----------------

with open("sample.txt", "w") as outfile:
    outfile.write(cell_value)

# ----------------------read text file to dataframe------------------------
dataframe = pd.read_fwf(r"sample.txt", encoding="latin1", header=None)
df = dataframe[(dataframe[0].str.contains(':'))]
df[0] = df[0].str[3:]
df = df[0].str.split(':', expand=True)
print(df)

print("\n__________________________________________________\n")

# ------------------load workobook with xlwings module --------------------
app = xw.App(visible=False)
workbook = xw.Book("ceonenew.xlsx")
sheet = workbook.sheets['Sheet1']

# first column of dataframe
sheet.range('G15').options(index=False, header=False).value = df[0]
sheet.range('Z15').options(
    index=False, header=False).value = df[1]  # second column

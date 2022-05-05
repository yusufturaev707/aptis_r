import pandas as pd
import os

file = "main/09.04.2022/sheets.xlsx"

xls = pd.ExcelFile(file)
sheets = xls.sheet_names
# n = len(sheets)
results = []

data = pd.read_excel(xls, sheet_name=sheets[0])
print(data)
df = pd.DataFrame(data)
df.fillna(0, inplace=True)
n = df.shape[0]

users = []
for i in range(n):
    id = df.iloc[i, 0]
    fam = df.iloc[i, 1]
    ism = df.iloc[i, 2]

    user = {
        'id': id,
        'fio': f"{ism} {fam}",
    }
    users.append(user)

# for i in users:
#     print(i)
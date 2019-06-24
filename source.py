import os
import glob
import pandas as pd

path_sources = r"./sources/*.xls"
path_output = r"./output/output.xlsx"

data = pd.DataFrame()
for file in glob.glob(path_sources):
    df = pd.read_excel(file, header=None)
    data = data.append(df, ignore_index=True, sort=False)

data = data.dropna(how="all")

if os.path.exists(path_output):
    os.remove(path_output)
with pd.ExcelWriter(path_output) as writer:
    data.to_excel(writer, sheet_name="merged")
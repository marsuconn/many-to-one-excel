import os
import pandas as pd

cwd = os.path.abspath('') 
files = os.listdir(cwd) 
print(cwd)
print(files)

df = pd.DataFrame()

for file in files:
    if file.endswith('.xlsx'):
        df = df.append(pd.read_excel(file, sheet_name='Sheet2'), ignore_index=True) 
        
df.to_excel('combined_files.xlsx')
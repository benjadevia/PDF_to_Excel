# STEP 1: 
# The following commands need to be installed via Terminal
# pip install tabula-py pandas openpyxl

import tabula 
import pandas as pd

# List to store DataFrames from each page
dfs = []

# Read all pages of the PDF file and append the DataFrames to the list
# STEP2: 
#Here, provide the path to the PDF file you want to convert
dfs = tabula.read_pdf("/Users/benjamindeviaservanti/Desktop/Code/A0410006.pdf", pages='all')

# Get the total number of pages
total_pages = len(dfs)

# Range of pages to process
initial_page = 0
final_page = total_pages

# List to store DataFrames from all pages
dfs_total = []

# Iterate over all DataFrames and append them to the list
for df in dfs:
    dfs_total.append(df)

# Concatenate all DataFrames into one
df_final = pd.concat(dfs_total, ignore_index=True)

# Save the combined DataFrame to an Excel file

# STEP 3 
#Here, provide the destination path and also add the desired .xlsx file name
df_final.to_excel('/Users/benjamindeviaservanti/Desktop/Code/output.xlsx', index=False)


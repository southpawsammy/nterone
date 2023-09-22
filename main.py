import pandas as pd 
import numpy as np
import pathlib
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

#library needed to use xlsx files 

#import pip
#pip.main(["install", "openpyxl"])

#takes a sheet and returns a df, removes unnecessary columns 
def create_df(file, sheet):
    df = pd.read_excel(file, sheet_name = sheet)
    df = df[df['Type'] == 'Bill']

    df = df.drop('Num', axis=1)
    df = df.drop('Name', axis=1)
    df = df.drop('Date', axis=1)
    df = df.drop('Memo', axis=1)

    return df 

#takes a dataframe and reformats to show subtotal breakdown
def sum_group_by(df):
    df = df.groupby(['Account']).sum()
    df.loc['Total'] = [df['Amount'].sum()]    
    return df

#takes an input excel file saves an output file to the directory with the Project cost breakdowns sheet by sheet
def output_file(file):
    wb = load_workbook(file)

    for sheet in wb.worksheets:
        sheet_name = sheet.title
        df = create_df(file, sheet_name)
        df = sum_group_by(df)
        del wb[sheet_name]
        wb.create_sheet(title=sheet_name)

        for r in dataframe_to_rows(df, index=False, header=True):
            wb[sheet_name].append(r)

        wb.save(file)
        df.to_excel("output.xlsx")




files = [f.name for f in pathlib.Path().glob("*.xlsx")]

for file in files:
    output_file(file)




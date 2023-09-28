import pandas as pd 
import numpy as np
import pathlib
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

#library needed to use xlsx files 

#import pip
#pip.main(["install", "openpyxl"])

#takes a dataframe and reformats to show subtotal breakdown
def reformat(df):
    project_name = df['Project'].iloc[0]
    df = df.groupby(['Account']).sum()
    df.loc['Total'] = [df['Amount'].sum()] 
    df.loc['Project Code'] = f"{project_name}"
    return df

#takes a sheet and returns a df, removes unnecessary columns 
def create_df_list(file, sheet):
    df = pd.read_excel(file, sheet_name = sheet)
    df = df[df['Type'] == 'Bill']

    df = df.drop('Num', axis=1)
    df = df.drop('Name', axis=1)
    df = df.drop('Type', axis=1)
    df = df.drop('Date', axis=1)
    df = df.drop('Memo', axis=1)

    #create list of dataframes for each project code in file
    df_list = []
    new_df = [] # this is actually a list but will become a dataframe later
    project = df['Project'].iloc[0] # holding project name from first row to keep track for comparisons 


    for i, row in df.iterrows():
        current_project = df['Project'][i]

        if project != current_project:
            append_df = reformat(pd.DataFrame(new_df))
            df_list.append(append_df)
            project = current_project 
            new_df.clear()
        
        new_df.append(row)

    df_list.append(reformat(pd.DataFrame(new_df))) #appending final DataFrame 
    return df_list 

#takes an input excel file saves an output file to the directory with the Project cost breakdowns sheet by sheet
def output_file(file):
    wb = load_workbook(file)

    for sheet in wb.worksheets:
        sheet_name = sheet.title
        df_list = create_df_list(file, sheet_name)
        
        for df in df_list:
            new_sheet = df['Amount'].iloc[-1]
        #del wb[sheet_name]
        #wb.create_sheet(title=sheet_name)

        #for r in dataframe_to_rows(df, index=False, header=True):
        #    wb[sheet_name].append(r)

        #wb.save('output_' + file)




files = [f.name for f in pathlib.Path().glob("*.xlsx")]

#incrementing through all files present in the nterone folder 
for file in files:
    output_file(file)




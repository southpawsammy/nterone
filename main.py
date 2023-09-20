import pandas as pd 
import numpy as np
import pathlib

#library needed to use xlsx files 

#import pip
#pip.main(["install", "openpyxl"])


files = [f.name for f in pathlib.Path().glob("*.xlsx")]
df = pd.read_excel(files[0], sheet_name= 'QBreport2') #need to fix this sheet_name logic in the future

#filter dataframe to just rows that are Bills 
bill_df = df[df['Type'] == 'Bill']

#remove unnecessary columns
bill_df = bill_df.drop('Num', axis=1)
bill_df = bill_df.drop('Name', axis=1)
bill_df = bill_df.drop('Date', axis=1)
bill_df = bill_df.drop('Memo', axis=1)


#save dataframe and save to directory 
final_xlsx = bill_df.to_excel("output.xlsx")


# nterone
Purpose
This program is meant to take compatible Quickbooks excel files, filter to just invoice rows, and create a summary of Nterone weekly spending. The summary breakdowns the input excel file by Project Code and Account type. The summary is saved as 'output_*'.xlsx in the same folder as the executable.  

*Requirments for running properly*
# 1 - files need to be in .xlsx file format
# 2 - Project codes in input file must be in alphabetical order


*How to run program*
# 1 - add any number of compatible Quickbooks files to the same folder as the .exe
# 2 - double click .exe file
# 3 - the new 'output_*' files will be created in the same folder
# 4 - remove all 'output_*' files before running program again


*Description of how program works* 

Structure:
The program is broken into 6 helper functions that support the main function (output_file). The (output_file) function takes a input excel file and creates and saves an output file in the same folder. This process is repeated until there are no more unprocessed files in the folder.

Helper functions:
reformat: takes a dataframe and reformats to subtotal breakdown
create_df_list_summary: takes a sheet and returns a dataframe, removes unnecessary columns 
create_summary: takes a list of dataframes and creates a summary of all of the projects in list
append_sheet: adds rows from input dataframe to selected workbook
reformat_workbook: change column width and font of worksheet 

Progression (from beginning of output_file function):
# 1 - Load workbook from input file and run a loop on all sheets in workbook
# 2 - In each sheet in the workbook:
    A - run (create_df_list) and reorganize the sheet by Project Code into a list of dataframes 
    B - Use the output of (create_df_list) to create a singular summary dataframe that highlights all Project Code and Account types in the current sheet
    C - Append the summary dataframe to the workbook
    D - Dataframe by dataframe, append all of the dataframes in (df_list) as a new sheet and replace the old sheet
    E - Reformat the font, size, and emphasis on various values in the workbook
    F - Save the workbook as a new file in the current folder
# 3 - Repeat this process for all sheets in the input file (each sheet in the input file should create a new file)
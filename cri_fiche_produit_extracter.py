import pandas as pd 
import openpyxl
import os
import static_info
import general_functions
import xlrd
import datetime
import multiprocessing

from openpyxl.workbook import Workbook
from pathlib import Path

start_time=datetime.datetime.now()

wb=Workbook()
ws=wb.active

result_df=pd.DataFrame()
fp_folder=Path(Path.cwd() / "list")

counter=0
for excel_file in fp_folder.rglob("*.xls*"):
    temp_file_name=str(excel_file.name)
    if "~$" in temp_file_name:
        print("I skipped this file: ", excel_file.name)
        continue
    counter+=1
    parent_folder=excel_file.parts[len(excel_file.parts)-2]
    print("I will start working this file: ", excel_file.parent, " - ", excel_file.name)

    #Here I decide how many sheet do I have
    try:
        temp_wb=pd.ExcelFile(excel_file)
        total_sheet=len(temp_wb.sheet_names)
    except:
        continue

    # Here I decide which sheet should I work on
    if total_sheet==1:
        try:
            df=pd.read_excel(excel_file)
            #print("Dataframe is assigned, WB has only 1 sheet", excel_file.name)
        except:
            pass

        #print("Sample data of current file: ", df[:10])
    else: # If workbook has more than one sheet
        specific_sheet_found=False
        for temp_sh in temp_wb.sheet_names:
            if str(temp_sh.title) in excel_file.name:
                try:
                    df=pd.read_excel(excel_file, sheet_name=temp_sh.title)
                    #print("Dataframe is assigned, WB has only 1 sheet", excel_file.name)
                    specific_sheet_found=True
                    break
                except:
                    #print(f"I couldn't open {excel_file.name}")
                    break

        if not specific_sheet_found:
            df=pd.read_excel(excel_file)

    # Here I loop through specific sheet of current excel file, and assign necessary values into my destination file/dataframe
    for i, rows in df.iterrows():

        # Create a list that tracks if column(key) info is found or not
        found_items=[]
        for j, cols in rows.iteritems():
            column_no=df.columns.get_loc(j)
            xvalue=str(cols)
            if str(xvalue)=="nan":
                continue
            # Check if cell value is equal to column titles
            for key in static_info.title:
                if xvalue.__contains__(key):
                    #print ("Match found: ", xvalue, " contains this: ", key)
                    if static_info.title[key][0]=="regular":
                        #print("I will loop from j=",column_no, " to 25")
                        current_cell_value=df.iloc[i][column_no]
                        #print("Current cell value: ", current_cell_value)
                        if column_no>25:
                            continue
                        desired_value_found=False
                        for to_right in range(column_no+1, 25):
                            try:
                                currently_active_iloc_value=df.iloc[i][to_right]
                                #print("I am inside to_right column no: ", to_right, " and I will analyze this cell: ", currently_active_iloc_value)
                                if general_functions.is_there_value(currently_active_iloc_value):
                                    # print("This cell is what I need, I chose this: ", currently_active_iloc_value)
                                    static_info.result_dict[int(static_info.title[key][1])].append(currently_active_iloc_value)
                                    desired_value_found=True
                                    break
                            except:
                                pass
                        
                    else: #if it is not regular
                        try:
                            english_desc_found=False
                            for to_right in range(column_no, 25):
                                currently_active_iloc_value=df.iloc[i+1][to_right]
                                #print("I will loop from j=",column_no, " to 25")
                                current_cell_value=df.iloc[i][column_no]
                                #print("Current cell value: ", current_cell_value)
                                #print("I am inside to_right column no: ", to_right, " and I will analyze this cell: ", df.iloc[i+1][to_right])
                                if general_functions.is_there_value(currently_active_iloc_value):
                                    if len(currently_active_iloc_value)>7:
                                        if english_desc_found==False:
                                            # print("Analyzing this cell: ", currently_active_iloc_value)
                                            static_info.result_dict[int(static_info.title[key][1])].append(currently_active_iloc_value)
                                            english_desc_found=True 
                                        else: # if english description is found
                                            static_info.result_dict[int(static_info.title[key][2])].append(currently_active_iloc_value)
                                            english_desc_found=True
                                            break
                        except:
                            pass
    
    # Independent from loop, update folder name and excel file name for once here
    static_info.result_dict[13].append(parent_folder)
    static_info.result_dict[14].append(excel_file.stem)

    # To make all columns same size, I write "" into every blank cell here
    for colno in range(1 , len(static_info.result_dict)):
        if len(static_info.result_dict[colno])<counter:
            static_info.result_dict[colno].append("")

# Here I create empty dataframe and assign result dictionary into in it
result_df=pd.DataFrame()
for i in range(0, len(static_info.result_dict)):
    result_df[static_info.formal_title[i]]=static_info.result_dict[i+1]

result_df.to_excel("result.xlsx")

end_time=datetime.datetime.now()
total_time_spent=end_time-start_time
print("Total spent time is: ", total_time_spent)



# if __name__ == '__main__':
#     multiprocessing.freeze_support()
#     print("I am inside multitest.")
#     pool=multiprocessing.Pool(processes=7)
#     pool.map(multi_test,range(1))
#     pool.close()
import pandas as pd
import os

def XLScombine(path, shname = None, cols = None):
    count = 0
    df = []
    dir = path # Modify this. This is the path to the Excel file
    file_list = os.list(path)

    for i in file_list:
        data = pd.read_excel(dir + i, sheetname=shname, header=0, usecols=cols)  # Modify the sheetname argument based on how your sheets are named
        df.append(data)
        count = count + 1
    final = f"{dir}Result_{shname}.xls"  # Path to the file in which new sheet will be saved.
    df = pd.concat(df)
    df.to_excel(final)
    return final


ui_cols = None
ui_path = input("Enter path to parent folder:   ")
ui_shname = input("\n\nEnter the shared sheet name to be combined:  ")
ui_cols_q = input("Combine all columns? (Y/N)")
if ui_cols_q is "N":
    ui_cols = input("Please enter a comma separated list of Excel column letters and column ranges(e.g. “A: E” or “A, C, E: F”)\n\n")
elif ui_cols_q is "Y":
    pass
else:
    print("::INVALID INPUT::\nContinuing using all columns")
    pass

xls_result_path =XLScombine(ui_path, ui_shname, ui_cols)
print(f"The concatenated Excel document is located at {xls_result_path}")
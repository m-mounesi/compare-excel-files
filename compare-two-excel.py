#########################################################
## Compare Two Excel files 
#########################################################
#########################################################
## Author: Sanny Patel
## Version: 1.0.1
## Mmaintainer: Sanny Patel
## Email: p_sunny27@live.com, sanny.patel@capgemini.com
#########################################################

import argparse
import pandas as pd
import sys
import xlsxwriter
from pathlib import Path
from datetime import datetime
from multiprocessing import Process

def excel_diff(df_OLD, df_NEW, path_OLD, path_NEW, suffix):

    print("\nStart Time for ", suffix, " :", datetime.now().strftime("%H:%M:%S"))
    
    ## Perform Diff
    dfDiff = df_NEW.copy()
    newRows = []

    cols_OLD = df_OLD.columns
    cols_NEW = df_NEW.columns
    sharedCols = list(set(cols_OLD).intersection(cols_NEW))
    
    ## Compare Rows
    for row in dfDiff.index:
        if (row in df_OLD.index) and (row in df_NEW.index):
            for col in sharedCols:
                value_OLD = df_OLD.loc[row,col]
                value_NEW = df_NEW.loc[row,col]
                if value_OLD==value_NEW:
                    dfDiff.loc[row,col] = df_NEW.loc[row,col]
                else:
                    dfDiff.loc[row,col] = ('{}→{}').format(value_OLD,value_NEW)
        else:
            newRows.append(row)

    dfDiff = dfDiff.sort_index().fillna('')

    print("\nProcessing time for ", suffix, ":", datetime.now().strftime("%H:%M:%S"))
    print("\nTotal new rows in new report for", suffix, ": ", len(newRows), "\n")
    
    print("\nSaving the outputs in new excel for ", suffix, "... Current Time: ", datetime.now().strftime("%H:%M:%S")) 
    
     
    ## Save output and format
    fname = '{} vs {} - {}.xlsx'.format(path_OLD.stem,path_NEW.stem,suffix)
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')

    dfDiff.to_excel(writer, sheet_name='DIFF', index=True)
    df_NEW.to_excel(writer, sheet_name=path_NEW.stem, index=True)
    df_OLD.to_excel(writer, sheet_name=path_OLD.stem, index=True)

    ## Get xlsxwriter objects
    workbook  = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.set_default_row(15)

    ## Define formats
    highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    new_fmt = workbook.add_format({'bg_color': '#32CD32', 'bold':True})

    
    ## Highlight changed cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                            'criteria': 'containing',
                                            'value':'→',
                                            'format': highlight_fmt})
    ## Highlight new rows
    for row, row_data in enumerate(dfDiff.index):
        if row_data in newRows:
            worksheet.set_row(row+1, 15, new_fmt)
    
    ## Saving Workbook
    writer.save()
    print("\nSaved Excel sheet for ", suffix, "in ", fname ,".")
    print("\nEnd Time for {} : {}".format(suffix,datetime.now().strftime("%H:%M:%S")))

def main(args):

    print("Start Time of Program :", datetime.now().strftime("%H:%M:%S"), "\n")

    ## Set Variables
    #indexColName = sys.argv[3]
    
    #path_OLD = Path(sys.argv[1])
    #path_NEW = Path(sys.argv[2])
    indexColName = args.index
    
    path_OLD = Path(args.old_excel)
    path_NEW = Path(args.new_excel)
 
    ## Starting Multiprocessing for Environment sheets
    procs = []

    for (env, sheet_num) in zip(["PROD", "TEST"],[0, 1]):
        
        ## Reading Sheets from Excel files
        print("\nReading sheet for ", env, "...")
        df_OLD = pd.read_excel(path_OLD, sheet_name=sheet_num, index_col=indexColName).fillna(0)
        df_NEW = pd.read_excel(path_NEW, sheet_name=sheet_num, index_col=indexColName).fillna(0)
        
        proc = Process(target=excel_diff, args=(df_OLD, df_NEW, path_OLD, path_NEW, env))
        procs.append(proc)
        proc.start()

    ## Complete the processes
    for proc in procs:
        proc.join()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Compare Two Excel Workbooks."
    )
    parser.add_argument(
        '-o',
        '--old-excel',
        metavar='old_file.xlsx',
        help='Old Excel File path',
        required=True
    )
    parser.add_argument(
        '-n',
        '--new-excel',
        metavar='new_file.xlsx',
        help='New Excel file path',
        required=True
    )
    parser.add_argument(
        '-i',
        '--index',
        metavar='Account Number',
        help='Common Index Column',
        required=True
    )
    args = parser.parse_args()
    main(args)
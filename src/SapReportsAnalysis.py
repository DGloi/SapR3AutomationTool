from importlib.machinery import PathFinder
import pandas as pd
import numpy as np
import glob as glob
import time
import os

#Variable directories set up
pathin=r"input"
pathout=r"output"
path_mapping=r"mapping.xlsx"

#File aquisition and column name modification for accommodation:

pnl_files=glob.glob(pathin + "\*.txt")
CM_mapping=pd.read_excel(path_mapping,sheet_name="CM")
PM_mapping=pd.read_excel(path_mapping,sheet_name="PM")
timestr = time.strftime("%Y%m%d")

#some columns in the report were not needed and needed renaming
del_cols=[0,8]
column_names=["col1","col2","col4","col5","col6","col7","col8","col9","col10"]
cols_format=['col1','col2','col3','col4']

#Loop over all txt files found in path input, suppose only sap exports will be located there
for f in pnl_files:

    #Preparing outputfile name and directory
    filename=os.path.basename(f)
    filename_without_ext=os.path.splitext(filename)[0]
    output=os.path.join("\\",str(pathout),filename_without_ext) 
    df1=pd.read_csv(f,sep="|", skiprows=11,names=column_names,skipinitialspace = True)

    #Data Clearing and formating
    df1.drop(df1.columns[del_cols],axis=1,inplace=True)
    df1.dropna(axis=0,how="any",inplace=True)
    df1["LE"]=df1["LE"].astype(int)

    for col in cols_format:
        df1[col]=[x.replace(',', '') if ',' in x else x for x in df1[col]]
        df1[col]=[x.replace(' ', '') if ' ' in x else x for x in df1[col]]
        df1.loc[df1[col].str.endswith("-"), col] = ("-" + df1.loc[df1[col].str.endswith("-"), col].str.strip("-"))
        df1[col]=df1[col].astype(float)

    #Setting up  index for relevant data (possible breaks in ledger)

    df1.loc[(df1['col1'] >= 10) | (df1['col2'] <= -10), 'ISSUE'] = 'ISSUE TYPE 1'
    df1.loc[(df1['col3'] >= 10) | (df1['col4'] <= -10),'ISSUE'] = 'ISSUE TYPE 1'
    df1.loc[(df1['col5'] !=0)&(df1['col6'] !=0) ,'ISSUE'] ='ISSUE TYPE 2'

    #Filtering on the break index
    filter=['ISSUE TYPE 2','ISSUE TYPE 1']
    df1=df1[df1[ 'ISSUE' ].isin(filter)]
    print(df1)

    #Scenario variation on file type:

    if "file type 1" in f:
        df2=pd.merge(df1,
                    CM_mapping,
                    on="COMPANY",
                    how="left")
        df2.to_excel(output+" "+timestr+".xlsx",index=False, sheet_name="ACCOUNTING REPORT")

    if "file type 2"in f:
        df2=pd.merge(df1,
                    PM_mapping,
                    on="LE",
                    how="left")
        df2.to_excel(output+" "+timestr+".xlsx",index=False, sheet_name="ACCOUNTING REPORT")
    else:

        print("All ACCOUNTING REPORT processed")
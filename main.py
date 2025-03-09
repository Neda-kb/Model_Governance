## Required Libraries
# sqlalchemy
# sqlachemy-access
# os
# pandas
# openpyxl
# datetim
# configparser
# logging

## Required database : 'Modell.accdb'

## Required files :
# MDB_overview_template.xlsx : template file
# MDB_Modelle.xlsx

## Project Structure
# It consists of the following files:
# main.py: The main file that coordinates the entire process.
# functions.py

## All inputs are initialized in config

import os
import sqlalchemy as sa
import pandas as pd
import openpyxl as xlsx
import re
import sys
from datetime import datetime
path_functions=r""
sys.path.append(path_functions)
import functions as f

#%%
# Create a config folder
# If it has been already created it will show FileExistsError
try:
    config_folder=f.creat_folder(path=r"",folder_name="configfile")
except:
    print('Config folder already exists.')

#%%
# Create a config file
# If it has been already created it will show FileExistsError.
# It should run just for first time this script is run and for next time you should ignore it because if you run it again everything you already changed in the config file will reset
try:
    cfg_file_path = r"configfile\export_from_MDB.conf"
    f.creat_configfile(cfg_file=cfg_file_path)
except:
    print('Config file already exists.')

#%%
# In the last step confige file was created in the configfile folder
# everything has been initialized just open it and check

#%%
# Read config file
cfg=f.config(cfg_file=r"configfile\export_from_MDB.conf")

#%%

current_time=datetime.now().strftime('%Y-%m-%d_%H-%M')
# Extracting paths and filenames from the configuration
output_path = cfg["Output"]["Path"]
int_output_name =( cfg["Output"]["Filename"])
mddb_path = cfg["Database"]["Path"]
mddb = cfg["Database"]["Filename"]
contact_person=cfg["Contact_person"]["Path"]
sheet_name=cfg["Contact_person"]['sheet_name']
template_path=cfg['DEFAULT']['path']

#%%
# Output name
# Split the format of output name
file_name,file_extention=os.path.splitext(int_output_name)
output_name=f'{file_name}_{current_time}{file_extention}'
output_filename=os.path.join(output_path, output_name)

#%%
# DB connection
connection_url = sa.engine.url.URL.create(
        "access+pyodbc",
        query={"odbc_connect":
               r""
               + os.path.join(mddb_path, mddb)
               + ";"})
engine = sa.create_engine(connection_url)

#%%
# Take data from DB and excel

# OverviewModelsFindingsMeasures
overview_query="SELECT * FROM OverviewModelsFindingsMeasures"
overview=f.read_table(overview_query,engine)

# Contact.xlsx
cantact_p=pd.read_excel(contact_person,sheet_name)

#%%
# Select coulmns
# Select 'model owner Name' and 'model validator Name' from Contact.xlsx
selected_columns_contact=['model number','model owner Name','Mmodel validator Name']
df_contact=cantact_p[selected_columns_contact]

#%%
# Merge data

# Merge df_overview with cantact_p whiout selecting columns in df_overview
df_int_suffixes=df_contact.merge(overview,how='right', right_on='Modellnummer',left_on='model number', suffixes=('','drop'))
# Drop column with suffixes
df_int=df_int_suffixes.loc[:,~df_int_suffixes.columns.str.endswith('_drop')]
df_int_2=df_int.drop(['model number'],axis=1)

#%%
# Select active models
df_1=df_int_2[df_int_2.iloc[:,9]=='active']
df = df_1[(df_1.iloc[:,49]=='Validierung') | (df_1.iloc[:,49]=='Governance')]

#%%
# Export data to Excel
def clean_string(df):
    '''
   This function finds any special characters in df that they can't be written in the Excel.
    '''
    regex = r'[\x00-\x1F\x7F-\x9F\r\n\t]'
    if isinstance(df, str):
       return re.sub(regex,'',df)
    return df

# Df whithout special characters
df_clean=df.applymap(clean_string)

# Write in Excel
wb=xlsx.load_workbook(template_path)
sheet=wb['Overview']
for row_idx,row_data in enumerate(df_clean.values,start=3):
    for col_idx ,value in enumerate(row_data,start=3):
        sheet.cell(row=row_idx,column=col_idx,value=value)
wb.save(output_filename)

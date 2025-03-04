# Required Libraries
# sqlalchemy
# sqlachemy-access - for installing a new library in a new environment please review Virtual Environment .html.
# os
# pandas
# openpyxl
# datetim

## Required database : 'ModellDBF.accdb'

## Required files which both are stored in 'U:\Risk_Control\ALLE\MODELL-DB\Tracking Validation Activities' :
# MDB_overview_template.xlsx : template file
# MDB_Modelle_Ansprechpartner.xlsx

## Project Structure
# It consists of the following files:
# main.py: The main file that coordinates the entire process.
# functions.py

## All inputs are initialized in config

## Output file will be stored in 'U:\Risk_Control\ALLE\MODELL-DB\Tracking Validation Activities'

import os
import sqlalchemy as sa
import pandas as pd
import openpyxl as xlsx
from datetime import datetime
from src import functions as f

#%%
# By running this part, it will create a folder which called `configfile` in same directory of Python files and store the config file there.
# All required inputs are initialized in config therefore it doesn't need to change anything, just open and check it.
# `Notic:` if desired to change anything in there, Be careful next time after changing this part should not run because all changes will reset.

# Create a config folder
# If it has been already created it will show FileExistsError
try:
    config_folder=f.creat_folder(path=r"",folder_name="configfile")
except:
    print("config folder already exists.")

# Create a config file
# If it has been already created it will show FileExistsError.
cfg_file_path = r"configfile\export_from_MDB.conf"
f.creat_configfile(cfg_file=cfg_file_path)

#%%
# Second time(if you change something in configfile )it should be continue from here.
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
               r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
               + os.path.join(mddb_path, mddb)
               + ";"})
engine = sa.create_engine(connection_url)

#%%
# Take data from DB and excel

# qryÜbersichtModelleFeststellungenMaßnahmen
overview_query="SELECT * FROM qryÜbersichtModelleFeststellungenMaßnahmen"
overview=f.read_table(overview_query,engine)

# MDB_Modelle_Ansprechpartner.xlsx
cantact_p=pd.read_excel(contact_person,sheet_name)

#%%
# Select coulmns
# Select 'Modellverantwortlicher Name' and 'Modellvalidierer Name' from MDB_Modelle_Ansprechpartner.xlsx
selected_columns_contact=['Modell-\nnummer','Modellverantwortlicher Name','Modellvalidierer Name']
df_contact=cantact_p[selected_columns_contact]

#%%
# Merge data

# Merge df_overview with cantact_p whiout selecting columns in df_overview
df_int_suffixes=df_contact.merge(overview,how='right', right_on='Modellnummer',left_on='Modell-\nnummer', suffixes=('','drop'))
# Drop column with suffixes
df_int=df_int_suffixes.loc[:,~df_int_suffixes.columns.str.endswith('_drop')]
df_int_2=df_int.drop(['Modell-\nnummer'],axis=1)

#%%
# Select active models
df_1=df_int_2[df_int_2.iloc[:,9]=='aktiv']
df = df_1[(df_1.iloc[:,49]=='Validierung') | (df_1.iloc[:,49]=='Governance')]

#%%
# Export data to Excel

wb=xlsx.load_workbook(template_path)
sheet=wb['Overview']
for row_idx,row_data in enumerate(df.values,start=3):
    for col_idx ,value in enumerate(row_data,start=3):
        sheet.cell(row=row_idx,column=col_idx,value=value)
wb.save(output_filename)


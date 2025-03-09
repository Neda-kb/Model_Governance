import os
import configparser
import logging
import pandas as pd
from configparser import ConfigParser

def creat_folder(path,folder_name):
    path_name=os.path.join(path,folder_name)
    if os.path.exists(path_name):
       raise FileExistsError()

    else:
        os.makedirs(path_name)
        print('Folder created successfully.')

def creat_configfile(cfg_file):
    if not os.path.exists(cfg_file):
       # Initialize the ConfigParser
       config = configparser.ConfigParser()
       config['DEFAULT']={'Path':''}
       config['Main']={'Loglevel':'Info'}
       config['Database']={'Path':'','Filename':''}
       config['Output']={'Path':'','Filename':''}
       config['Contact_person']={'path':'','sheet_name':''}


       with open(cfg_file, 'w') as configfile:
           config.write(configfile)
       print("Config file created successfully.")
    else :
       raise FileExistsError()

    return config

def config(cfg_file):
    # Open the config file
    cfg = ConfigParser(
    delimiters=(":","="),
    converters={"list": lambda x: [i.strip() for i in x.split("\n")]}
    )
    cfg.read(cfg_file, encoding="utf-8")

    # Set up logging level
    loglevel = {
        "Info": logging.INFO,
        "Warn": logging.WARNING,
        "Debug": logging.DEBUG
    }
    logging.basicConfig(level=loglevel[cfg["Main"]["Loglevel"]])
    return cfg

def read_table(query, engine):
    return pd.read_sql(query, engine)
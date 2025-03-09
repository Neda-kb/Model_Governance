# Database and Excel Automation Project
## Overview
This project automates data extraction and processing tasks involving an Access database (Modell.accdb) and Excel templates (MDB_overview_template.xlsx, MDB_Modelle.xlsx). It utilizes Python for handling database queries, file operations, and logging.

## Project Structure
- main-2.py - The main script that orchestrates the process.
- functions.py - Contains reusable functions for file management, configuration handling, and logging.
- config.ini - Stores initialization parameters.
- Modell.accdb - The required Microsoft Access database.
- MDB_overview_template.xlsx, MDB_Modelle.xlsx - Required Excel files.

## Requirements
### Dependencies
Ensure you have the following Python libraries installed:
- pip install sqlalchemy sqlalchemy-access pandas openpyxl configparser logging

### Required Files
Before running the project, ensure the following files exist in the working directory:
- Modell.accdb (Access Database)
- MDB_overview_template.xlsx (Excel Template)
- MDB_Modelle.xlsx (Excel Data File)
- config.ini (Configuration File)

## Usage
1. Setup Configuration
- Modify the config.ini file with the appropriate database and file paths.
2. Run the Main Script
`python main-2.py`
3. Logging
- Logs are generated automatically to track execution progress and errors.

## Functions Overview
`functions.py`
This module includes:
- create_folder(path, folder_name): Creates a new folder.
- create_config_file(cfg_file): Generates a default configuration file if it does not exist.
- Additional functions for file and database operations.

## License

This project is intended for internal use. Modify and distribute as needed.
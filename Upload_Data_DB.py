import pandas as pd
from sqlalchemy import create_engine
import argparse
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
log_file = './logs/script_process.log'
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler(log_file)
file_handler.setLevel(logging.DEBUG)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Clear the log file for a fresh start
with open(log_file, 'w'):
    pass

# xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

def create_database(engine, database_name):
    with engine.connect() as conn:
        if engine.dialect.name == 'mysql':
            conn.execute(f"CREATE DATABASE IF NOT EXISTS {database_name}")
        elif engine.dialect.name == 'postgresql':
            conn.execute(f"CREATE DATABASE {database_name}")
        elif engine.dialect.name == 'mssql':
            conn.execute(f"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{database_name}') CREATE DATABASE {database_name}")

def transfer_data_to_database(engine, database_name, excel_file):
    # Connect to the newly created database
    db_engine = create_engine(f"{engine.url}/{database_name}")

    # Read Excel file
    xls = pd.ExcelFile(excel_file)
    
    for sheet_name in xls.sheet_names:
        # Read each sheet into a DataFrame
        df = xls.parse(sheet_name)

        # Transfer data to the database
        df.to_sql(sheet_name, db_engine, if_exists='replace', index=False)
        print(f"Transferred sheet '{sheet_name}' to database '{database_name}'")

def main():
    parser = argparse.ArgumentParser(description="Transfer Excel data to a database.")
    parser.add_argument('db_type', choices=['mysql', 'postgresql', 'mssql'], help="Target database type")
    parser.add_argument('connection_string', help="SQLAlchemy connection string for the database server")
    parser.add_argument('database_name', help="Name of the database to create")
    parser.add_argument('excel_file', help="Path to the Excel (.xlsx) file")
    
    args = parser.parse_args()

    # Connect to the database server
    engine = create_engine(args.connection_string)

    # Create the database
    create_database(engine, args.database_name)

    # Transfer data from Excel to the database
    transfer_data_to_database(engine, args.database_name, args.excel_file)

if __name__ == "__main__":
    main()

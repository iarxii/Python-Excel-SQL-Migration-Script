import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.sql import text
import argparse
import logging
import time

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
    try:
        with engine.connect() as conn:
            if engine.dialect.name == 'mysql':
                conn.execute(text(f"CREATE DATABASE IF NOT EXISTS {database_name}"))
            elif engine.dialect.name == 'postgresql':
                conn.execute(text(f"CREATE DATABASE {database_name}"))
            elif engine.dialect.name == 'mssql':
                conn.execute(text(f"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{database_name}') CREATE DATABASE {database_name}"))
            logger.info(f"Database '{database_name}' created successfully.")
    except Exception as e:
        logger.error(f"Error creating database '{database_name}': {e}")
        raise

def transfer_data_to_database(engine, database_name, excel_file, is_continuous):
    start_time = time.time()
    try:
        # Connect to the newly created database
        db_engine = create_engine(f"{engine.url}/{database_name}")

        # Read Excel file
        xls = pd.ExcelFile(excel_file)
        
        # Check column consistency across all sheets
        reference_columns = None
        table_name = "continuous_table" if is_continuous else None

        for sheet_name in xls.sheet_names:
            try:
                logger.info(f"Validating sheet '{sheet_name}'.")
                # Only read the header
                data_header = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=0)

                if reference_columns is None:
                    reference_columns = data_header.columns.tolist()
                elif data_header.columns.tolist() != reference_columns:
                    if is_continuous:
                        logger.critical(f"Inconsistent columns detected in sheet '{sheet_name}'. Script terminated.")
                        raise ValueError("Inconsistent columns detected in continuous mode.")
            except Exception as e:
                logger.error(f"Error checking sheet '{sheet_name}': {e}")
                raise
            
            # try:
            #     df = xls.parse(sheet_name)

            #     if reference_columns is None:
            #         reference_columns = df.columns.tolist()
            #     elif df.columns.tolist() != reference_columns:
            #         if is_continuous:
            #             logger.critical(f"Inconsistent columns detected in sheet '{sheet_name}'. Script terminated.")
            #             raise ValueError("Inconsistent columns detected in continuous mode.")

            # except Exception as e:
            #     logger.error(f"Error checking sheet '{sheet_name}': {e}")
            #     raise

        # If column consistency checks pass, start transferring data
        for sheet_name in xls.sheet_names:
            try:
                logger.info(f"Processing sheet '{sheet_name}' for data transfer.")
                data = pd.read_excel(excel_file, sheet_name=sheet_name)

                # Process in chunks manually
                chunk_size = 10000
                for i in range(0, len(data), chunk_size):
                    chunk = data.iloc[i:i + chunk_size]
                    logger.info(f"Processing chunk {i // chunk_size + 1} of sheet '{sheet_name}'. Rows: {len(chunk)}")

                    if is_continuous:
                        chunk.to_sql(table_name, db_engine, if_exists='append', index=False)
                        logger.debug(f"Inserted chunk {i // chunk_size + 1} of sheet '{sheet_name}' into table '{table_name}'.")
                    else:
                        table_name = sheet_name
                        chunk.to_sql(table_name, db_engine, if_exists='replace' if i == 0 else 'append', index=False)
                        logger.debug(f"Inserted chunk {i // chunk_size + 1} of sheet '{sheet_name}' into table '{table_name}'.")
                logger.info(f"Completed processing sheet '{sheet_name}'.")
            except Exception as e:
                logger.error(f"Error transferring sheet '{sheet_name}': {e}")
                raise
            
            # try:
                # logger.info(f"Processing sheet '{sheet_name}' for data transfer.")
                # reader = pd.read_excel(excel_file, sheet_name=sheet_name, chunksize=10000)
                # for i, chunk in enumerate(reader):
                    # logger.info(f"Processing chunk {i + 1} of sheet '{sheet_name}'. Rows: {len(chunk)}")

                    # if is_continuous:
                        # chunk.to_sql(table_name, db_engine, if_exists='append', index=False, chunksize=500)
                        # logger.debug(f"Inserted chunk {i + 1} of sheet '{sheet_name}' into table '{table_name}'.")
                    # else:
                        # table_name = sheet_name
                        # chunk.to_sql(table_name, db_engine, if_exists='replace', index=False, chunksize=500)
                        # logger.debug(f"Inserted chunk {i + 1} of sheet '{sheet_name}' into table '{table_name}'.")
                # logger.info(f"Completed processing sheet '{sheet_name}'.")
            # except Exception as e:
                # logger.error(f"Error transferring sheet '{sheet_name}': {e}")
                # raise

        elapsed_time = time.time() - start_time
        logger.info(f"Data transfer completed in {elapsed_time:.2f} seconds. Please review the Database and Table(s).")
        
    except Exception as e:
        logger.error(f"Error transferring data to database '{database_name}': {e}")
        raise

def main():
    parser = argparse.ArgumentParser(description="Transfer Excel data to a database.")
    parser.add_argument('db_type', choices=['mysql', 'postgresql', 'mssql'], help="Target database type")
    parser.add_argument('connection_string', help="SQLAlchemy connection string for the database server")
    parser.add_argument('database_name', help="Name of the database to create")
    parser.add_argument('excel_file', help="Path to the Excel (.xlsx) file")
    parser.add_argument('--continuous', action='store_true', help="Indicate if the workbook contains a continuous table")
    
    args = parser.parse_args()

    try:
        # Connect to the database server
        engine = create_engine(args.connection_string)

        # Create the database
        create_database(engine, args.database_name)

        # Transfer data from Excel to the database
        transfer_data_to_database(engine, args.database_name, "./DATA/"+args.excel_file, args.continuous)
    except Exception as e:
        logger.critical(f"Script failed: {e}")
        raise

if __name__ == "__main__":
    main()

# SQL Database Upload Utility

The script allows the user to select the target database type (SQL Server, MySQL, or PostgreSQL) and transfer data from an Excel file into the chosen database.

### Key Features
1. **Database Creation**: 
   - Supports SQL Server, MySQL, and PostgreSQL.
   - Dynamically creates the database if it doesn’t exist.

2. **Data Transfer**:
   - Reads data from an Excel file.
   - Transfers each sheet into a corresponding table in the database.

3. **Parameterization**:
   - User specifies the database type, connection string, target database name, and Excel file via command-line arguments.

### Usage
1. Install required packages:
   ```bash
   pip install pandas sqlalchemy pymysql psycopg2 pyodbc
   ```
2. Run the script:
   ```bash
   python script_name.py <db_type> <connection_string> <database_name> <excel_file>
   ```

   Example for MySQL:
   ```bash
   python script_name.py mysql "mysql+pymysql://user:password@localhost" my_database data.xlsx
   ```

   Example for PostgreSQL:
   ```bash
   python script_name.py postgresql "postgresql+psycopg2://user:password@localhost" my_database data.xlsx
   ```

   Example for SQL Server:
   ```bash
   python script_name.py mssql "mssql+pyodbc://user:password@localhost/driver=ODBC+Driver+17+for+SQL+Server" my_database data.xlsx
   ```

### Notes
- Ensure the provided connection string is valid and compatible with SQLAlchemy.
- Excel sheet names are used as table names in the database.
- For SQL Server, ensure you have the necessary ODBC driver installed.


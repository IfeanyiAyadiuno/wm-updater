import pyodbc
import pandas as pd
from pathlib import Path


import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="pandas")
# path to Access db
db_path = Path(r"I:\ResEng\Tools\Programmers Paradise\GUI_WM\PCE_WM1.accdb")

conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    rf"DBQ={db_path};"
)

with pyodbc.connect(conn_str) as conn:
    # check tables
    cur = conn.cursor()
    for row in cur.tables(tableType="TABLE"):
        print(row.table_name)

    # grab some data into pandas
    df = pd.read_sql("SELECT TOP 5 * FROM PCE_WM;", conn)
    print(df)
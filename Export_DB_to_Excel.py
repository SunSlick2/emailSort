import sqlite3
import pandas as pd
import os

db_path = 'SMTP_cache.db' # Ensure this matches your filename

def export_to_excel():
    if not os.path.exists(db_path):
        print(f"Error: Database {db_path} not found.")
        return

    try:
        conn = sqlite3.connect(db_path)
        # Read the entire table into a DataFrame
        df = pd.read_sql_query("SELECT * FROM smtp_cache", conn)
        conn.close()

        # Save to Excel
        output_file = 'Cache_Review.xlsx'
        df.to_excel(output_file, index=False)
        print(f"Success! You can now open {output_file} in Excel.")
        os.startfile(output_file) # Automatically opens the file
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    export_to_excel()
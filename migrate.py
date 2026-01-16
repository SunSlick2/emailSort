import pandas as pd
import sqlite3
import os

# --- SETTINGS ---
excel_path = r'C:\Path\To\Your\MailboxTables.xlsm' # Update this
db_path = 'mailbox_cache.db'

def migrate():
    if not os.path.exists(excel_path):
        print("Excel file not found!")
        return

    print("Reading Excel cache...")
    try:
        df = pd.read_excel(excel_path, sheet_name='SMTP_Cache')
        # Clean data: lowercase addresses and remove duplicates
        df['ExchangeAddress'] = df['ExchangeAddress'].str.lower().str.strip()
        df = df.drop_duplicates(subset=['ExchangeAddress'])
        
        print(f"Found {len(df)} entries. Connecting to SQLite...")
        conn = sqlite3.connect(db_path)
        
        # Create table and upload data
        # If table exists, it will replace it
        df.to_sql('smtp_cache', conn, if_exists='replace', index=False)
        
        # Create an index to make lookups lightning fast
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ex_addr ON smtp_cache (ExchangeAddress)")
        
        conn.close()
        print(f"Success! Database created at: {os.path.abspath(db_path)}")
    except Exception as e:
        print(f"Error during migration: {e}")

if __name__ == "__main__":
    migrate()
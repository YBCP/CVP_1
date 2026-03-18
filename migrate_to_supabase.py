"""
Migrar datos existentes de CSV a Supabase.
Ejecutar una sola vez: python migrate_to_supabase.py
"""
import pandas as pd
from supabase import create_client

SUPABASE_URL = "https://sglmtigafinkujzphhld.supabase.co"
SUPABASE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNnbG10aWdhZmlua3VqenBoaGxkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM4NTE2MjMsImV4cCI6MjA4OTQyNzYyM30"
    ".plTgLwR0y5sHi2bcJVgXpPAbb9OYrlmwBcHopK6HY_s"
)

sb = create_client(SUPABASE_URL, SUPABASE_KEY)


def migrate_table(csv_path, table_name, conflict_col="NUM_VISITA"):
    try:
        df = pd.read_csv(csv_path, dtype=str)
        df.columns = df.columns.str.strip()
        records = df.where(pd.notna(df), None).to_dict("records")
        if records:
            sb.table(table_name).upsert(records, on_conflict=conflict_col).execute()
            print(f"OK {table_name}: {len(records)} filas migradas")
        else:
            print(f"WARN {table_name}: CSV vacio, nada que migrar")
    except Exception as e:
        print(f"ERROR {table_name}: {e}")


if __name__ == "__main__":
    migrate_table("data/visitas.csv", "visitas")
    migrate_table("data/resultados.csv", "resultados")
    print("\nMigracion completa.")

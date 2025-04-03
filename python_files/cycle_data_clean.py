import pandas as pd
from sqlalchemy import create_engine

# Load the Excel file
excel_file = r"C:\Users\migli\OneDrive\Desktop\case_study\cycle_december.xlsx"

# Connect to PostgreSQL
engine = create_engine("postgresql://postgres:%40Supermanredson1@localhost:5432/cycle_data")

# Load all sheets into a dictionary
sheets = pd.read_excel(excel_file, sheet_name=None)  # None loads all sheets

for sheet_name, df in sheets.items():
    print(f"Processing sheet: {sheet_name}")

    # Remove negative ride duration values if the column exists
    if 'ride_length' in df.columns:
        df = df[df['ride_length'].notnull()]  # Remove null ride_length values
    
    # Remove duplicate rows
    df = df.drop_duplicates()

    # Insert into PostgreSQL
    table_name = sheet_name.lower().replace(" ", "_")
    df.to_sql(table_name, engine, if_exists="replace", index=False)

    print(f"Successfully imported: {sheet_name} -> {table_name}")

print("All sheets imported successfully!")

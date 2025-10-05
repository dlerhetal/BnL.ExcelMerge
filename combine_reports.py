
import pandas as pd
import os

# Define file paths
expenses_file = 'bnl.expenses.csv'
revenue_file = 'bnl.revenue.csv'
output_file = 'combined_report.xlsx'

# --- Load Data ---
try:
    # Load expenses, skipping footer rows
    expenses_df = pd.read_csv(expenses_file, skipfooter=4, engine='python')
    # Load revenue, skipping footer rows
    revenue_df = pd.read_csv(revenue_file, skipfooter=4, engine='python')
except FileNotFoundError as e:
    print(f"Error: {e}. Make sure the input files are in the same directory as the script.")
    exit()

# --- Data Cleaning ---
# Rename columns for clarity and consistency
expenses_df.rename(columns={'Phase ID': 'Phase Description', 'Phase Description': 'Cost Code Description'}, inplace=True)
revenue_df.rename(columns={'Phase ID': 'Phase Description'}, inplace=True)

# Standardize key columns for merging
key_cols = ['Job ID', 'Cost Code ID', 'Phase Description']

# --- Data Merging ---
# Merge the two dataframes on the key columns
combined_df = pd.merge(expenses_df, revenue_df, on=key_cols, how='outer')

# --- Final Touches ---
# Reorder columns for better readability
desired_order = [
    'Job ID',
    'Cost Code ID',
    'Phase Description',
    'Cost Code Description',
    'Est. Expenses',
    'Act. Expenses',
    'Diff. Expenses',
    'Est. Revenue',
    'Act. Revenue',
    'Diff. Revenue',
    'Est. Exp. Units',
    'Act. Exp. Units',
    'Diff. Exp. Units'
]
combined_df = combined_df[desired_order]

# --- Save to Excel ---
try:
    combined_df.to_excel(output_file, index=False)
    print(f"Successfully combined files into '{output_file}'")
except Exception as e:
    print(f"Error saving to Excel: {e}")


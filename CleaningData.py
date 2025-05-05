import pandas as pd

# Load Excel file
file_path = "Sheet1 (Nike Dataset)_Sheet1.xlsx"
df = pd.read_excel(file_path)

# Display initial data overview
print("Initial Data Shape:", df.shape)
print("Initial Columns:", df.columns.tolist())

# Rename columns for clarity and consistency
df.columns = [col.strip().replace(" ", "_").lower() for col in df.columns]

# Parse 'invoice_date' as datetime
df['invoice_date'] = pd.to_datetime(df['invoice_date'], errors='coerce')

# Drop rows with missing essential values
required_columns = ['invoice_date', 'product', 'region', 'retailer', 'sales_method', 
                    'state', 'price_per_unit', 'total_sales', 'units_sold']
df.dropna(subset=required_columns, inplace=True)

# Remove rows with negative or zero values in key numeric columns
df = df[(df['price_per_unit'] > 0) & (df['total_sales'] > 0) & (df['units_sold'] > 0)]

# Standardize text columns (strip whitespace, title case)
text_columns = ['product', 'region', 'retailer', 'sales_method', 'state']
for col in text_columns:
    df[col] = df[col].astype(str).str.strip().str.title()

# Optional: Reset index after cleaning
df.reset_index(drop=True, inplace=True)

# Save cleaned data
cleaned_file = "cleaned_nike_sales_data.xlsx"
df.to_excel(cleaned_file, index=False)

print("✅ Data cleaning completed. Cleaned file saved as:", cleaned_file)
print("Final Data Shape:", df.shape)

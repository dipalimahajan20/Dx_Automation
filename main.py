import pandas as pd

# Try reading with different header row positions
for header_row in range(10):
    try:
        print(f"Trying header row {header_row}...")
        df = pd.read_excel("employees_raw.xlsx", sheet_name="Data", header=header_row)
        print(f"Columns with header row {header_row}: {df.columns.tolist()}")
        if any('employee' in str(col).lower() or 'id' in str(col).lower() for col in df.columns):
            print(f"Found relevant columns at header row {header_row}")
            break
    except Exception as e:
        print(f"Error with header row {header_row}: {e}")
        continue

# Group by Employee ID
def team_merge(teams):
    return "Multiple Teams" if len(set(teams)) > 1 else list(teams)[0]

print("DataFrame shape:", df.shape)
print("First few rows:")
print(df.head(10))
print("\nColumn names:")
print(df.columns.tolist())

# Check if we need to skip header rows
if df.iloc[0, 0] == 'Data Refreshed on - 4th Sep at 12:30PM (IST)':
    print("\nSkipping header rows...")
    # Look for the actual data headers by examining more rows
    for i in range(min(20, len(df))):
        row_data = df.iloc[i].tolist()
        print(f"Row {i}: {row_data}")
        # Check if this row contains column headers (look for common employee data fields)
        if any(keyword in str(row_data).lower() for keyword in ['employee', 'id', 'name', 'region', 'project', 'team', 'workspace']):
            print(f"Found potential headers at row {i}")
            # Use this row as column names
            df = df.iloc[i+1:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            break

print("\nAfter processing:")
print("DataFrame shape:", df.shape)
print("Column names:", df.columns.tolist())
print("First few rows:")
print(df.head())

# Check if Employee ID column exists
if 'Employee ID' in df.columns:
    result = df.groupby("Employee ID").agg({
        "Employee Name": "first",
        "Sales Region": "first", 
        "Project Name": lambda x: ",".join(set(x)),
        "Employee Role": team_merge,
        "Employee Home Office": lambda x: ",".join(set(x))
    }).reset_index()
    
    # Save cleaned sheet
    result.to_excel("employees_cleaned.xlsx", index=False)
    print("Successfully created employees_cleaned.xlsx")
    print(f"Processed {len(result)} unique employees")
else:
    print("Employee ID column not found. Available columns:", df.columns.tolist())
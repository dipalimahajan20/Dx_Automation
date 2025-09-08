import pandas as pd
import datetime


# Read raw data
df = pd.read_excel("employees_raw.xlsx", sheet_name="Data")

def team_merge(teams):
    return "Multiple Teams" if len(set(teams)) > 1 else list(teams)[0]

def account_merge(accounts):
    return "Multiple Accounts" if len(set(accounts)) > 1 else list(accounts)[0]

def market_merge(markets):
    return "Multiple Markets" if len(set(markets)) > 1 else list(markets)[0]



# Check if Employee ID column exists
if 'Employee ID' in df.columns:
    result = df.groupby("Employee ID").agg({
        "Employee Name": "first",
        "Sales Region": "first",
        "Market": market_merge,
        "Account Name": account_merge,
        "Project Name": team_merge,
        "Employee Role": team_merge,
        "Employee Grade": "first",
        "Employee Home Office": "first",
        "Resource Email ID": "first",
    }).reset_index()
    
    # Replace NaN values with "NA" for better visibility in Excel
    result = result.fillna("NA")
    
    # Save cleaned sheet
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
result.to_excel(f"employees_cleaned_{timestamp}.xlsx", index=True)
    print("Successfully created employees_cleaned.xlsx")
    print(f"Processed {len(result)} unique employees")
else:
    print("Employee ID column not found. Available columns:", df.columns.tolist())
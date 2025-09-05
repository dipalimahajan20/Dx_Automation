import pandas as pd
# Read raw data
df = pd.read_excel("employees_raw.xlsx", sheet_name="Data")

# Group by Employee ID
def team_merge(teams):
    return "Multiple Teams" if len(set(teams)) > 1 else list(teams)[0]

print(df.columns.tolist())
print(df.columns[4])

cell_value = df.iloc[4, 4]
print("Value at E5:", cell_value)

result = df.groupby("Employee ID").agg({
    "Name": "first",
    "Region": "first",
    "Project": lambda x: ",".join(set(x)),
    "Team": team_merge,
    "Workspace": lambda x: ",".join(set(x))
}).reset_index()

# Save cleaned sheet
result.to_excel("employees_cleaned.xlsx", index=False)
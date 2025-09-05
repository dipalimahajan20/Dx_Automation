import pandas as pd
# Read raw data
df = pd.read_excel("employees_raw.xlsx", sheet_name="Data")

# Group by Employee ID
def team_merge(teams):
    return "Multiple Teams" if len(set(teams)) > 1 else list(teams)[0]

result = df.groupby("Employee ID").agg({
    "Name": "first",
    "Region": "first",
    "Project": lambda x: ",".join(set(x)),
    "Team": team_merge,
    "Workspace": lambda x: ",".join(set(x))
}).reset_index()

# Save cleaned sheet
result.to_excel("employees_cleaned.xlsx", index=1)
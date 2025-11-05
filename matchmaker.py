import pandas as pd

# --- Load Excel files ---
operators = pd.read_excel("data/requirements_cleaned.xlsx")  
buildings = pd.read_excel("data/Data Workshop office spreadsheet.xlsx")

# --- Show available operators ---
print("Available operators:\n")
print(operators[["Operator", "MinSize", "MaxSize"]].reset_index())

# --- Select operator ---
row_num = int(input("\nEnter the row number of the operator you want to match: "))
op = operators.iloc[row_num]

# --- Extract operator info ---
operator_name = op["Operator"]
min_size = op["MinSize"]
max_size = op["MaxSize"]

print(f"\n🔍 Selected Operator: {operator_name}")
print(f"Required Size Range: {min_size} – {max_size} sq ft")

# --- Convert building sizes to numeric ---
buildings["size"] = pd.to_numeric(buildings["size"], errors="coerce")

# --- Find matches ---
matches = buildings[
    (buildings["size"] >= min_size) &
    (buildings["size"] <= max_size)
]

# --- Output results ---
if matches.empty:
    print("\n❌ No matching buildings found for this operator.")
else:
    print(f"\n✅ Found {len(matches)} matching buildings:\n")
    print(matches[["city", "post code", "size"]])

    output_file = f"Matches_for_{operator_name}.xlsx"
    matches.to_excel(output_file, index=False)
    print(f"\n💾 Results saved to: {output_file}")

import pandas as pd
import os

# Load Excel files
operators = pd.read_excel("data/requirements_cleaned.xlsx")
buildings = pd.read_excel("data/Data Workshop office spreadsheet.xlsx")

# Ensure output directory exists
os.makedirs("results", exist_ok=True)

# Display available operators
print("Available operators:\n")
print(operators[["Operator", "MinSize", "MaxSize"]].reset_index())

# Select operator
row_num = int(input("\nEnter the row number of the operator you want to match: "))
op = operators.iloc[row_num]

# Extract operator info
operator_name = op["Operator"]
min_size = op["MinSize"]
max_size = op["MaxSize"]

print(f"\nSelected Operator: {operator_name}")
print(f"Required Size Range: {min_size} – {max_size} sq ft")

# Convert building sizes to numeric
buildings["size"] = pd.to_numeric(buildings["size"], errors="coerce")

# Find matches
matches = buildings[
    (buildings["size"] >= min_size) &
    (buildings["size"] <= max_size)
]

# Output results
if matches.empty:
    print("\nNo matching buildings found for this operator.")
else:
    print(f"\nFound {len(matches)} matching buildings.")

    # Choose output folder and filename
    safe_name = str(operator_name).replace("/", "_").replace("\\", "_").strip()
    output_file = os.path.join("results", f"Matches_for_{safe_name}.xlsx")

    # Save results to Excel
    matches.to_excel(output_file, index=False)

    print(f"\nResults saved to: {output_file}")

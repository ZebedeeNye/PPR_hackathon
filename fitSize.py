import pandas as pd
import os


def find_matches(operator_name, operators_path="data/requirements_cleaned.xlsx", buildings_path="data/Data Workshop office spreadsheet.xlsx"):
    """
    Finds building matches for a given operator based on size requirements.

    Parameters
    ----------
    operator_name : str
        The name of the operator to match.
    operators_path : str, optional
        Path to the operators Excel file.
    buildings_path : str, optional
        Path to the buildings Excel file.

    Returns
    -------
    pd.DataFrame
        DataFrame containing matching buildings for the operator.
    """

    # Load Excel files
    operators = pd.read_excel(operators_path)
    buildings = pd.read_excel(buildings_path)

    # Ensure output directory exists
    os.makedirs("results", exist_ok=True)

    # Find the operator row
    op_rows = operators[operators["Operator"].astype(str).str.lower() == operator_name.lower()]
    if op_rows.empty:
        print(f"No operator found with name '{operator_name}'.")
        return pd.DataFrame()

    op = op_rows.iloc[0]
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

    if matches.empty:
        print("No matching buildings found.")
    else:
        print(f"Found {len(matches)} matching buildings.")
        safe_name = str(operator_name).replace("/", "_").replace("\\", "_").strip()
        output_file = os.path.join("results", f"Matches_for_{safe_name}.xlsx")
        matches.to_excel(output_file, index=False)
        print(f"Results saved to: {output_file}")

    return matches


# Example usage:
if __name__ == "__main__":
    operator = input("Enter operator name: ").strip()
    find_matches(operator)

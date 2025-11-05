import pandas as pd
import os


def find_matches(
    operator_name,
    operators_path="data/requirements_cleaned.xlsx",
    buildings_path="data/Data Workshop office spreadsheet.xlsx",
    output_path="results/matched_buildings.xlsx",
):
    """
    Finds building matches for a given operator based on size requirements
    and saves them to a fixed Excel file for use in subsequent processing.

    Parameters
    ----------
    operator_name : str
        The name of the operator to match.
    operators_path : str, optional
        Path to the operators Excel file.
    buildings_path : str, optional
        Path to the buildings Excel file.
    output_path : str, optional
        Path where the matched buildings Excel file will be saved.

    Returns
    -------
    pd.DataFrame
        DataFrame containing matching buildings for the operator.
    """

    # Load Excel files
    operators = pd.read_excel(operators_path)
    buildings = pd.read_excel(buildings_path)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

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

    # Output results
    if matches.empty:
        print("No matching buildings found.")
        return pd.DataFrame()

    print(f"Found {len(matches)} matching buildings.")
    matches.to_excel(output_path, index=False)
    print(f"Results saved to: {output_path}")

    return matches


# Example usage
if __name__ == "__main__":
    operator = input("Enter operator name: ").strip()
    find_matches(operator)

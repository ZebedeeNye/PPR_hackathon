import pandas as pd
import os
from fitSize import find_matches

operators = pd.read_excel("data/requirements_cleaned.xlsx")
buildings = pd.read_excel("data/Data Workshop office spreadsheet.xlsx")

matches = find_matches("Co-op, Sainsuburys, Tesco, Asda, Morrisons")

area = find_area()


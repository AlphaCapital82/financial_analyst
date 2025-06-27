import gspread
from google.oauth2.service_account import Credentials

# Authenticate
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file(
    'C:/Jon/skole/Master thesis/Python/XGBoost - Master thesis/PythonProject1/sheets/credentials.json',
    scopes=scopes
)
client = gspread.authorize(creds)

# Open worksheet
worksheet = client.open_by_key("1bc_ZZ_I2gtwL5kxKylKxIrYsbzk9XuQtCuexm_BDU4I").sheet1

def compute_average_and_write(cell_range_str, target_cell, min_val=None, max_val=None):
    cell_range = worksheet.get(cell_range_str)
    valid_values = []

    for row in cell_range:
        if row:
            raw = row[0].strip()
            raw = raw.replace(',', '.')  # Handle comma decimals
            if raw and not raw.startswith('#'):
                try:
                    num = float(raw)
                    # Apply filtering for absurd values if thresholds are given
                    if (min_val is not None and num < min_val) or (max_val is not None and num > max_val):
                        continue
                    valid_values.append(num)
                except (ValueError, TypeError):
                    continue

    if valid_values:
        avg = round(sum(valid_values) / len(valid_values), 3)
        print(f"Average of valid {cell_range_str} cells = {avg}")
        worksheet.update(range_name=target_cell, values=[[avg]])
    else:
        print(f"No valid numeric values in {cell_range_str}.")

# Process C23:C27 → C28
compute_average_and_write('C23:C27', 'C28')

# Process H23:H27 → H28
compute_average_and_write('H23:H27', 'H28')

# Process B6:B8 → C9, exclude values < -0.75 or > 5.0
compute_average_and_write('B6:B8', 'C9', min_val=-0.75, max_val=5.0)

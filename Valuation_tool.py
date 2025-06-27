import gspread
import time
import tkinter as tk
from google.oauth2.service_account import Credentials

# === Authenticate ===
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file(
    'C:/Jon/skole/Master thesis/Python/XGBoost - Master thesis/PythonProject1/sheets/credentials.json',
    scopes=scopes
)
client = gspread.authorize(creds)
worksheet = client.open_by_key("1bc_ZZ_I2gtwL5kxKylKxIrYsbzk9XuQtCuexm_BDU4I").sheet1

# === Averaging function ===
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

# === GUI Window ===
def run_gui():
    def on_submit():
        ticker = ticker_entry.get().strip()
        if not ticker:
            status_label.config(text="Please enter a ticker", fg="red")
            return

        status_label.config(text=f"Loading data for '{ticker}'...", fg="black")
        root.update()

        # Write ticker
        worksheet.update(range_name='A2', values=[[ticker]])

        # Run calculations
        compute_average_and_write('C23:C27', 'C28')
        compute_average_and_write('H23:H27', 'H28')
        compute_average_and_write('B6:B8', 'C9', min_val=-0.75, max_val=5.0)

        time.sleep(2)  # let spreadsheet update

        try:
            dcf_price = worksheet.acell("F17").value
            avg_growth = worksheet.acell("C9").value
            ebitda_price = worksheet.acell("D27").value
            roic_price = worksheet.acell("J27").value

            # Format growth as percentage
            try:
                avg_growth_pct = f"{round(float(str(avg_growth).replace(',', '.')) * 100, 2)}%"
            except:
                avg_growth_pct = "N/A"

            # Clear previous content and show result
            for widget in output_frame.winfo_children():
                widget.destroy()

            tk.Label(output_frame, text=f"DCF & Multiple Price", font=("Helvetica", 14, "bold"), bg="#f5f5f5", fg="#003366").pack(pady=(10, 0))
            tk.Label(output_frame, text=f"{dcf_price}", font=("Helvetica", 24, "bold"), bg="#f5f5f5", fg="#007700").pack(pady=(0, 10))

            tk.Label(output_frame, text=f"Avg Growth: {avg_growth_pct}", font=("Helvetica", 11), bg="#f5f5f5").pack()
            tk.Label(output_frame, text=f"EBITDA Yield Total (D27): {ebitda_price}", font=("Helvetica", 11), bg="#f5f5f5").pack()
            tk.Label(output_frame, text=f"ROIC Total Price (J27): {roic_price}", font=("Helvetica", 11), bg="#f5f5f5").pack()

            status_label.config(text="✔ Data loaded successfully", fg="green")
        except Exception as e:
            status_label.config(text=f"Error fetching data: {e}", fg="red")

    # Create window
    root = tk.Tk()
    root.title("Stock Valuation Tool")
    root.geometry("420x400")
    root.configure(bg="#f5f5f5")

    tk.Label(root, text="Enter Stock Ticker:", font=("Helvetica", 12), bg="#f5f5f5").pack(pady=(20, 5))
    ticker_entry = tk.Entry(root, font=("Helvetica", 12), width=20, justify='center')
    ticker_entry.pack()

    tk.Button(root, text="Get Valuation", command=on_submit, font=("Helvetica", 11, "bold"), bg="#0077cc", fg="white").pack(pady=10)

    output_frame = tk.Frame(root, bg="#f5f5f5")
    output_frame.pack(pady=10)

    status_label = tk.Label(root, text="", font=("Helvetica", 10), bg="#f5f5f5")
    status_label.pack()

    root.mainloop()

# === Run the app ===
run_gui()

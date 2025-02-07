import os
import string
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyomo.environ as pyo

# --------------------------
# Utility functions to perform tier lookups
# --------------------------
def lookup_percentage(table_df, supplier, volume, col_name):
    """
    Given a DataFrame (for discount or rebate),
    supplier (string) and volume (numeric),
    look up the applicable percentage.
    The DataFrame is assumed to have columns:
       "Supplier Name", "Min Volume", "Max Volume", and the percentage column (e.g., "% Discount" or "% Rebate").
    Returns a float (as a decimal, e.g., 0.05 for 5%).
    """
    # Filter rows for the given supplier.
    df = table_df[table_df["Supplier Name"].str.strip() == supplier]
    if df.empty:
        return 0.0
    # Convert min and max volumes to numeric if necessary.
    df["Min Volume"] = pd.to_numeric(df["Min Volume"], errors="coerce")
    df["Max Volume"] = pd.to_numeric(df["Max Volume"], errors="coerce")
    # Sort by Min Volume.
    df = df.sort_values("Min Volume")
    for _, row in df.iterrows():
        if volume >= row["Min Volume"] and volume <= row["Max Volume"]:
            # Remove "%" if present and convert to decimal.
            perc = str(row[col_name]).strip()
            if "%" in perc:
                return float(perc.replace("%", "")) / 100.0
            else:
                return float(perc)
    return 0.0

def get_discount_and_rebate(supplier, volume, discount_df, rebate_df):
    """
    For a given supplier and volume, look up discount and rebate percentages.
    Returns a tuple: (discount, rebate) as decimals.
    """
    discount = lookup_percentage(discount_df, supplier, volume, "% Volume Discount")
    rebate = lookup_percentage(rebate_df, supplier, volume, "% Rebate")
    return discount, rebate

# --------------------------
# Global variables for file paths and DataFrames
# --------------------------
excel_file_path = None
sheet_names = []  # To be populated when an Excel file is loaded

bids_df = None
rebates_df = None
discount_df = None
results_df = None

bid_sheet_var = None
rebate_sheet_var = None
discount_sheet_var = None

bid_sheet_menu = None
rebate_sheet_menu = None
discount_sheet_menu = None

# --------------------------
# GUI Functions: Load Excel File and Set Sheet Names
# --------------------------
def load_excel_file():
    global excel_file_path, sheet_names, bid_sheet_menu, rebate_sheet_menu, discount_sheet_menu
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File with 3 Tabs",
        filetypes=[("Excel files", "*.xlsx *.xls")])
    if not excel_file_path:
        return
    try:
        xls = pd.ExcelFile(excel_file_path)
        sheet_names = xls.sheet_names
        messagebox.showinfo("File Loaded", f"Loaded file:\n{excel_file_path}\nAvailable sheets: {', '.join(sheet_names)}")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading Excel file:\n{e}")
        return
    if bid_sheet_menu is not None:
        bid_sheet_menu['values'] = sheet_names
    if rebate_sheet_menu is not None:
        rebate_sheet_menu['values'] = sheet_names
    if discount_sheet_menu is not None:
        discount_sheet_menu['values'] = sheet_names
    if sheet_names:
        bid_sheet_var.set(sheet_names[0])
        rebate_sheet_var.set(sheet_names[0])
        discount_sheet_var.set(sheet_names[0])

# --------------------------
# Main Model Function
# --------------------------
def run_model():
    global excel_file_path, bids_df, rebates_df, discount_df, results_df
    if not excel_file_path:
        messagebox.showerror("Input Error", "Please load an Excel file first.")
        return
    try:
        bids_df = pd.read_excel(excel_file_path, sheet_name=bid_sheet_var.get())
        rebates_df = pd.read_excel(excel_file_path, sheet_name=rebate_sheet_var.get())
        discount_df = pd.read_excel(excel_file_path, sheet_name=discount_sheet_var.get())
    except Exception as e:
        messagebox.showerror("Error", f"Error reading sheets:\n{e}")
        return

    # --------------------------
    # Prepare Bid Data
    # --------------------------
    try:
        bid_data = {}
        supplier_capacity = {}
        # Expected columns (case-sensitive): 
        # "Bid ID", "Bid Supplier Name", "Bid Volume", "Baseline Price",
        # "Current Price", "Incumbent", "Bid Supplier Capacity", "Bid Price", "Facility"
        for idx, row in bids_df.iterrows():
            orig_bid_id = str(row["Bid ID"]).strip()
            supplier = str(row["Bid Supplier Name"]).strip()
            unique_bid_id = orig_bid_id + "_" + supplier
            volume = float(row["Bid Volume"])
            bid_data[unique_bid_id] = {
                'orig_bid_id': orig_bid_id,
                'supplier': supplier,
                'facility': str(row["Facility"]).strip(),
                'volume': volume,
                'baseline_price': float(str(row["Baseline Price"]).replace("$", "").strip()),
                'current_price': float(str(row["Current Price"]).replace("$", "").strip()),
                'bid_price': float(str(row["Bid Price"]).replace("$", "").strip()),
                'capacity': float(row["Bid Supplier Capacity"]),
                'incumbent': str(row["Incumbent"]).strip()
            }
            # Assume capacity is the same for each supplier; store it.
            if supplier not in supplier_capacity:
                supplier_capacity[supplier] = float(row["Bid Supplier Capacity"])
            # Pre-compute discount and rebate for this bid row using the bid volume.
            discount, rebate = get_discount_and_rebate(supplier, volume, discount_df, rebates_df)
            bid_data[unique_bid_id]['discount'] = discount
            bid_data[unique_bid_id]['rebate'] = rebate
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing bid data: {e}")
        return

    orig_bid_ids = set(bid_data[b]['orig_bid_id'] for b in bid_data)

    # --------------------------
    # Build the Pyomo MILP Model
    # --------------------------
    model = pyo.ConcreteModel()

    model.BIDS = pyo.Set(initialize=list(bid_data.keys()))
    model.ORIG_BIDS = pyo.Set(initialize=sorted(orig_bid_ids))
    suppliers_list = sorted(set(bid_data[b]['supplier'] for b in bid_data))
    model.SUPPLIERS = pyo.Set(initialize=suppliers_list)

    def volume_init(model, b):
        return bid_data[b]['volume']
    model.volume = pyo.Param(model.BIDS, initialize=volume_init)

    def baseline_price_init(model, b):
        return bid_data[b]['baseline_price']
    model.baseline_price = pyo.Param(model.BIDS, initialize=baseline_price_init)

    def bid_price_init(model, b):
        return bid_data[b]['bid_price']
    model.bid_price = pyo.Param(model.BIDS, initialize=bid_price_init)

    def facility_of_bid(model, b):
        return bid_data[b]['facility']
    model.facility_of_bid = pyo.Param(model.BIDS, initialize=facility_of_bid, within=pyo.Any)

    def supplier_of_bid(model, b):
        return bid_data[b]['supplier']
    model.supplier_of_bid = pyo.Param(model.BIDS, initialize=supplier_of_bid, within=pyo.Any)

    # Decision variable: y[b] is binary: 1 if bid row b is awarded, 0 otherwise.
    model.y = pyo.Var(model.BIDS, domain=pyo.Binary)

    # For each original Bid ID, exactly one bid row is awarded.
    def one_bid_per_orig_bid_rule(model, orig_bid):
        return sum(model.y[b] for b in model.BIDS if bid_data[b]['orig_bid_id'] == orig_bid) == 1
    model.one_bid_per_orig_bid = pyo.Constraint(model.ORIG_BIDS, rule=one_bid_per_orig_bid_rule)

    # For each supplier, total awarded volume cannot exceed capacity.
    def supplier_capacity_rule(model, s):
        return sum(bid_data[b]['volume'] * model.y[b] for b in model.BIDS if bid_data[b]['supplier'] == s) <= supplier_capacity[s]
    model.supplier_capacity_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_capacity_rule)

    # Now, define the saving for each bid row.
    # For bid row b (with supplier s), let:
    #   discount = d_b, rebate = r_b (both parameters computed earlier),
    #   then Discounted Awarded Supplier Price = bid_price[b]*(1 - d_b),
    #   Baseline Savings = volume[b]*(baseline_price[b] - bid_price[b]*(1 - d_b)),
    #   Rebate Savings = volume[b]*(bid_price[b]*(1 - d_b))*r_b.
    # Total Saving for bid row b = volume[b]*(baseline_price[b] - bid_price[b]*(1 - d_b)*(1 - r_b)).
    savings = {b: bid_data[b]['volume'] * (bid_data[b]['baseline_price'] - bid_data[b]['bid_price'] * (1 - bid_data[b]['discount']) * (1 - bid_data[b]['rebate']))
               for b in bid_data}
    
    # Objective: maximize total savings over all awarded bid rows.
    def obj_rule(model):
        return sum(savings[b] * model.y[b] for b in model.BIDS)
    model.obj = pyo.Objective(rule=obj_rule, sense=pyo.maximize)

    # --------------------------
    # Solve the Model using GLPK
    # --------------------------
    try:
        solver = pyo.SolverFactory('glpk')
        result = solver.solve(model, tee=True)
    except Exception as e:
        messagebox.showerror("Solver Error", f"Error during model solve: {e}")
        return

    # --------------------------
    # Prepare Output Data
    # --------------------------
    output_rows = []
    for b in model.BIDS:
        if pyo.value(model.y[b]) > 0.5:
            info = bid_data[b]
            awarded_volume = info['volume']  # all-or-nothing awarding
            supplier = info['supplier']
            d_val = info['discount']
            r_val = info['rebate']
            orig_price = info['bid_price']
            discounted_price = orig_price * (1 - d_val)
            baseline_savings = awarded_volume * (info['baseline_price'] - orig_price*(1 - d_val))
            rebate_savings = awarded_volume * (orig_price*(1 - d_val)) * r_val
            output_rows.append({
                "Bid ID": info['orig_bid_id'],
                # "Bid ID Split" will be assigned later (here, since it's all-or-nothing, it will be "A")
                "Facility": info['facility'],
                "Incumbent": info['incumbent'],
                "Baseline Price": info['baseline_price'],
                "Awarded Supplier": supplier,
                "Original Awarded Supplier Price": orig_price,
                "Percentage Volume Discount": f"{d_val*100:.0f}%",
                "Discounted Awarded Supplier Price": discounted_price,
                "Awarded Volume": awarded_volume,
                "Awarded Supplier Capacity": info['capacity'],
                "Baseline Savings": baseline_savings,
                "Rebate Savings": rebate_savings
            })
    if not output_rows:
        messagebox.showwarning("No Awarded Bids", "No bids were awarded by the model.")
        return

    # Since each bid is all-or-nothing, every original Bid ID will have exactly one row.
    for row in output_rows:
        row["Bid ID Split"] = "A"

    final_columns = ["Bid ID", "Bid ID Split", "Facility", "Incumbent", "Baseline Price",
                     "Awarded Supplier", "Original Awarded Supplier Price",
                     "Percentage Volume Discount", "Discounted Awarded Supplier Price",
                     "Awarded Volume", "Awarded Supplier Capacity", "Baseline Savings", "Rebate Savings"]
    results_df = pd.DataFrame(output_rows)[final_columns]

    # --------------------------
    # Save Output Automatically to Downloads Folder
    # --------------------------
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    output_filepath = os.path.join(downloads_folder, "model_output.xlsx")
    try:
        results_df.to_excel(output_filepath, index=False)
        messagebox.showinfo("Success", f"Model run complete.\nOutput saved to:\n{output_filepath}")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving output:\n{e}")

def save_output():
    pass

# --------------------------
# Build the Tkinter GUI
# --------------------------
root = tk.Tk()
root.title("Bid Optimization Model (All-or-Nothing, GLPK)")

bid_sheet_var = tk.StringVar(root)
rebate_sheet_var = tk.StringVar(root)
discount_sheet_var = tk.StringVar(root)

top_frame = tk.Frame(root)
top_frame.pack(pady=10)

load_file_btn = tk.Button(top_frame, text="Load Excel File (3 Tabs)", width=30, command=load_excel_file)
load_file_btn.grid(row=0, column=0, columnspan=3, padx=5, pady=5)

tk.Label(top_frame, text="Bid Data Sheet:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
bid_sheet_menu = ttk.Combobox(top_frame, textvariable=bid_sheet_var, values=sheet_names, state="readonly", width=20)
bid_sheet_menu.grid(row=1, column=1, padx=5, pady=5)

tk.Label(top_frame, text="Rebate Data Sheet:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
rebate_sheet_menu = ttk.Combobox(top_frame, textvariable=rebate_sheet_var, values=sheet_names, state="readonly", width=20)
rebate_sheet_menu.grid(row=2, column=1, padx=5, pady=5)

tk.Label(top_frame, text="Volume Discount Data Sheet:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
discount_sheet_menu = ttk.Combobox(top_frame, textvariable=discount_sheet_var, values=sheet_names, state="readonly", width=20)
discount_sheet_menu.grid(row=3, column=1, padx=5, pady=5)

run_model_btn = tk.Button(root, text="Run Model", width=30, command=run_model)
run_model_btn.pack(pady=10)

root.mainloop()

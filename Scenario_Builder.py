import os
import string
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyomo.environ as pyo

# --------------------------
# Helper functions for tier lookups
# --------------------------
def lookup_percentage(df, supplier, volume, col_name):
    """
    Given a DataFrame (for discount or rebate) with columns:
    "Supplier Name", "Min Volume", "Max Volume", and the percentage column,
    return the applicable percentage (as a decimal) for the given supplier and volume.
    """
    df_sup = df[df["Supplier Name"].str.strip() == supplier]
    if df_sup.empty:
        return 0.0
    # Ensure numeric conversion
    df_sup["Min Volume"] = pd.to_numeric(df_sup["Min Volume"], errors="coerce")
    df_sup["Max Volume"] = pd.to_numeric(df_sup["Max Volume"], errors="coerce")
    df_sup = df_sup.sort_values("Min Volume")
    for _, row in df_sup.iterrows():
        if volume >= row["Min Volume"] and volume <= row["Max Volume"]:
            perc = str(row[col_name]).strip()
            if "%" in perc:
                return float(perc.replace("%", "")) / 100.0
            else:
                return float(perc)
    return 0.0

def get_discount_and_rebate(supplier, volume, discount_df, rebate_df):
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

# Global OptionMenu variables for sheet selection
bid_sheet_var = None
rebate_sheet_var = None
discount_sheet_var = None

bid_sheet_menu = None
rebate_sheet_menu = None
discount_sheet_menu = None

# --------------------------
# Discrete awarding options
# --------------------------
OPTIONS = [1, 2, 3, 4]  # Option indices
option_percent = {1: 0.25, 2: 0.50, 3: 0.75, 4: 1.00}

# --------------------------
# GUI: Load Excel File
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
# Main Model Function (MILP: All-or-Nothing, Dual-Sourcing, Discrete Splits)
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
    # Prepare Bid Data and Precompute Savings for each option
    # --------------------------
    try:
        bid_data = {}
        supplier_capacity = {}
        S_full = {}  # key: (bid_row, option) → computed savings if that bid row is awarded with that fraction
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
            if supplier not in supplier_capacity:
                supplier_capacity[supplier] = float(row["Bid Supplier Capacity"])
            bid_data[unique_bid_id]['option'] = {}
            for k in OPTIONS:
                p = option_percent[k]
                awarded_volume = volume * p
                discount, rebate = get_discount_and_rebate(supplier, awarded_volume, discount_df, rebates_df)
                bid_data[unique_bid_id]['option'][k] = {'discount': discount, 'rebate': rebate}
                # Compute savings for awarding the bid row with fraction p.
                S_full[(unique_bid_id, k)] = awarded_volume * (bid_data[unique_bid_id]['baseline_price'] - 
                                            bid_data[unique_bid_id]['bid_price']*(1 - discount)*(1 - rebate))
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing bid data: {e}")
        return

    orig_bid_ids = set(bid_data[b]['orig_bid_id'] for b in bid_data)

    # --------------------------
    # Build the MILP Model
    # --------------------------
    model = pyo.ConcreteModel()
    model.BIDS = pyo.Set(initialize=list(bid_data.keys()))
    model.ORIG_BIDS = pyo.Set(initialize=sorted(orig_bid_ids))
    model.SUPPLIERS = pyo.Set(initialize=sorted(set(bid_data[b]['supplier'] for b in bid_data)))
    model.OPTIONS = pyo.Set(initialize=OPTIONS)

    # Decision variables: x[b,k] ∈ {0,1} indicating if bid row b is awarded with option k.
    model.x = pyo.Var(model.BIDS, model.OPTIONS, domain=pyo.Binary)

    # Constraint: For each bid row, at most one awarding option is chosen.
    def one_option_per_bid_rule(model, b):
        return sum(model.x[b, k] for k in model.OPTIONS) <= 1
    model.one_option_per_bid = pyo.Constraint(model.BIDS, rule=one_option_per_bid_rule)

    # Constraint: For each original Bid ID, the total awarded fraction equals 1.
    def full_allocation_rule(model, orig_bid):
        return sum(option_percent[k] * model.x[b, k]
                   for b in model.BIDS if bid_data[b]['orig_bid_id'] == orig_bid
                   for k in model.OPTIONS) == 1
    model.full_allocation = pyo.Constraint(model.ORIG_BIDS, rule=full_allocation_rule)

    # Constraint: For each original Bid ID, at most 2 bid rows are awarded.
    def dual_source_rule(model, orig_bid):
        return sum(model.x[b, k]
                   for b in model.BIDS if bid_data[b]['orig_bid_id'] == orig_bid
                   for k in model.OPTIONS) <= 2
    model.dual_source = pyo.Constraint(model.ORIG_BIDS, rule=dual_source_rule)

    # Supplier capacity: For each supplier, total awarded volume ≤ capacity.
    def supplier_capacity_rule(model, s):
        return sum(bid_data[b]['volume'] * option_percent[k] * model.x[b, k]
                   for b in model.BIDS if bid_data[b]['supplier'] == s
                   for k in model.OPTIONS) <= supplier_capacity[s]
    model.supplier_capacity = pyo.Constraint(model.SUPPLIERS, rule=supplier_capacity_rule)

    # Objective: maximize total savings.
    def obj_rule(model):
        return sum(S_full[(b, k)] * model.x[b, k] for b in model.BIDS for k in model.OPTIONS)
    model.obj = pyo.Objective(rule=obj_rule, sense=pyo.maximize)

    # --------------------------
    # Solve the MILP using GLPK
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
        for k in model.OPTIONS:
            if pyo.value(model.x[b, k]) > 0.5:
                info = bid_data[b]
                fraction = option_percent[k]
                awarded_volume = info['volume'] * fraction
                supplier = info['supplier']
                d_val = info['option'][k]['discount']
                r_val = info['option'][k]['rebate']
                orig_price = info['bid_price']
                discounted_price = orig_price * (1 - d_val)
                baseline_savings = awarded_volume * (info['baseline_price'] - orig_price*(1 - d_val))
                rebate_savings = awarded_volume * (orig_price*(1 - d_val)) * r_val
                output_rows.append({
                    "Bid ID": info['orig_bid_id'],
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

    # Assign Bid ID Split letters. Since each original Bid ID can have 1 or 2 awarded rows,
    # we label them as A and B.
    grouped = {}
    for row in output_rows:
        bid_id = row["Bid ID"]
        if bid_id not in grouped:
            grouped[bid_id] = []
        grouped[bid_id].append(row)
    final_rows = []
    for bid_id in sorted(grouped.keys()):
        group = grouped[bid_id]
        group_sorted = sorted(group, key=lambda r: r["Awarded Supplier"])
        for i, row in enumerate(group_sorted):
            row["Bid ID Split"] = string.ascii_uppercase[i] if i < len(string.ascii_uppercase) else "Z"
            final_rows.append(row)
    
    final_columns = ["Bid ID", "Bid ID Split", "Facility", "Incumbent", "Baseline Price",
                     "Awarded Supplier", "Original Awarded Supplier Price",
                     "Percentage Volume Discount", "Discounted Awarded Supplier Price",
                     "Awarded Volume", "Awarded Supplier Capacity", "Baseline Savings", "Rebate Savings"]
    results_df = pd.DataFrame(final_rows)[final_columns]

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
root.title("Bid Optimization Model (Dual-Sourcing MILP, GLPK)")

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

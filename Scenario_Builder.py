import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyomo.environ as pyo

# Global variables for file path and data
excel_file_path = None
sheet_names = []       # List to store sheet names from the loaded Excel file

# Global variables for DataFrames (will be set after reading the selected sheets)
bids_df = None
rebates_df = None
discount_df = None
results_df = None

# Global OptionMenu variables for sheet selection
bid_sheet_var = None
rebate_sheet_var = None
discount_sheet_var = None

# Global combobox widget variables (to update values later)
bid_sheet_menu = None
rebate_sheet_menu = None
discount_sheet_menu = None

def load_excel_file():
    global excel_file_path, sheet_names, bid_sheet_menu, rebate_sheet_menu, discount_sheet_menu
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File with 3 Tabs",
        filetypes=[("Excel files", "*.xlsx *.xls")])
    if not excel_file_path:
        return

    try:
        # Load the Excel file to get the sheet names
        xls = pd.ExcelFile(excel_file_path)
        sheet_names = xls.sheet_names
        messagebox.showinfo("File Loaded", f"Loaded file:\n{excel_file_path}\nAvailable sheets: {', '.join(sheet_names)}")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading Excel file:\n{e}")
        return

    # Update the dropdown menus with the retrieved sheet names
    if bid_sheet_menu is not None:
        bid_sheet_menu['values'] = sheet_names
    if rebate_sheet_menu is not None:
        rebate_sheet_menu['values'] = sheet_names
    if discount_sheet_menu is not None:
        discount_sheet_menu['values'] = sheet_names

    # Set the default selection for each dropdown to the first sheet (if available)
    if sheet_names:
        bid_sheet_var.set(sheet_names[0])
        rebate_sheet_var.set(sheet_names[0])
        discount_sheet_var.set(sheet_names[0])

def run_model():
    global excel_file_path, bids_df, rebates_df, discount_df, results_df
    if not excel_file_path:
        messagebox.showerror("Input Error", "Please load an Excel file first.")
        return

    try:
        # Read the selected sheets into DataFrames
        bids_df = pd.read_excel(excel_file_path, sheet_name=bid_sheet_var.get())
        rebates_df = pd.read_excel(excel_file_path, sheet_name=rebate_sheet_var.get())
        discount_df = pd.read_excel(excel_file_path, sheet_name=discount_sheet_var.get())
    except Exception as e:
        messagebox.showerror("Error", f"Error reading sheets from Excel file:\n{e}")
        return

    # --------------------------
    # Prepare Bid Data
    # --------------------------
    try:
        bid_data = {}
        supplier_capacity = {}
        # Expected columns in bids_df:
        # "Bid ID", "Bid Supplier Name", "Bid Volume", "Baseline Price",
        # "Current Price", "Incumbent", "Bid Supplier Capacity", "Bid Price", "Facility"
        for idx, row in bids_df.iterrows():
            orig_bid_id = str(row["Bid ID"]).strip()
            supplier = str(row["Bid Supplier Name"]).strip()
            # Create a unique key by combining Bid ID and Bid Supplier Name
            unique_bid_id = orig_bid_id + "_" + supplier
            bid_data[unique_bid_id] = {
                'orig_bid_id': orig_bid_id,  # For output, we preserve the original Bid ID
                'supplier': supplier,
                'facility': str(row["Facility"]).strip(),
                'volume': float(row["Bid Volume"]),
                'baseline_price': float(str(row["Baseline Price"]).replace("$", "").strip()),
                'current_price': float(str(row["Current Price"]).replace("$", "").strip()),
                'bid_price': float(str(row["Bid Price"]).replace("$", "").strip()),
                'capacity': float(row["Bid Supplier Capacity"]),
                'incumbent': str(row["Incumbent"]).strip()
            }
            if supplier not in supplier_capacity:
                supplier_capacity[supplier] = float(row["Bid Supplier Capacity"])
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing bid data: {e}")
        return

    # --------------------------
    # Prepare Rebate Data
    # --------------------------
    rebate_data = []
    try:
        # Expected columns in rebates_df: "Supplier Name", "Min Volume", "Max Volume", "% Rebate"
        for idx, row in rebates_df.iterrows():
            supplier = str(row["Supplier Name"]).strip()
            min_vol = float(row["Min Volume"])
            max_vol = float(row["Max Volume"])
            r_str = str(row["% Rebate"]).strip()
            if "%" in r_str:
                rebate_val = float(r_str.replace("%", "")) / 100.0
            else:
                rebate_val = float(r_str)
            rebate_data.append({
                'supplier': supplier,
                'min_volume': min_vol,
                'max_volume': max_vol,
                'rebate': rebate_val
            })
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing rebate data: {e}")
        return

    # --------------------------
    # Prepare Volume Discount Data
    # --------------------------
    discount_data = []
    try:
        # Expected columns in discount_df: "Supplier Name", "Min Volume", "Max Volume", "% Volume Discount"
        for idx, row in discount_df.iterrows():
            supplier = str(row["Supplier Name"]).strip()
            min_vol = float(row["Min Volume"])
            max_vol = float(row["Max Volume"])
            d_str = str(row["% Volume Discount"]).strip()
            if "%" in d_str:
                discount_val = float(d_str.replace("%", "")) / 100.0
            else:
                discount_val = float(d_str)
            discount_data.append({
                'supplier': supplier,
                'min_volume': min_vol,
                'max_volume': max_vol,
                'discount': discount_val
            })
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing discount data: {e}")
        return

    # --------------------------
    # Build the Pyomo Model
    # --------------------------
    model = pyo.ConcreteModel()

    # Sets: Unique bids, Facilities, and Suppliers
    model.BIDS = pyo.Set(initialize=list(bid_data.keys()))
    facilities = set(bid['facility'] for bid in bid_data.values())
    model.ITEMS = pyo.Set(initialize=facilities)
    suppliers = set(bid['supplier'] for bid in bid_data.values())
    model.SUPPLIERS = pyo.Set(initialize=suppliers)

    # Parameters from bid_data
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
    model.facility_of_bid = pyo.Param(model.BIDS, initialize=facility_of_bid)

    def supplier_of_bid(model, b):
        return bid_data[b]['supplier']
    model.supplier_of_bid = pyo.Param(model.BIDS, initialize=supplier_of_bid)

    # Decision variable: x[b] = 1 if bid b is awarded.
    model.x = pyo.Var(model.BIDS, domain=pyo.Binary)

    # Awarded Volume per Supplier: V[s]
    model.V = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def supplier_volume_rule(model, s):
        return model.V[s] == sum(model.volume[b] * model.x[b] for b in model.BIDS if bid_data[b]['supplier'] == s)
    model.supplier_volume_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_volume_rule)

    # Supplier Bid Cost: C[s]
    model.C = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def supplier_cost_rule(model, s):
        return model.C[s] == sum(model.bid_price[b] * model.volume[b] * model.x[b]
                                   for b in model.BIDS if bid_data[b]['supplier'] == s)
    model.supplier_cost_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_cost_rule)

    # --------------------------
    # Rebate Tiers (Kick-back)
    # --------------------------
    tiers_by_supplier = {}
    for s in suppliers:
        tiers = [ (row['min_volume'], row['max_volume'], row['rebate'])
                  for row in rebate_data if row['supplier'] == s ]
        if tiers:
            tiers.sort(key=lambda x: x[0])
            if tiers[0][0] > 0:
                tiers.insert(0, (0, tiers[0][0]-1, 0.0))
        else:
            tiers = [(0, supplier_capacity[s], 0.0)]
        tiers_by_supplier[s] = tiers

    tier_index_set = []
    for s in suppliers:
        for r in range(len(tiers_by_supplier[s])):
            tier_index_set.append((s, r))
    model.TIERS = pyo.Set(initialize=tier_index_set, dimen=2)

    model.y = pyo.Var(model.TIERS, domain=pyo.Binary)
    model.z = pyo.Var(model.TIERS, domain=pyo.NonNegativeReals)
    def one_tier_per_supplier_rule(model, s):
        return sum(model.y[s, r] for r in range(len(tiers_by_supplier[s]))) == 1
    model.one_tier_per_supplier = pyo.Constraint(model.SUPPLIERS, rule=one_tier_per_supplier_rule)
    M_vol = { s: supplier_capacity[s] for s in suppliers }
    def tier_volume_lower_rule(model, s, r):
        return model.V[s] >= tiers_by_supplier[s][r][0] * model.y[s, r]
    model.tier_volume_lower = pyo.Constraint(model.TIERS, rule=tier_volume_lower_rule)
    def tier_volume_upper_rule(model, s, r):
        return model.V[s] <= tiers_by_supplier[s][r][1] + M_vol[s]*(1 - model.y[s, r])
    model.tier_volume_upper = pyo.Constraint(model.TIERS, rule=tier_volume_upper_rule)
    def z_upper1_rule(model, s, r):
        return model.z[s, r] <= model.V[s]
    model.z_upper1 = pyo.Constraint(model.TIERS, rule=z_upper1_rule)
    def z_upper2_rule(model, s, r):
        return model.z[s, r] <= M_vol[s] * model.y[s, r]
    model.z_upper2 = pyo.Constraint(model.TIERS, rule=z_upper2_rule)
    def z_lower_rule(model, s, r):
        return model.z[s, r] >= model.V[s] - M_vol[s]*(1 - model.y[s, r])
    model.z_lower = pyo.Constraint(model.TIERS, rule=z_lower_rule)
    model.rebate_saving = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def rebate_saving_rule(model, s):
        return model.rebate_saving[s] == sum(tiers_by_supplier[s][r][2] * model.z[s, r]
                                              for r in range(len(tiers_by_supplier[s])))
    model.rebate_saving_con = pyo.Constraint(model.SUPPLIERS, rule=rebate_saving_rule)

    # --------------------------
    # Discount Tiers (Volume Discount)
    # --------------------------
    discount_tiers_by_supplier = {}
    for s in suppliers:
        dt = [ (row['min_volume'], row['max_volume'], row['discount'])
               for row in discount_data if row['supplier'] == s ]
        if dt:
            dt.sort(key=lambda x: x[0])
            if dt[0][0] > 0:
                dt.insert(0, (0, dt[0][0]-1, 0.0))
        else:
            dt = [(0, supplier_capacity[s], 0.0)]
        discount_tiers_by_supplier[s] = dt

    discount_index_set = []
    for s in suppliers:
        for r in range(len(discount_tiers_by_supplier[s])):
            discount_index_set.append((s, r))
    model.DTIERS = pyo.Set(initialize=discount_index_set, dimen=2)

    model.d_y = pyo.Var(model.DTIERS, domain=pyo.Binary)
    model.d_z = pyo.Var(model.DTIERS, domain=pyo.NonNegativeReals)
    def one_discount_tier_rule(model, s):
        return sum(model.d_y[s, r] for r in range(len(discount_tiers_by_supplier[s]))) == 1
    model.one_discount_tier = pyo.Constraint(model.SUPPLIERS, rule=one_discount_tier_rule)
    M_cost = { s: 1e8 for s in suppliers }
    def discount_z_upper1_rule(model, s, r):
        return model.d_z[s, r] <= model.C[s]
    model.discount_z_upper1 = pyo.Constraint(model.DTIERS, rule=discount_z_upper1_rule)
    def discount_z_upper2_rule(model, s, r):
        return model.d_z[s, r] <= M_cost[s] * model.d_y[s, r]
    model.discount_z_upper2 = pyo.Constraint(model.DTIERS, rule=discount_z_upper2_rule)
    def discount_z_lower_rule(model, s, r):
        return model.d_z[s, r] >= model.C[s] - M_cost[s]*(1 - model.d_y[s, r])
    model.discount_z_lower = pyo.Constraint(model.DTIERS, rule=discount_z_lower_rule)
    model.discount_saving = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def discount_saving_rule(model, s):
        return model.discount_saving[s] == sum(discount_tiers_by_supplier[s][r][2] * model.d_z[s, r]
                                                for r in range(len(discount_tiers_by_supplier[s])))
    model.discount_saving_con = pyo.Constraint(model.SUPPLIERS, rule=discount_saving_rule)

    # --------------------------
    # Objective: Maximize Overall Savings
    # --------------------------
    def obj_rule(model):
        baseline_savings = sum((model.baseline_price[b] - model.bid_price[b]) * model.volume[b] * model.x[b]
                               for b in model.BIDS)
        rebate_bonus = sum(model.rebate_saving[s] for s in model.SUPPLIERS)
        discount_bonus = sum(model.discount_saving[s] for s in model.SUPPLIERS)
        return baseline_savings + rebate_bonus + discount_bonus
    model.obj = pyo.Objective(rule=obj_rule, sense=pyo.maximize)

    # --------------------------
    # Constraint: Exactly One Bid Awarded per Facility
    # --------------------------
    def one_bid_per_facility_rule(model, i):
        return sum(model.x[b] for b in model.BIDS if bid_data[b]['facility'] == i) == 1
    model.one_bid_per_facility = pyo.Constraint(model.ITEMS, rule=one_bid_per_facility_rule)

    # --------------------------
    # Solve the Model using GLPK
    # --------------------------
    try:
        solver = pyo.SolverFactory('glpk')
        result = solver.solve(model, tee=False)
    except Exception as e:
        messagebox.showerror("Solver Error", f"Error during model solve: {e}")
        return

    # --------------------------
    # Prepare Output Data
    # --------------------------
    output_rows = []
    for b in model.BIDS:
        if pyo.value(model.x[b]) > 0.5:
            info = bid_data[b]
            baseline_savings = (info['baseline_price'] - info['bid_price']) * info['volume']
            supplier = info['supplier']
            rebate_save = pyo.value(model.rebate_saving[supplier])
            discount_save = pyo.value(model.discount_saving[supplier])
            # Determine the discount percentage from the discount tier selection
            disc_pct = 0.0
            for r in range(len(discount_tiers_by_supplier[supplier])):
                if pyo.value(model.d_y[supplier, r]) > 0.5:
                    disc_pct = discount_tiers_by_supplier[supplier][r][2]
                    break
            orig_price = info['bid_price']
            discounted_price = orig_price * (1 - disc_pct)
            # Build the output row with the new columns in place of "Awarded Supplier Price"
            output_rows.append({
                "Bid ID": info['orig_bid_id'],
                "Facility": info['facility'],
                "Incumbent": info['incumbent'],
                "Baseline Price": info['baseline_price'],
                "Awarded Supplier": supplier,
                "Awarded Volume": info['volume'],
                "Awarded Supplier Capacity": info['capacity'],
                "Baseline Savings": baseline_savings,
                "Rebate Savings": rebate_save,
                "Original Awarded Supplier Price": orig_price,
                "Percentage Volume Discount": disc_pct,
                "Discounted Awarded Supplier Price": discounted_price
            })
    if not output_rows:
        messagebox.showwarning("No Awarded Bids", "No bids were awarded by the model.")
        return

    results_df = pd.DataFrame(output_rows)

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
    # This function is no longer needed because output is saved automatically.
    pass

# --------------------------
# Build the Tkinter GUI
# --------------------------
root = tk.Tk()
root.title("Bid Optimization Model (Excel I/O with Tabs & Discounts)")

# Initialize global OptionMenu (Combobox) variables
bid_sheet_var = tk.StringVar(root)
rebate_sheet_var = tk.StringVar(root)
discount_sheet_var = tk.StringVar(root)

# Frame for file and sheet selection
top_frame = tk.Frame(root)
top_frame.pack(pady=10)

# Button to load the Excel file (which contains 3 tabs)
load_file_btn = tk.Button(top_frame, text="Load Excel File (3 Tabs)", width=30, command=load_excel_file)
load_file_btn.grid(row=0, column=0, columnspan=3, padx=5, pady=5)

# Dropdown menus for sheet selection (initially empty; will be updated once a file is loaded)
tk.Label(top_frame, text="Bid Data Sheet:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
bid_sheet_menu = ttk.Combobox(top_frame, textvariable=bid_sheet_var, values=sheet_names, state="readonly", width=20)
bid_sheet_menu.grid(row=1, column=1, padx=5, pady=5)

tk.Label(top_frame, text="Rebate Data Sheet:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
rebate_sheet_menu = ttk.Combobox(top_frame, textvariable=rebate_sheet_var, values=sheet_names, state="readonly", width=20)
rebate_sheet_menu.grid(row=2, column=1, padx=5, pady=5)

tk.Label(top_frame, text="Volume Discount Data Sheet:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
discount_sheet_menu = ttk.Combobox(top_frame, textvariable=discount_sheet_var, values=sheet_names, state="readonly", width=20)
discount_sheet_menu.grid(row=3, column=1, padx=5, pady=5)

# Button to run the model (which will automatically save output to the Downloads folder)
run_model_btn = tk.Button(root, text="Run Model", width=30, command=run_model)
run_model_btn.pack(pady=10)

root.mainloop()

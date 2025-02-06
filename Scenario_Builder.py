import os
import string
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyomo.environ as pyo

# Global variables for file path and sheet names
excel_file_path = None
sheet_names = []  # To be populated when an Excel file is loaded

# Global DataFrames (set after reading the selected sheets)
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
        xls = pd.ExcelFile(excel_file_path)
        sheet_names = xls.sheet_names
        messagebox.showinfo("File Loaded", f"Loaded file:\n{excel_file_path}\nAvailable sheets: {', '.join(sheet_names)}")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading Excel file:\n{e}")
        return
    # Update dropdown menus with sheet names
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
        messagebox.showerror("Error", f"Error reading sheets from Excel file:\n{e}")
        return

    # --------------------------
    # Prepare Bid Data
    # --------------------------
    try:
        bid_data = {}
        supplier_capacity = {}
        # Expected columns: "Bid ID", "Bid Supplier Name", "Bid Volume", "Baseline Price",
        # "Current Price", "Incumbent", "Bid Supplier Capacity", "Bid Price", "Facility"
        for idx, row in bids_df.iterrows():
            orig_bid_id = str(row["Bid ID"]).strip()
            supplier = str(row["Bid Supplier Name"]).strip()
            unique_bid_id = orig_bid_id + "_" + supplier  # Unique key per row
            bid_data[unique_bid_id] = {
                'orig_bid_id': orig_bid_id,  # preserve original Bid ID for output
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

    # Create set of unique original Bid IDs
    orig_bid_ids = set(bid_data[b]['orig_bid_id'] for b in bid_data)

    # --------------------------
    # Prepare Rebate Data
    # --------------------------
    rebate_data = []
    try:
        # Expected columns: "Supplier Name", "Min Volume", "Max Volume", "% Rebate"
        for idx, row in rebates_df.iterrows():
            supplier = str(row["Supplier Name"]).strip()
            min_vol = float(row["Min Volume"])
            max_vol = float(row["Max Volume"])
            r_str = str(row["% Rebate"]).strip()
            rebate_val = float(r_str.replace("%", "")) / 100.0 if "%" in r_str else float(r_str)
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
        # Expected columns: "Supplier Name", "Min Volume", "Max Volume", "% Volume Discount"
        for idx, row in discount_df.iterrows():
            supplier = str(row["Supplier Name"]).strip()
            min_vol = float(row["Min Volume"])
            max_vol = float(row["Max Volume"])
            d_str = str(row["% Volume Discount"]).strip()
            discount_val = float(d_str.replace("%", "")) / 100.0 if "%" in d_str else float(d_str)
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

    # Define sets. Use sorted lists for deterministic ordering.
    model.BIDS = pyo.Set(initialize=list(bid_data.keys()))
    model.ORIG_BIDS = pyo.Set(initialize=sorted(orig_bid_ids))
    suppliers_list = sorted(set(bid_data[b]['supplier'] for b in bid_data))
    model.SUPPLIERS = pyo.Set(initialize=suppliers_list)

    # Parameters
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

    # Decision variable: x[b] is the fraction of bid b's volume awarded (allowing splits).
    model.x = pyo.Var(model.BIDS, domain=pyo.UnitInterval)

    # Constraint: For each original Bid ID, the sum of x[b] (across rows with that Bid ID) equals 1.
    def one_bid_per_orig_bid_rule(model, orig_bid):
        return sum(model.x[b] for b in model.BIDS if bid_data[b]['orig_bid_id'] == orig_bid) == 1
    model.one_bid_per_orig_bid = pyo.Constraint(model.ORIG_BIDS, rule=one_bid_per_orig_bid_rule)

    # Awarded Volume per Supplier: V[s] = sum(volume[b]*x[b]) for bids from supplier s.
    model.V = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def supplier_volume_rule(model, s):
        return model.V[s] == sum(model.volume[b]*model.x[b] for b in model.BIDS if bid_data[b]['supplier'] == s)
    model.supplier_volume_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_volume_rule)

    # Supplier Capacity Constraint: V[s] <= capacity for each supplier.
    def supplier_capacity_rule(model, s):
        return model.V[s] <= supplier_capacity[s]
    model.supplier_capacity_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_capacity_rule)

    # Supplier Bid Cost: C[s] = sum(bid_price[b]*volume[b]*x[b]) for bids from supplier s.
    model.C = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def supplier_cost_rule(model, s):
        return model.C[s] == sum(model.bid_price[b]*model.volume[b]*model.x[b]
                                   for b in model.BIDS if bid_data[b]['supplier'] == s)
    model.supplier_cost_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_cost_rule)

    # --------------------------
    # Rebate Tiers (Kick-back)
    # --------------------------
    tiers_by_supplier = {}
    for s in model.SUPPLIERS.value:
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
    for s in model.SUPPLIERS.value:
        for r in range(len(tiers_by_supplier[s])):
            tier_index_set.append((s, r))
    model.TIERS = pyo.Set(initialize=tier_index_set, dimen=2)

    model.y = pyo.Var(model.TIERS, domain=pyo.Binary)
    model.z = pyo.Var(model.TIERS, domain=pyo.NonNegativeReals)
    def one_tier_per_supplier_rule(model, s):
        return sum(model.y[s, r] for r in range(len(tiers_by_supplier[s]))) == 1
    model.one_tier_per_supplier = pyo.Constraint(model.SUPPLIERS, rule=one_tier_per_supplier_rule)
    M_vol = { s: supplier_capacity[s] for s in model.SUPPLIERS.value }
    def tier_volume_lower_rule(model, s, r):
        return model.V[s] >= tiers_by_supplier[s][r][0]*model.y[s, r]
    model.tier_volume_lower = pyo.Constraint(model.TIERS, rule=tier_volume_lower_rule)
    def tier_volume_upper_rule(model, s, r):
        return model.V[s] <= tiers_by_supplier[s][r][1] + M_vol[s]*(1 - model.y[s, r])
    model.tier_volume_upper = pyo.Constraint(model.TIERS, rule=tier_volume_upper_rule)
    def z_upper1_rule(model, s, r):
        return model.z[s, r] <= model.V[s]
    model.z_upper1 = pyo.Constraint(model.TIERS, rule=z_upper1_rule)
    def z_upper2_rule(model, s, r):
        return model.z[s, r] <= M_vol[s]*model.y[s, r]
    model.z_upper2 = pyo.Constraint(model.TIERS, rule=z_upper2_rule)
    def z_lower_rule(model, s, r):
        return model.z[s, r] >= model.V[s] - M_vol[s]*(1 - model.y[s, r])
    model.z_lower = pyo.Constraint(model.TIERS, rule=z_lower_rule)
    # Define the supplier rebate percentage as a variable r_s with an equality constraint.
    model.r_s = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals, bounds=(0,1))
    def r_s_rule(model, s):
        return model.r_s[s] == sum(tiers_by_supplier[s][r][2]*model.y[s, r] for r in range(len(tiers_by_supplier[s])))
    model.r_s_con = pyo.Constraint(model.SUPPLIERS, rule=r_s_rule)

    # --------------------------
    # Discount Tiers (Volume Discount)
    # --------------------------
    discount_tiers_by_supplier = {}
    for s in model.SUPPLIERS.value:
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
    for s in model.SUPPLIERS.value:
        for r in range(len(discount_tiers_by_supplier[s])):
            discount_index_set.append((s, r))
    model.DTIERS = pyo.Set(initialize=discount_index_set, dimen=2)
    
    model.d_y = pyo.Var(model.DTIERS, domain=pyo.Binary)
    model.d_z = pyo.Var(model.DTIERS, domain=pyo.NonNegativeReals)
    def one_discount_tier_rule(model, s):
        return sum(model.d_y[s, r] for r in range(len(discount_tiers_by_supplier[s]))) == 1
    model.one_discount_tier = pyo.Constraint(model.SUPPLIERS, rule=one_discount_tier_rule)
    M_cost = { s: 1e8 for s in model.SUPPLIERS.value }
    def discount_z_upper1_rule(model, s, r):
        return model.d_z[s, r] <= model.C[s]
    model.discount_z_upper1 = pyo.Constraint(model.DTIERS, rule=discount_z_upper1_rule)
    def discount_z_upper2_rule(model, s, r):
        return model.d_z[s, r] <= M_cost[s]*model.d_y[s, r]
    model.discount_z_upper2 = pyo.Constraint(model.DTIERS, rule=discount_z_upper2_rule)
    def discount_z_lower_rule(model, s, r):
        return model.d_z[s, r] >= model.C[s] - M_cost[s]*(1 - model.d_y[s, r])
    model.discount_z_lower = pyo.Constraint(model.DTIERS, rule=discount_z_lower_rule)
    model.d_s = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals, bounds=(0,1))
    def d_s_rule(model, s):
        return model.d_s[s] == sum(discount_tiers_by_supplier[s][r][2]*model.d_y[s, r] for r in range(len(discount_tiers_by_supplier[s])))
    model.d_s_con = pyo.Constraint(model.SUPPLIERS, rule=d_s_rule)
    
    model.discount_saving = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def discount_saving_rule(model, s):
        return model.discount_saving[s] == sum(discount_tiers_by_supplier[s][r][2]*model.d_z[s, r] for r in range(len(discount_tiers_by_supplier[s])))
    model.discount_saving_con = pyo.Constraint(model.SUPPLIERS, rule=discount_saving_rule)
    
    # --------------------------
    # Define F[s] = 1 - d_s[s] - r_s[s] + f[s]  (but we want to remove the product x[b]*F[s] from the objective)
    # Instead, we will introduce a new variable for each bid row to linearize x[b]*F[s].
    # First, define F[s] as an expression:
    model.F = pyo.Expression(model.SUPPLIERS, rule=lambda model, s: 1 - model.d_s[s] - model.r_s[s] + 0)
    # (Here we set f[s] = 0 because our desired formulation is:
    # Discounted Price = Original Price*(1 - d_s); and rebate is separately applied.)
    # You can adjust this if you wish to incorporate a further product term.
    
    # --------------------------
    # New: Define z_bid[b] = x[b] * F[s] for each bid row b, where s = supplier_of_bid(b).
    # We know x[b] in [0,1] and F[s] is an expression with bounds: 
    # F[s] = 1 - d_s[s] - r_s[s]. With d_s, r_s in [0,1], F[s] is in [-1,1] ideally.
    # To be safe, we assume F[s] ∈ [-1,2] (if needed adjust bounds).
    model.z_bid = pyo.Var(model.BIDS, domain=pyo.Reals, bounds=(-1,2))
    def mcCormick_z_bid_lower(model, b):
        s = bid_data[b]['supplier']
        # x in [0,1], let F_lb = -1 and F_ub = 2.
        return model.z_bid[b] >= - model.x[b]
    def mcCormick_z_bid_lower2(model, b):
        s = bid_data[b]['supplier']
        return model.z_bid[b] >= model.F[s] - 2*(1 - model.x[b])
    def mcCormick_z_bid_upper(model, b):
        s = bid_data[b]['supplier']
        return model.z_bid[b] <= 2*model.x[b]
    def mcCormick_z_bid_upper2(model, b):
        s = bid_data[b]['supplier']
        return model.z_bid[b] <= model.F[s] + 1 - model.x[b]
    model.mc1 = pyo.Constraint(model.BIDS, rule=mcCormick_z_bid_lower)
    model.mc2 = pyo.Constraint(model.BIDS, rule=mcCormick_z_bid_lower2)
    model.mc3 = pyo.Constraint(model.BIDS, rule=mcCormick_z_bid_upper)
    model.mc4 = pyo.Constraint(model.BIDS, rule=mcCormick_z_bid_upper2)
    
    # --------------------------
    # Objective: Maximize Overall Savings (Linearized)
    # For each bid row b, awarded volume = volume[b]*x[b].
    # Discounted Awarded Supplier Price = bid_price[b]*(F[s]) where s = supplier_of_bid(b),
    # but we linearize x[b]*F[s] by z_bid[b].
    # So the contribution is: volume[b] * ( baseline_price[b]*x[b] - bid_price[b]*z_bid[b] ).
    def obj_rule(model):
        return sum(model.volume[b]*(model.baseline_price[b]*model.x[b] - model.bid_price[b]*model.z_bid[b])
                   for b in model.BIDS)
    model.obj = pyo.Objective(rule=obj_rule, sense=pyo.maximize)
    
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
        awarded_fraction = pyo.value(model.x[b])
        if awarded_fraction > 1e-6:
            info = bid_data[b]
            awarded_volume = info['volume'] * awarded_fraction
            supplier = info['supplier']
            # Retrieve supplier's discount and rebate percentages
            d_val = pyo.value(model.d_s[supplier])
            r_val = pyo.value(model.r_s[supplier])
            # Original Awarded Supplier Price:
            orig_price = info['bid_price']
            # Discounted Awarded Supplier Price is orig_price * (1 - d_val)
            discounted_price = orig_price * (1 - d_val)
            baseline_savings = awarded_volume * (info['baseline_price'] - orig_price * (1 - d_val))
            rebate_savings = awarded_volume * (orig_price * (1 - d_val)) * r_val
            output_rows.append({
                "Bid ID": info['orig_bid_id'],
                # "Bid ID Split" will be assigned later
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
    
    # --------------------------
    # Assign "Bid ID Split" Letters per Original Bid ID
    # --------------------------
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
root.title("Bid Optimization Model (Excel I/O with Tabs, Capacity, & Splitting)")

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

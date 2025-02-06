import tkinter as tk
from tkinter import messagebox, scrolledtext
import pyomo.environ as pyo

def run_model():
    # Retrieve text from the bid and rebate data text boxes.
    bid_text = bid_text_widget.get("1.0", tk.END).strip()
    rebate_text = rebate_text_widget.get("1.0", tk.END).strip()

    if not bid_text:
        messagebox.showerror("Input Error", "Please enter bid data.")
        return

    # --------------------------
    # Parse Bid Data
    # --------------------------
    try:
        bid_lines = bid_text.splitlines()
        # Expected header:
        # "Bid ID, Supplier Name, Item Description, Bid Volume, Incumbent Supplier, Current Price, Bid Price, Supplier Capacity"
        header = [col.strip() for col in bid_lines[0].split(',')]
        bid_data = {}
        supplier_capacity = {}
        for line in bid_lines[1:]:
            if line.strip() == "":
                continue
            parts = [p.strip() for p in line.split(',')]
            bid_id = parts[0]
            supplier = parts[1]
            item = parts[2]
            volume = float(parts[3])
            incumbent = parts[4].lower() in ['yes', 'true', '1']
            current_price = float(parts[5])
            bid_price = float(parts[6])
            capacity = float(parts[7])
            bid_data[bid_id] = {
                'supplier': supplier,
                'item': item,
                'volume': volume,
                'incumbent': incumbent,
                'current_price': current_price,
                'bid_price': bid_price,
                'capacity': capacity
            }
            # Save supplier capacity (assume consistency among bids)
            if supplier not in supplier_capacity:
                supplier_capacity[supplier] = capacity
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing bid data: {e}")
        return

    # --------------------------
    # Parse Rebate Data
    # --------------------------
    rebate_data = []
    try:
        if rebate_text:
            rebate_lines = rebate_text.splitlines()
            # Expected header:
            # "Supplier Name, Min Volume, Max Volume, % Rebate"
            header_rebate = [col.strip() for col in rebate_lines[0].split(',')]
            for line in rebate_lines[1:]:
                if line.strip() == "":
                    continue
                parts = [p.strip() for p in line.split(',')]
                supplier = parts[0]
                min_volume = float(parts[1])
                max_volume = float(parts[2])
                rebate_str = parts[3]
                if "%" in rebate_str:
                    rebate_val = float(rebate_str.replace("%", "")) / 100.0
                else:
                    rebate_val = float(rebate_str)
                rebate_data.append({
                    'supplier': supplier,
                    'min_volume': min_volume,
                    'max_volume': max_volume,
                    'rebate': rebate_val
                })
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Error parsing rebate data: {e}")
        return

    # --------------------------
    # Build the Pyomo Model
    # --------------------------
    model = pyo.ConcreteModel()

    # Create sets for bids, items, and suppliers.
    model.BIDS = pyo.Set(initialize=list(bid_data.keys()))
    items = set(bid['item'] for bid in bid_data.values())
    model.ITEMS = pyo.Set(initialize=items)
    suppliers = set(bid['supplier'] for bid in bid_data.values())
    model.SUPPLIERS = pyo.Set(initialize=suppliers)

    # Parameters from bid_data.
    def volume_init(model, b):
        return bid_data[b]['volume']
    model.volume = pyo.Param(model.BIDS, initialize=volume_init)

    def current_price_init(model, b):
        return bid_data[b]['current_price']
    model.current_price = pyo.Param(model.BIDS, initialize=current_price_init)

    def bid_price_init(model, b):
        return bid_data[b]['bid_price']
    model.bid_price = pyo.Param(model.BIDS, initialize=bid_price_init)

    def item_of_bid(model, b):
        return bid_data[b]['item']
    model.item_of_bid = pyo.Param(model.BIDS, initialize=item_of_bid)

    def supplier_of_bid(model, b):
        return bid_data[b]['supplier']
    model.supplier_of_bid = pyo.Param(model.BIDS, initialize=supplier_of_bid)

    # Decision variable: x[b] == 1 if bid b is selected.
    model.x = pyo.Var(model.BIDS, domain=pyo.Binary)

    # --------------------------
    # Calculate Awarded Volume per Supplier
    # --------------------------
    # Create a variable V[s] for total awarded volume for supplier s.
    model.V = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def supplier_volume_rule(model, s):
        return model.V[s] == sum(model.volume[b]*model.x[b] for b in model.BIDS if bid_data[b]['supplier'] == s)
    model.supplier_volume_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_volume_rule)

    # --------------------------
    # Process Rebate Data into Tiers
    # --------------------------
    # Build a dictionary mapping each supplier to its list of rebate tiers.
    # Each tier is a tuple: (min_volume, max_volume, rebate_percentage).
    # We add a "dummy" tier for volumes below the first offered tier.
    tiers_by_supplier = {}
    for s in suppliers:
        tiers = [ (row['min_volume'], row['max_volume'], row['rebate']) 
                  for row in rebate_data if row['supplier'] == s ]
        if tiers:
            tiers.sort(key=lambda x: x[0])
            if tiers[0][0] > 0:
                # Add dummy tier: 0 volume up to one less than the first tier's minimum gets 0% rebate.
                dummy_tier = (0, tiers[0][0]-1, 0.0)
                tiers.insert(0, dummy_tier)
        else:
            # If no rebate info provided, create a single tier with 0 rebate.
            tiers = [(0, supplier_capacity[s], 0.0)]
        tiers_by_supplier[s] = tiers

    # Create a set for (supplier, tier index) for all tiers.
    tier_index_set = []
    for s in suppliers:
        for r in range(len(tiers_by_supplier[s])):
            tier_index_set.append((s, r))
    model.TIERS = pyo.Set(initialize=tier_index_set, dimen=2)

    # --------------------------
    # Rebate Tier Variables and Constraints
    # --------------------------
    # For each supplier s and tier r, let y[s,r] be a binary variable that indicates if tier r is active.
    model.y = pyo.Var(model.TIERS, domain=pyo.Binary)
    # Also, define z[s,r] to represent V[s]*y[s,r] (this product is linearized).
    model.z = pyo.Var(model.TIERS, domain=pyo.NonNegativeReals)

    # For each supplier, exactly one tier must be selected.
    def one_tier_per_supplier_rule(model, s):
        return sum(model.y[s, r] for r in range(len(tiers_by_supplier[s]))) == 1
    model.one_tier_per_supplier = pyo.Constraint(model.SUPPLIERS, rule=one_tier_per_supplier_rule)

    # For each (s,r), if tier r is active then V[s] must be within [min_volume, max_volume].
    # We use the supplier capacity as a "big M" value.
    M = { s: supplier_capacity[s] for s in suppliers }

    def tier_volume_lower_rule(model, s, r):
        min_vol = tiers_by_supplier[s][r][0]
        return model.V[s] >= min_vol * model.y[s, r]
    model.tier_volume_lower = pyo.Constraint(model.TIERS, rule=tier_volume_lower_rule)

    def tier_volume_upper_rule(model, s, r):
        max_vol = tiers_by_supplier[s][r][1]
        return model.V[s] <= max_vol + M[s]*(1 - model.y[s, r])
    model.tier_volume_upper = pyo.Constraint(model.TIERS, rule=tier_volume_upper_rule)

    # Linearize z[s,r] = V[s] * y[s,r] using a standard formulation.
    def z_upper1_rule(model, s, r):
        return model.z[s, r] <= model.V[s]
    model.z_upper1 = pyo.Constraint(model.TIERS, rule=z_upper1_rule)

    def z_upper2_rule(model, s, r):
        return model.z[s, r] <= M[s] * model.y[s, r]
    model.z_upper2 = pyo.Constraint(model.TIERS, rule=z_upper2_rule)

    def z_lower_rule(model, s, r):
        return model.z[s, r] >= model.V[s] - M[s]*(1 - model.y[s, r])
    model.z_lower = pyo.Constraint(model.TIERS, rule=z_lower_rule)

    # Define a variable for the rebate saving for each supplier.
    model.rebate_saving = pyo.Var(model.SUPPLIERS, domain=pyo.NonNegativeReals)
    def rebate_saving_rule(model, s):
        # The rebate saving is the sum over tiers of (rebate percentage * z[s,r]).
        return model.rebate_saving[s] == sum(tiers_by_supplier[s][r][2] * model.z[s, r]
                                              for r in range(len(tiers_by_supplier[s])))
    model.rebate_saving_con = pyo.Constraint(model.SUPPLIERS, rule=rebate_saving_rule)

    # --------------------------
    # Objective: Maximize Bid Savings + Rebate Bonus
    # --------------------------
    # Bid savings: (current_price - bid_price)*volume for each selected bid.
    # Rebate bonus: Sum of rebate_saving over all suppliers.
    def obj_rule(model):
        bid_savings = sum((model.current_price[b] - model.bid_price[b]) * model.volume[b] * model.x[b]
                          for b in model.BIDS)
        rebate_bonus = sum(model.rebate_saving[s] for s in model.SUPPLIERS)
        return bid_savings + rebate_bonus
    model.obj = pyo.Objective(rule=obj_rule, sense=pyo.maximize)

    # --------------------------
    # Other Constraints: One bid per item and supplier capacity (already enforced by V[s] <= capacity).
    # --------------------------
    def one_bid_per_item_rule(model, i):
        return sum(model.x[b] for b in model.BIDS if bid_data[b]['item'] == i) == 1
    model.one_bid_per_item = pyo.Constraint(model.ITEMS, rule=one_bid_per_item_rule)

    def supplier_capacity_rule(model, s):
        return model.V[s] <= supplier_capacity[s]
    model.supplier_capacity_con = pyo.Constraint(model.SUPPLIERS, rule=supplier_capacity_rule)

    # --------------------------
    # Solve the Model
    # --------------------------
    try:
        solver = pyo.SolverFactory('glpk')
        result = solver.solve(model, tee=False)
    except Exception as e:
        messagebox.showerror("Solver Error", f"Error during model solve: {e}")
        return

    # --------------------------
    # Display the Results
    # --------------------------
    output_text = ""
    output_text += "--- Selected Bids ---\n"
    for b in model.BIDS:
        if pyo.value(model.x[b]) > 0.5:
            info = bid_data[b]
            output_text += (f"Bid {b}: Item={info['item']}, Supplier={info['supplier']}, "
                            f"Volume={info['volume']}, Current Price={info['current_price']}, "
                            f"Bid Price={info['bid_price']}\n")
    output_text += "\n--- Supplier Summary ---\n"
    for s in suppliers:
        awarded_vol = pyo.value(model.V[s])
        rebate_save = pyo.value(model.rebate_saving[s])
        # Determine which tier was chosen.
        chosen_tier = None
        for r in range(len(tiers_by_supplier[s])):
            if pyo.value(model.y[s, r]) > 0.5:
                chosen_tier = tiers_by_supplier[s][r]
                break
        tier_str = f"Tier (min={chosen_tier[0]}, max={chosen_tier[1]}, rebate={chosen_tier[2]*100:.1f}%)" if chosen_tier else "None"
        output_text += (f"Supplier {s}: Awarded Volume = {awarded_vol}, "
                        f"Rebate Saving = {rebate_save:.2f}, Selected {tier_str}\n")

    result_text_widget.delete("1.0", tk.END)
    result_text_widget.insert(tk.END, output_text)

# --------------------------
# Create the Tkinter GUI
# --------------------------
root = tk.Tk()
root.title("Bid Optimization with Rebates")

# Bid Data Input
bid_label = tk.Label(root, text="Enter Bid Data (CSV Format):\n"
                                "Header: Bid ID, Supplier Name, Item Description, Bid Volume, Incumbent Supplier, Current Price, Bid Price, Supplier Capacity")
bid_label.pack(pady=(10,0))
bid_text_widget = scrolledtext.ScrolledText(root, width=100, height=10)
bid_text_widget.pack(padx=10, pady=5)
bid_text_widget.insert(tk.END,
"""Bid ID,Supplier Name,Item Description,Bid Volume,Incumbent Supplier,Current Price,Bid Price,Supplier Capacity
B1,Supplier A,Widget,1000,Yes,50,45,1500
B2,Supplier B,Widget,1000,No,50,44,2000
B3,Supplier C,Widget,1000,No,50,46,1800
B4,Supplier A,Gadget,500,No,30,28,1500
B5,Supplier B,Gadget,500,Yes,30,27,2000
B6,Supplier C,Gadget,500,No,30,29,1800""")

# Rebate Data Input
rebate_label = tk.Label(root, text="Enter Rebate Data (CSV Format):\n"
                                   "Header: Supplier Name, Min Volume, Max Volume, % Rebate")
rebate_label.pack(pady=(10,0))
rebate_text_widget = scrolledtext.ScrolledText(root, width=100, height=5)
rebate_text_widget.pack(padx=10, pady=5)
rebate_text_widget.insert(tk.END,
"""Supplier Name,Min Volume,Max Volume,% Rebate
Supplier A,500,1000,1.0%
Supplier A,1001,1500,1.5%
Supplier B,500,1200,1.2%
Supplier B,1201,2000,2.0%
Supplier C,500,800,0.8%
Supplier C,801,1800,1.7%""")

# Run Button
run_button = tk.Button(root, text="Run Model", command=run_model)
run_button.pack(pady=10)

# Results Output
result_label = tk.Label(root, text="Results:")
result_label.pack()
result_text_widget = scrolledtext.ScrolledText(root, width=100, height=10)
result_text_widget.pack(padx=10, pady=5)

root.mainloop()

import os
import pandas as pd
import PySimpleGUI as sg
import pulp

#############################################
# DEFAULT DATA (used if no file is uploaded)
#############################################

# Define suppliers and items (fixed for our example)
suppliers = ['A', 'B', 'C']
items = ['item1', 'item2', 'item3']

# Default Item Attributes (each item now has Facility, BusinessUnit, Incumbent, Capacity Group)
default_item_attributes = {
    'item1': {'BusinessUnit': 'A', 'Incumbent': 'A', 'Capacity Group': 'Widgets', 'Facility': 'Facility1'},
    'item2': {'BusinessUnit': 'B', 'Incumbent': 'B', 'Capacity Group': 'Gadgets', 'Facility': 'Facility2'},
    'item3': {'BusinessUnit': 'A', 'Incumbent': 'C', 'Capacity Group': 'Widgets', 'Facility': 'Facility3'}
}

# Default Prices: dictionary with key (supplier, item)
default_price = {
    ('A', 'item1'): 50,
    ('A', 'item2'): 70,
    ('A', 'item3'): 55,
    ('B', 'item1'): 60,
    ('B', 'item2'): 80,
    ('B', 'item3'): 65,
    ('C', 'item1'): 55,
    ('C', 'item2'): 75,
    ('C', 'item3'): 60
}

# Default Demand: dictionary with key item
default_demand = {
    'item1': 600,
    'item2': 1000,
    'item3': 800
}

# Default Rebate Structure: multi-tier per supplier
# Each tier is defined as (min_volume, max_volume, rebate_percentage)
default_rebate_tiers = {
    'A': [(0, 500, 0.0), (500, 1000, 0.10), (1000, float('inf'), 0.10)],
    'B': [(0, 500, 0.0), (500, 1000, 0.05), (1000, float('inf'), 0.05)],
    'C': [(0, 700, 0.0), (700, 1500, 0.08), (1500, float('inf'), 0.08)]
}

# Default Discount Structure: multi-tier per supplier
# Each tier is defined as (min_volume, max_volume, discount_percentage)
default_discount_tiers = {
    'A': [(0, 500, 0.0), (500, 1000, 0.05), (1000, float('inf'), 0.07)],
    'B': [(0, 500, 0.0), (500, 1000, 0.03), (1000, float('inf'), 0.03)],
    'C': [(0, 500, 0.0), (500, 1500, 0.04), (1500, float('inf'), 0.04)]
}

# Default Baseline Price: dictionary with key item
default_baseline_price = {
    'item1': 45,
    'item2': 65,
    'item3': 75
}

# Default Per‑Item Capacity (if not using global)
default_per_item_capacity = {
    ('A', 'item1'): 500,
    ('A', 'item2'): 400,
    ('A', 'item3'): 300,
    ('B', 'item1'): 400,
    ('B', 'item2'): 800,
    ('B', 'item3'): 600,
    ('C', 'item1'): 300,
    ('C', 'item2'): 500,
    ('C', 'item3'): 700
}

# Default Global Capacity DataFrame (each supplier has a capacity for each capacity group)
default_global_capacity_df = pd.DataFrame({
    "Supplier Name": ["A", "A", "B", "B", "C", "C"],
    "Capacity Group": ["Widgets", "Gadgets", "Widgets", "Gadgets", "Widgets", "Gadgets"],
    "Capacity": [1000, 900, 1200, 1100, 800, 950]
})
default_global_capacity = {
    (str(row["Supplier Name"]).strip(), str(row["Capacity Group"]).strip()): row["Capacity"]
    for idx, row in default_global_capacity_df.iterrows()
}

# For big‑M values, we compute U_volume and U_spend from per‑item capacity.
def compute_U_volume(per_item_cap):
    total = {}
    for s in suppliers:
        tot = 0
        for j in items:
            tot += per_item_cap.get((s, j), 0)
        total[s] = tot
    return total

default_U_volume = compute_U_volume(default_per_item_capacity)

def compute_U_spend(per_item_cap):
    total = {}
    for s in suppliers:
        tot = 0
        for j in items:
            tot += default_price.get((s, j), 0) * per_item_cap.get((s, j), 0)
        total[s] = tot
    return total

default_U_spend = compute_U_spend(default_per_item_capacity)

epsilon = 1e-6

#############################################
# OPTIMIZATION MODEL FUNCTION
#############################################
def run_optimization(use_global, capacity_data, demand_data, item_attr_data, price_data,
                     rebate_tiers, discount_tiers, baseline_price_data):
    """
    Runs the optimization model with multi-tier rebates and volume discounts.
    Parameters:
      use_global: Boolean. True if capacity_data uses keys (supplier, capacity group), else (supplier, item).
      demand_data: dict mapping item -> demand.
      item_attr_data: dict mapping item -> attributes.
      price_data: dict mapping (supplier, item) -> price.
      rebate_tiers: dict mapping supplier -> list of (min, max, rebate%) tuples.
      discount_tiers: dict mapping supplier -> list of (min, max, discount%) tuples.
      baseline_price_data: dict mapping item -> baseline price.
    Returns: (output_file, feasibility_notes, model_status)
    """
    U_volume = default_U_volume   # For big‑M constraints.
    U_spend = default_U_spend

    lp_problem = pulp.LpProblem("Sourcing_with_MultiTier_Rebates_Discounts", pulp.LpMinimize)

    # Decision Variables: x[s,j] = awarded units.
    x = {}
    for s in suppliers:
        for j in items:
            x[(s, j)] = pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')

    # For each supplier: Base spend S0[s] and Effective spend S[s].
    S0 = {}
    S = {}
    for s in suppliers:
        S0[s] = pulp.LpVariable(f"S0_{s}", lowBound=0, cat='Continuous')
        S[s]  = pulp.LpVariable(f"S_{s}", lowBound=0, cat='Continuous')

    # Total awarded volume V[s].
    V = {}
    for s in suppliers:
        V[s] = pulp.LpVariable(f"V_{s}", lowBound=0, cat='Continuous')

    # Create binary variables for each discount tier per supplier.
    z_discount = {}
    for s in suppliers:
        z_discount[s] = {}
        tiers = discount_tiers[s]
        for k in range(len(tiers)):
            z_discount[s][k] = pulp.LpVariable(f"z_discount_{s}_{k}", cat='Binary')
        lp_problem += pulp.lpSum(z_discount[s][k] for k in range(len(tiers))) == 1, f"DiscountTierSelect_{s}"

    # Create binary variables for each rebate tier per supplier.
    y_rebate = {}
    for s in suppliers:
        y_rebate[s] = {}
        tiers = rebate_tiers[s]
        for k in range(len(tiers)):
            y_rebate[s][k] = pulp.LpVariable(f"y_rebate_{s}_{k}", cat='Binary')
        lp_problem += pulp.lpSum(y_rebate[s][k] for k in range(len(tiers))) == 1, f"RebateTierSelect_{s}"

    # Discount amount for supplier s.
    d = {}
    for s in suppliers:
        d[s] = pulp.LpVariable(f"d_{s}", lowBound=0, cat='Continuous')
    # Rebate amount for supplier s.
    rebate_var = {}
    for s in suppliers:
        rebate_var[s] = pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous')

    # Objective: Minimize total effective cost = sum(S[s] - rebate_var[s])
    lp_problem += pulp.lpSum(S[s] - rebate_var[s] for s in suppliers), "Total_Effective_Cost"

    # 1. Demand Constraints.
    for j in items:
        lp_problem += pulp.lpSum(x[(s, j)] for s in suppliers) == demand_data[j], f"Demand_{j}"

    # 2. Capacity Constraints.
    if use_global:
        supplier_capacity_groups = {}
        all_groups = set(item_attr_data[j].get("Capacity Group", None) for j in items if item_attr_data[j].get("Capacity Group", None))
        for s in suppliers:
            supplier_capacity_groups[s] = {g: [] for g in all_groups}
            for j in items:
                group = item_attr_data[j].get("Capacity Group", None)
                if group is not None:
                    supplier_capacity_groups[s][group].append(j)
        for s in suppliers:
            for group, item_list in supplier_capacity_groups[s].items():
                cap = capacity_data.get((s, group), 1e9)
                lp_problem += pulp.lpSum(x[(s, j)] for j in item_list) <= cap, f"GlobalCapacity_{s}_{group}"
    else:
        for s in suppliers:
            for j in items:
                cap = capacity_data.get((s, j), 1e9)
                lp_problem += x[(s, j)] <= cap, f"PerItemCapacity_{s}_{j}"

    # 3. Base Spend and Total Volume.
    for s in suppliers:
        lp_problem += S0[s] == pulp.lpSum(price_data[(s, j)] * x[(s, j)] for j in items), f"BaseSpend_{s}"
        lp_problem += V[s] == pulp.lpSum(x[(s, j)] for j in items), f"Volume_{s}"

    # 4. Discount Activation & Linearization (multi-tier).
    for s in suppliers:
        tiers = discount_tiers[s]
        M_discount = U_spend[s]
        for k, (Dmin, Dmax, Dperc) in enumerate(tiers):
            lp_problem += V[s] >= Dmin * z_discount[s][k], f"DiscountTierMin_{s}_{k}"
            if Dmax < float('inf'):
                lp_problem += V[s] <= Dmax + M_discount*(1 - z_discount[s][k]), f"DiscountTierMax_{s}_{k}"
            lp_problem += d[s] >= Dperc * S0[s] - M_discount*(1 - z_discount[s][k]), f"DiscountTierLower_{s}_{k}"
            lp_problem += d[s] <= Dperc * S0[s] + M_discount*(1 - z_discount[s][k]), f"DiscountTierUpper_{s}_{k}"

    # 5. Effective Spend Calculation.
    for s in suppliers:
        lp_problem += S[s] == S0[s] - d[s], f"EffectiveSpend_{s}"

    # 6. Rebate Activation & Linearization (multi-tier).
    for s in suppliers:
        tiers = rebate_tiers[s]
        M_rebate = U_spend[s]
        for k, (Rmin, Rmax, Rperc) in enumerate(tiers):
            lp_problem += V[s] >= Rmin * y_rebate[s][k], f"RebateTierMin_{s}_{k}"
            if Rmax < float('inf'):
                lp_problem += V[s] <= Rmax + M_rebate*(1 - y_rebate[s][k]), f"RebateTierMax_{s}_{k}"
            lp_problem += rebate_var[s] >= Rperc * S[s] - M_rebate*(1 - y_rebate[s][k]), f"RebateTierLower_{s}_{k}"
            lp_problem += rebate_var[s] <= Rperc * S[s] + M_rebate*(1 - y_rebate[s][k]), f"RebateTierUpper_{s}_{k}"

    # 7. Custom Constraint: For items with Business Unit "A", award must go only to the incumbent.
    for j in items:
        if item_attr_data[j]['BusinessUnit'] == 'A':
            incumbent_supplier = item_attr_data[j]['Incumbent']
            for s in suppliers:
                if s != incumbent_supplier:
                    lp_problem += x[(s, j)] == 0, f"IncumbentConstraint_{j}_{s}"

    lp_problem.solve()
    model_status = pulp.LpStatus[lp_problem.status]

    # Build feasibility notes.
    feasibility_notes = ""
    if model_status == "Infeasible":
        feasibility_notes += "Model is infeasible. Possible causes:\n"
        for j in items:
            if item_attr_data[j]['BusinessUnit'] == 'A':
                incumbent = item_attr_data[j]['Incumbent']
                if use_global:
                    group = item_attr_data[j].get("Capacity Group", "Unknown")
                    allowed = capacity_data.get((incumbent, group), 0)
                    feasibility_notes += f"  Item {j} (BU A, incumbent {incumbent}): demand = {demand_data[j]}, capacity for group {group} = {allowed}\n"
                else:
                    allowed = capacity_data.get((incumbent, j), 0)
                    feasibility_notes += f"  Item {j} (BU A, incumbent {incumbent}): demand = {demand_data[j]}, capacity = {allowed}\n"
            else:
                if use_global:
                    group = item_attr_data[j].get("Capacity Group", "Unknown")
                    total = sum(capacity_data.get((s, group), 0) for s in suppliers)
                    feasibility_notes += f"  Item {j}: demand = {demand_data[j]}, total capacity for group {group} = {total}\n"
                else:
                    total = sum(capacity_data.get((s, j), 0) for s in suppliers)
                    feasibility_notes += f"  Item {j}: demand = {demand_data[j]}, total capacity = {total}\n"
        feasibility_notes += "Please review capacities and/or bidding data for potential issues.\n"
    else:
        feasibility_notes = "Model is optimal."

    #############################################
    # Build Excel Output (Results & Feasibility Notes)
    #############################################
    excel_rows = []
    letter_list = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    # For each item, output one row per supplier awarded a positive volume.
    for idx, j in enumerate(items, start=1):
        awarded_list = []
        for s in suppliers:
            award_val = pulp.value(x[(s, j)])
            if award_val is None:
                award_val = 0
            if award_val > 0:
                awarded_list.append((s, award_val))
        if not awarded_list:
            awarded_list = [("None", 0)]
        awarded_list.sort(key=lambda tup: (-tup[1], tup[0]))
        for i, (s, award_val) in enumerate(awarded_list):
            bid_split = letter_list[i] if i < len(letter_list) else f"Split{i+1}"
            orig_price = price_data.get((s, j), 0)
            # Determine active discount tier for supplier s.
            active_discount = 0
            for k, tier in enumerate(discount_tiers[s]):
                if pulp.value(z_discount[s][k]) is not None and pulp.value(z_discount[s][k]) >= 0.5:
                    active_discount = tier[2]
                    break
            discount_pct = active_discount
            discounted_price = orig_price * (1 - discount_pct)
            awarded_spend = discounted_price * award_val
            base_price = baseline_price_data[j]
            baseline_spend = base_price * award_val
            baseline_savings = baseline_spend - awarded_spend
            if use_global:
                group = item_attr_data[j].get("Capacity Group", "Unknown")
                awarded_capacity = capacity_data.get((s, group), 0)
            else:
                awarded_capacity = capacity_data.get((s, j), 0)
            # Corrected rebate savings: Awarded Supplier Spend * (active rebate percentage)
            active_rebate = 0
            for k, tier in enumerate(rebate_tiers[s]):
                if pulp.value(y_rebate[s][k]) is not None and pulp.value(y_rebate[s][k]) >= 0.5:
                    active_rebate = tier[2]
                    break
            rebate_savings = awarded_spend * active_rebate
            facility_val = item_attr_data[j].get("Facility", "")
            row = {
                "Bid ID": idx,
                "Capacity Group": item_attr_data[j].get("Capacity Group", "") if use_global else "",
                "Bid ID Split": bid_split,
                "Facility": facility_val,
                "Incumbent": item_attr_data[j]["Incumbent"],
                "Baseline Price": base_price,
                "Baseline Spend": baseline_spend,
                "Awarded Supplier": s,
                "Original Awarded Supplier Price": orig_price,
                "Percentage Volume Discount": f"{discount_pct*100:.0f}%" if s in discount_tiers else "0%",
                "Discounted Awarded Supplier Price": discounted_price,
                "Awarded Supplier Spend": awarded_spend,
                "Awarded Volume": award_val,
                "Awarded Supplier Capacity": awarded_capacity,
                "Baseline Savings": baseline_savings,
                "Rebate %": f"{active_rebate*100:.0f}%" if active_rebate is not None else "0%",
                "Rebate Savings": rebate_savings
            }
            excel_rows.append(row)
    
    df_results = pd.DataFrame(excel_rows)
    if use_global:
        cols = ["Bid ID", "Capacity Group", "Bid ID Split", "Facility", "Incumbent", "Baseline Price", "Baseline Spend",
                "Awarded Supplier", "Original Awarded Supplier Price", "Percentage Volume Discount",
                "Discounted Awarded Supplier Price", "Awarded Supplier Spend", "Awarded Volume",
                "Awarded Supplier Capacity", "Baseline Savings", "Rebate %", "Rebate Savings"]
    else:
        cols = ["Bid ID", "Bid ID Split", "Facility", "Incumbent", "Baseline Price", "Baseline Spend",
                "Awarded Supplier", "Original Awarded Supplier Price", "Percentage Volume Discount",
                "Discounted Awarded Supplier Price", "Awarded Supplier Spend", "Awarded Volume",
                "Awarded Supplier Capacity", "Baseline Savings", "Rebate %", "Rebate Savings"]
    df_results = df_results[cols]
    
    df_feasibility = pd.DataFrame({"Feasibility Notes": [feasibility_notes]})
    
    home_dir = os.path.expanduser("~")
    downloads_folder = os.path.join(home_dir, "Downloads")
    output_file = os.path.join(downloads_folder, "optimization_results.xlsx")
    
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_results.to_excel(writer, sheet_name="Results", index=False)
        df_feasibility.to_excel(writer, sheet_name="Feasibility Notes", index=False)
    
    return output_file, feasibility_notes, model_status

#############################################
# GUI CODE USING PySimpleGUI
#############################################

# Create sample Excel files for download.
# For Global Capacity, include only the Global Capacity table; for Per Item, include only that table.
df_global_capacity = default_global_capacity_df.copy()
df_per_item_capacity = pd.DataFrame([
    {"Supplier Name": s, "Item": j, "Capacity": default_per_item_capacity.get((s, j), None)}
    for s in suppliers for j in items
])
df_demand = pd.DataFrame([
    {"Item": j, "Demand": default_demand[j]} for j in items
])
df_item_attr = pd.DataFrame([
    {"Item": j, **default_item_attributes[j]} for j in items
])
df_price = pd.DataFrame([
    {"Supplier Name": s, "Item": j, "Price": default_price.get((s, j), None)}
    for s in suppliers for j in items
])
df_rebate = pd.DataFrame([
    {"Supplier Name": s, "Tier": k+1, "Min Volume": tier[0], "Max Volume": tier[1], "Rebate Percentage": tier[2]}
    for s in suppliers for k, tier in enumerate(default_rebate_tiers[s])
])
df_discount = pd.DataFrame([
    {"Supplier Name": s, "Tier": k+1, "Min Volume": tier[0], "Max Volume": tier[1], "Discount Percentage": tier[2]}
    for s in suppliers for k, tier in enumerate(default_discount_tiers[s])
])
df_baseline = pd.DataFrame([
    {"Item": j, "Baseline Price": default_baseline_price[j]} for j in items
])

layout = [
    [sg.Text("Upload Excel File (Optional):"), sg.Input(key="-FILE-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Text("Select Capacity Input Type:")],
    [sg.Radio("Global Capacity", "CAP_TYPE", key="-GLOBAL-", default=True),
     sg.Radio("Per Item Capacity", "CAP_TYPE", key="-PERITEM-")],
    [sg.Text("Select Data Tab for Capacity:"), sg.Combo(values=[], key="-CAPTAB-", size=(30, 1))],
    [sg.Text("Select Data Tab for Demand:"), sg.Combo(values=[], key="-DEMTAB-", size=(30, 1))],
    [sg.Text("Select Data Tab for Item Attributes:"), sg.Combo(values=[], key="-ITEMTAB-", size=(30, 1))],
    [sg.Text("Select Data Tab for Price:"), sg.Combo(values=[], key="-PRICETAB-", size=(30, 1))],
    [sg.Text("Select Data Tab for Rebate Structure:"), sg.Combo(values=[], key="-REBATTAB-", size=(30, 1))],
    [sg.Text("Select Data Tab for Discount Structure:"), sg.Combo(values=[], key="-DISCTAB-", size=(30, 1))],
    [sg.Text("Select Data Tab for Baseline Price:"), sg.Combo(values=[], key="-BASETAB-", size=(30, 1))],
    [sg.Button("Load Excel File")],
    [sg.Button("Download All Example Data")],
    [sg.Button("Run Optimization")],
    [sg.Button("Exit")]
]

window = sg.Window("Capacity & Optimization Input", layout)
uploaded_file = None

while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    elif event == "Load Excel File":
        file_path = values["-FILE-"]
        if file_path:
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                window["-CAPTAB-"].update(values=sheet_names, value=sheet_names[0])
                window["-DEMTAB-"].update(values=sheet_names, value=sheet_names[0])
                window["-ITEMTAB-"].update(values=sheet_names, value=sheet_names[0])
                window["-PRICETAB-"].update(values=sheet_names, value=sheet_names[0])
                window["-REBATTAB-"].update(values=sheet_names, value=sheet_names[0])
                window["-DISCTAB-"].update(values=sheet_names, value=sheet_names[0])
                window["-BASETAB-"].update(values=sheet_names, value=sheet_names[0])
                sg.popup("Excel file loaded successfully.\nSelect the corresponding sheet for each data table.")
                uploaded_file = file_path
            except Exception as e:
                sg.popup_error("Error reading Excel file:", e)
        else:
            sg.popup("No file selected.")
    elif event == "Download All Example Data":
        home_dir = os.path.expanduser("~")
        downloads_folder = os.path.join(home_dir, "Downloads")
        out_file = os.path.join(downloads_folder, "all_example_data.xlsx")
        with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
            if values["-GLOBAL-"]:
                df_global_capacity.to_excel(writer, sheet_name="Capacity", index=False)
            else:
                df_per_item_capacity.to_excel(writer, sheet_name="Capacity", index=False)
            df_demand.to_excel(writer, sheet_name="Demand", index=False)
            df_item_attr.to_excel(writer, sheet_name="Item Attributes", index=False)
            df_price.to_excel(writer, sheet_name="Price", index=False)
            df_rebate.to_excel(writer, sheet_name="Rebate Structure", index=False)
            df_discount.to_excel(writer, sheet_name="Discount Structure", index=False)
            df_baseline.to_excel(writer, sheet_name="Baseline Price", index=False)
        sg.popup("Example data saved to:", out_file)
    elif event == "Run Optimization":
        use_global = values["-GLOBAL-"]
        if uploaded_file:
            try:
                df_cap = pd.read_excel(uploaded_file, sheet_name=values["-CAPTAB-"])
                df_dem = pd.read_excel(uploaded_file, sheet_name=values["-DEMTAB-"])
                df_item = pd.read_excel(uploaded_file, sheet_name=values["-ITEMTAB-"])
                df_price = pd.read_excel(uploaded_file, sheet_name=values["-PRICETAB-"])
                df_reb = pd.read_excel(uploaded_file, sheet_name=values["-REBATTAB-"])
                df_disc = pd.read_excel(uploaded_file, sheet_name=values["-DISCTAB-"])
                df_base = pd.read_excel(uploaded_file, sheet_name=values["-BASETAB-"])
            except Exception as e:
                sg.popup_error("Error reading one or more data tables from Excel:", e)
                continue
            if use_global:
                cap_dict = {}
                for idx, row in df_cap.iterrows():
                    key = (str(row["Supplier Name"]).strip(), str(row["Capacity Group"]).strip())
                    cap_dict[key] = row["Capacity"]
            else:
                cap_dict = {}
                for idx, row in df_cap.iterrows():
                    key = (str(row["Supplier Name"]).strip(), str(row["Item"]).strip())
                    cap_dict[key] = row["Capacity"]
            demand_dict = {row["Item"]: row["Demand"] for idx, row in df_dem.iterrows()}
            item_attr_dict = {}
            for idx, row in df_item.iterrows():
                item = str(row["Item"]).strip()
                item_attr_dict[item] = {
                    "BusinessUnit": row["BusinessUnit"],
                    "Incumbent": row["Incumbent"],
                    "Capacity Group": row["Capacity Group"],
                    "Facility": row["Facility"]
                }
            price_dict = {}
            for idx, row in df_price.iterrows():
                key = (str(row["Supplier Name"]).strip(), str(row["Item"]).strip())
                price_dict[key] = row["Price"]
            rebate_tiers = {}
            for idx, row in df_reb.iterrows():
                supplier = str(row["Supplier Name"]).strip()
                rebate_tiers.setdefault(supplier, []).append((row["Min Volume"], row["Max Volume"], row["Rebate Percentage"]))
            discount_tiers = {}
            for idx, row in df_disc.iterrows():
                supplier = str(row["Supplier Name"]).strip()
                discount_tiers.setdefault(supplier, []).append((row["Min Volume"], row["Max Volume"], row["Discount Percentage"]))
            for s in suppliers:
                if s not in rebate_tiers:
                    rebate_tiers[s] = default_rebate_tiers[s]
                if s not in discount_tiers:
                    discount_tiers[s] = default_discount_tiers[s]
            baseline_dict = {row["Item"]: row["Baseline Price"] for idx, row in df_base.iterrows()}
        else:
            cap_dict = default_global_capacity if use_global else default_per_item_capacity
            demand_dict = default_demand
            item_attr_dict = default_item_attributes
            price_dict = default_price
            rebate_tiers = default_rebate_tiers
            discount_tiers = default_discount_tiers
            baseline_dict = default_baseline_price

        output_file, feasibility_notes, model_status = run_optimization(use_global, cap_dict,
                                                                         demand_dict, item_attr_dict,
                                                                         price_dict, rebate_tiers,
                                                                         discount_tiers,
                                                                         baseline_dict)
        sg.popup(f"Model Status: {model_status}\nExcel results written to:\n{output_file}\n\nFeasibility Notes:\n{feasibility_notes}")

window.close()

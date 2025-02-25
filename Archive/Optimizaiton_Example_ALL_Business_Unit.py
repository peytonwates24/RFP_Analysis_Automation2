import os
import pandas as pd
import pulp

# --------------------------
# 1. Problem Data
# --------------------------

# Define suppliers and items.
suppliers = ['A', 'B', 'C']
items = ['item1', 'item2', 'item3']

# Additional supplier attributes.
supplier_attributes = {
    'A': {'Facility': 'Facility1'},
    'B': {'Facility': 'Facility2'},
    'C': {'Facility': 'Facility3'}
}

# Additional item attributes.
# For each item, we specify its Business Unit and which supplier is its incumbent.
item_attributes = {
    'item1': {'BusinessUnit': 'A', 'Incumbent': 'A'},
    'item2': {'BusinessUnit': 'B', 'Incumbent': 'B'},
    'item3': {'BusinessUnit': 'A', 'Incumbent': 'C'}
}

# Prices (dollars per unit) for each supplier–item combination.
price = {
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

# Supplier capacities: maximum units that can be supplied for each item.
capacity = {
    ('A', 'item1'): 5000,
    ('A', 'item2'): 4000,
    ('A', 'item3'): 3000,
    ('B', 'item1'): 4000,
    ('B', 'item2'): 8000,
    ('B', 'item3'): 6000,
    ('C', 'item1'): 3000,
    ('C', 'item2'): 5000,
    ('C', 'item3'): 7000
}

# Demand for each item.
demand = {
    'item1': 600,
    'item2': 1000,
    'item3': 800
}

# Rebate structure (applied on effective spend).
rebate_threshold = {'A': 600, 'B': 500, 'C': 700}
rebate_percentage = {'A': 0.10, 'B': 0.05, 'C': 0.08}

# Volume discount structure (affects bid price).
discount_threshold = {'A': 500, 'B': 500, 'C': 500}
discount_percentage = {'A': 0.05, 'B': 0.03, 'C': 0.04}

# For the purpose of the Excel output, define baseline prices (external data).
baseline_price = {
    'item1': 45,
    'item2': 65,
    'item3': 75
}

# --------------------------
# Big-M Parameters and Epsilon
# --------------------------
# U_volume: Upper bound on awarded volume for each supplier (sum of capacities over all items).
U_volume = {s: sum(capacity[(s, j)] for j in items) for s in suppliers}

# U_spend: Upper bound on base spend for each supplier.
U_spend = {s: sum(price[(s, j)] * capacity[(s, j)] for j in items) for s in suppliers}

epsilon = 1e-6

# --------------------------
# 2. Define the MILP Problem
# --------------------------
lp_problem = pulp.LpProblem("Sourcing_with_Attributes_and_Incumbency", pulp.LpMinimize)

# --------------------------
# 3. Decision Variables
# --------------------------
# x[s, j]: Awarded units from supplier s for item j.
x = {}
for s in suppliers:
    for j in items:
        x[(s, j)] = pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')

# For each supplier:
# S0[s]: Base spend (without discount).
# S[s]: Effective spend after discount.
S0 = {}
S = {}
for s in suppliers:
    S0[s] = pulp.LpVariable(f"S0_{s}", lowBound=0, cat='Continuous')
    S[s]  = pulp.LpVariable(f"S_{s}", lowBound=0, cat='Continuous')

# V[s]: Total awarded volume (sum over items) for supplier s.
V = {}
for s in suppliers:
    V[s] = pulp.LpVariable(f"V_{s}", lowBound=0, cat='Continuous')

# Rebate variables:
# y[s]: Binary variable; 1 if rebate is activated.
# rebate[s]: Rebate amount.
y = {}
rebate = {}
for s in suppliers:
    y[s] = pulp.LpVariable(f"y_{s}", cat='Binary')
    rebate[s] = pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous')

# Discount variables:
# z[s]: Binary variable; 1 if discount is activated.
# d[s]: Discount amount.
z = {}
d = {}
for s in suppliers:
    z[s] = pulp.LpVariable(f"z_{s}", cat='Binary')
    d[s] = pulp.LpVariable(f"d_{s}", lowBound=0, cat='Continuous')

# --------------------------
# 4. Objective Function
# --------------------------
# Minimize total effective cost = sum_{s in suppliers} (Effective Spend S[s] - rebate[s]).
lp_problem += pulp.lpSum(S[s] - rebate[s] for s in suppliers), "Total_Effective_Cost"

# --------------------------
# 5. Constraints
# --------------------------

# 5.1 Demand Constraints: For each item, total awarded units must equal demand.
for j in items:
    lp_problem += pulp.lpSum(x[(s, j)] for s in suppliers) == demand[j], f"Demand_{j}"

# 5.2 Capacity Constraints: Awarded units for each supplier and item cannot exceed capacity.
for s in suppliers:
    for j in items:
        lp_problem += x[(s, j)] <= capacity[(s, j)], f"Capacity_{s}_{j}"

# 5.3 Base Spend and Total Volume Calculation:
for s in suppliers:
    lp_problem += S0[s] == pulp.lpSum(price[(s, j)] * x[(s, j)] for j in items), f"BaseSpend_{s}"
    lp_problem += V[s] == pulp.lpSum(x[(s, j)] for j in items), f"Volume_{s}"

# 5.4 Discount Activation Constraints:
for s in suppliers:
    lp_problem += V[s] >= discount_threshold[s] * z[s], f"DiscountThresholdLower_{s}"
    lp_problem += V[s] <= (discount_threshold[s] - epsilon) + U_volume[s] * z[s], f"DiscountThresholdUpper_{s}"

# 5.5 Discount Linearization Constraints:
for s in suppliers:
    lp_problem += d[s] <= discount_percentage[s] * S0[s], f"DiscountUpper1_{s}"
    lp_problem += d[s] <= discount_percentage[s] * U_spend[s] * z[s], f"DiscountUpper2_{s}"
    lp_problem += d[s] >= discount_percentage[s] * S0[s] - discount_percentage[s] * U_spend[s] * (1 - z[s]), f"DiscountLower_{s}"

# 5.6 Effective Spend Calculation:
for s in suppliers:
    lp_problem += S[s] == S0[s] - d[s], f"EffectiveSpend_{s}"

# 5.7 Rebate Activation Constraints:
for s in suppliers:
    lp_problem += V[s] >= rebate_threshold[s] * y[s], f"RebateThresholdLower_{s}"
    lp_problem += V[s] <= (rebate_threshold[s] - epsilon) + U_volume[s] * y[s], f"RebateThresholdUpper_{s}"

# 5.8 Rebate Linearization Constraints:
for s in suppliers:
    lp_problem += rebate[s] <= rebate_percentage[s] * S[s], f"RebateUpper1_{s}"
    lp_problem += rebate[s] <= rebate_percentage[s] * U_spend[s] * y[s], f"RebateUpper2_{s}"
    lp_problem += rebate[s] >= rebate_percentage[s] * S[s] - rebate_percentage[s] * U_spend[s] * (1 - y[s]), f"RebateLower_{s}"

# 5.9 Custom Constraint:
# For items with Business Unit "A", all award must go to the incumbent supplier.
for j in items:
    if item_attributes[j]['BusinessUnit'] == 'A':
        incumbent_supplier = item_attributes[j]['Incumbent']
        for s in suppliers:
            if s != incumbent_supplier:
                lp_problem += x[(s, j)] == 0, f"IncumbentConstraint_{j}_{s}"

# --------------------------
# 6. Solve the Problem
# --------------------------
lp_problem.solve()

# --------------------------
# 7. Diagnostics (if Infeasible)
# --------------------------
feasibility_notes = ""
status = pulp.LpStatus[lp_problem.status]
print("Status:", status)
print("Total Effective Cost: ${:.2f}".format(pulp.value(lp_problem.objective) if pulp.value(lp_problem.objective) is not None else 0))
print()

if status == "Infeasible":
    feasibility_notes += "Model is infeasible. Possible causes:\n"
    for j in items:
        if item_attributes[j]['BusinessUnit'] == 'A':
            incumbent = item_attributes[j]['Incumbent']
            allowed_capacity = capacity.get((incumbent, j), 0)
            feasibility_notes += (f"  Item {j} (Business Unit A, incumbent {incumbent}): "
                                  f"demand = {demand[j]}, capacity from incumbent = {allowed_capacity}\n")
        else:
            allowed_capacity = sum(capacity[(s, j)] for s in suppliers if (s, j) in capacity)
            feasibility_notes += f"  Item {j}: demand = {demand[j]}, total capacity = {allowed_capacity}\n"
    feasibility_notes += "Please review capacities and/or bidding data for potential issues.\n"
else:
    feasibility_notes = "Model is optimal."

# --------------------------
# 8. Clean and Display Text Output (Optional)
# --------------------------
for s in suppliers:
    print(f"Supplier {s} (Facility: {supplier_attributes[s]['Facility']}):")
    print("  Base Spend (S0): ${:.2f}".format(pulp.value(S0[s])))
    print("  Discount Active (z):", int(pulp.value(z[s])))
    print("  Discount Amount (d): ${:.2f}".format(pulp.value(d[s])))
    print("  Effective Spend (S): ${:.2f}".format(pulp.value(S[s])))
    print("  Total Awarded Volume (V): {:.2f} units".format(pulp.value(V[s])))
    print("  Discount Threshold: {} units, Discount Percentage: {:.2%}".format(
        discount_threshold[s], discount_percentage[s]))
    print("  Rebate Active (y):", int(pulp.value(y[s])))
    print("  Rebate Amount: ${:.2f}".format(pulp.value(rebate[s])))
    print("  Rebate Threshold: {} units, Rebate Percentage: {:.2%}".format(
        rebate_threshold[s], rebate_percentage[s]))
    for j in items:
        print("    Award for {}: {:.2f} units".format(j, pulp.value(x[(s, j)])))
    print()

# --------------------------
# 9. Build Excel Output and Write to Downloads Folder
# --------------------------
# For our Excel output, we create one row per item (bid).
# We determine the awarded supplier (the one with the highest awarded units for the item).
excel_rows = []
for idx, j in enumerate(items, start=1):
    # Determine the awarded supplier (the one with the maximum x[s,j])
    awarded_supplier = None
    max_award = 0
    for s in suppliers:
        award_val = pulp.value(x[(s, j)])
        if award_val is None:
            award_val = 0
        if award_val > max_award:
            max_award = award_val
            awarded_supplier = s
    if awarded_supplier is None:
        awarded_supplier = "None"
        max_award = 0

    # Retrieve the original bid price for the awarded supplier for this item.
    orig_price = price.get((awarded_supplier, j), 0)
    # Check if discount is active for the supplier.
    disc_active = pulp.value(z[awarded_supplier]) if awarded_supplier in z else 0
    if disc_active is None:
        disc_active = 0
    # Apply discount if active.
    if disc_active >= 0.5:
        discount_pct = discount_percentage[awarded_supplier]
    else:
        discount_pct = 0
    discounted_price = orig_price * (1 - discount_pct)

    # Awarded supplier spend for this bid.
    awarded_spend = discounted_price * max_award

    # Baseline values (from external data)
    base_price = baseline_price[j]
    baseline_spend = base_price * max_award
    baseline_savings = baseline_spend - awarded_spend

    # Awarded supplier capacity for this item.
    awarded_capacity = capacity.get((awarded_supplier, j), 0)

    # Allocate supplier-level rebate proportionally (if any).
    total_rebate_supplier = pulp.value(rebate[awarded_supplier]) if awarded_supplier in rebate else 0
    total_volume_supplier = pulp.value(V[awarded_supplier]) if awarded_supplier in V else 0
    if not total_volume_supplier:
        rebate_allocation = 0
    else:
        rebate_allocation = total_rebate_supplier * (max_award / total_volume_supplier)

    # Build the row.
    row = {
        "Bid ID": idx,
        "Bid ID Split": item_attributes[j]['BusinessUnit'],
        "Facility": supplier_attributes.get(awarded_supplier, {}).get('Facility', ""),
        "Incumbent": item_attributes[j]['Incumbent'],
        "Baseline Price": base_price,
        "Baseline Spend": baseline_spend,
        "Awarded Supplier": awarded_supplier,
        "Original Awarded Supplier Price": orig_price,
        "Percentage Volume Discount": f"{discount_percentage[awarded_supplier]*100:.0f}%" if awarded_supplier in discount_percentage else "0%",
        "Discounted Awarded Supplier Price": discounted_price,
        "Awarded Supplier Spend": awarded_spend,
        "Awarded Volume": max_award,
        "Awarded Supplier Capacity": awarded_capacity,
        "Baseline Savings": baseline_savings,
        "Rebate %": f"{rebate_percentage[awarded_supplier]*100:.0f}%" if awarded_supplier in rebate_percentage else "0%",
        "Rebate Savings": rebate_allocation
    }
    excel_rows.append(row)

# Create a DataFrame for the main results.
df_results = pd.DataFrame(excel_rows)

# Order the columns as required.
df_results = df_results[["Bid ID", "Bid ID Split", "Facility", "Incumbent", "Baseline Price", "Baseline Spend",
                           "Awarded Supplier", "Original Awarded Supplier Price", "Percentage Volume Discount",
                           "Discounted Awarded Supplier Price", "Awarded Supplier Spend", "Awarded Volume",
                           "Awarded Supplier Capacity", "Baseline Savings", "Rebate %", "Rebate Savings"]]

# Create a DataFrame for feasibility notes.
df_feasibility = pd.DataFrame({"Feasibility Notes": [feasibility_notes]})

# Determine the user's Downloads folder.
home_dir = os.path.expanduser("~")
downloads_folder = os.path.join(home_dir, "Downloads")
output_file = os.path.join(downloads_folder, "optimization_results.xlsx")

# Write both DataFrames to the Excel file in separate sheets.
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_results.to_excel(writer, sheet_name="Results", index=False)
    df_feasibility.to_excel(writer, sheet_name="Feasibility Notes", index=False)

print(f"Excel results written to {output_file}")

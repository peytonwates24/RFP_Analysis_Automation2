import pulp

# --------------------------
# 1. Problem Data
# --------------------------

# Define suppliers and items.
suppliers = ['A', 'B']
items = ['item1', 'item2']

# Prices (dollars per unit) for each supplier–item combination.
price = {('A', 'item1'): 50,
         ('A', 'item2'): 70,
         ('B', 'item1'): 60,
         ('B', 'item2'): 80}

# Supplier capacities: maximum units that can be supplied for each item.
capacity = {('A', 'item1'): 500,
            ('A', 'item2'): 400,
            ('B', 'item1'): 40000,
            ('B', 'item2'): 30000}

# Demand for each item.
demand = {'item1': 6000,
          'item2': 1000}

# Rebate structure:
# Each supplier offers a rebate if their awarded volume reaches a threshold.
rebate_threshold = {'A': 600,   # Minimum awarded units across all items for a rebate.
                    'B': 500}
rebate_percentage = {'A': 0.20,  # 10% rebate on effective spend.
                     'B': 0.05}  # 5% rebate on effective spend.

# Volume discount structure:
# A supplier offers a discount on their bid price if awarded volume reaches a threshold.
discount_threshold = {'A': 5000,  # If Supplier A is awarded at least 500 units, discount applies.
                      'B': 500}
discount_percentage = {'A': 0.10,  # 1% discount on base spend.
                       'B': 0.01}

# --------------------------
# Big-M Parameters
# --------------------------
# U_volume: Upper bound on awarded volume per supplier = sum of capacities over items.
U_volume = {}
for s in suppliers:
    U_volume[s] = sum(capacity[(s, j)] for j in items)

# U_spend: Upper bound on base spend per supplier.
U_spend = {}
for s in suppliers:
    U_spend[s] = sum(price[(s, j)] * capacity[(s, j)] for j in items)

# A very small epsilon to enforce strict cutoffs.
epsilon = 1e-6

# --------------------------
# 2. Define the MILP Problem
# --------------------------
lp_problem = pulp.LpProblem("Supplier_Sourcing_with_Rebates_and_Discounts", pulp.LpMinimize)

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
# d[s]: Discount amount (if discount is activated).
# S[s]: Effective spend after discount.
S0 = {}
S = {}
for s in suppliers:
    S0[s] = pulp.LpVariable(f"S0_{s}", lowBound=0, cat='Continuous')
    S[s]  = pulp.LpVariable(f"S_{s}", lowBound=0, cat='Continuous')

# V[s]: Total awarded volume (across items) from supplier s.
V = {}
for s in suppliers:
    V[s] = pulp.LpVariable(f"V_{s}", lowBound=0, cat='Continuous')

# Rebate variables:
# y[s]: Binary variable that equals 1 if the awarded volume qualifies for the rebate.
# rebate[s]: Rebate amount.
y = {}
rebate = {}
for s in suppliers:
    y[s] = pulp.LpVariable(f"y_{s}", cat='Binary')
    rebate[s] = pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous')

# Discount variables:
# z[s]: Binary variable that equals 1 if the awarded volume qualifies for the discount.
# d[s]: Discount amount (defined above as S0[s] * discount_percentage if active).
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

# 5.1 Demand Constraints:
# For each item, total awarded units across suppliers must equal the demand.
for j in items:
    lp_problem += (pulp.lpSum(x[(s, j)] for s in suppliers) == demand[j],
                   f"Demand_{j}")

# 5.2 Capacity Constraints:
# For each supplier and item, awarded units cannot exceed the supplier’s capacity.
for s in suppliers:
    for j in items:
        lp_problem += (x[(s, j)] <= capacity[(s, j)], f"Capacity_{s}_{j}")

# 5.3 Compute Base Spend S0[s] for each supplier.
for s in suppliers:
    lp_problem += (S0[s] == pulp.lpSum(price[(s, j)] * x[(s, j)] for j in items),
                   f"BaseSpend_{s}")

# 5.4 Compute Total Awarded Volume V[s] for each supplier.
for s in suppliers:
    lp_problem += (V[s] == pulp.lpSum(x[(s, j)] for j in items),
                   f"Volume_{s}")

# 5.5 Discount Activation Constraints:
# Activate discount if awarded volume V[s] reaches discount_threshold[s].
for s in suppliers:
    lp_problem += (V[s] >= discount_threshold[s] * z[s], f"DiscountThresholdLower_{s}")
    lp_problem += (V[s] <= (discount_threshold[s] - epsilon) + U_volume[s] * z[s],
                   f"DiscountThresholdUpper_{s}")

# 5.6 Discount Linearization Constraints:
# Ensure that d[s] equals discount_percentage[s] * S0[s] if discount is active, else 0.
for s in suppliers:
    lp_problem += (d[s] <= discount_percentage[s] * S0[s], f"DiscountUpper1_{s}")
    lp_problem += (d[s] <= discount_percentage[s] * U_spend[s] * z[s], f"DiscountUpper2_{s}")
    lp_problem += (d[s] >= discount_percentage[s] * S0[s] - discount_percentage[s] * U_spend[s] * (1 - z[s]),
                   f"DiscountLower_{s}")

# 5.7 Effective Spend Calculation:
# Effective Spend S[s] = Base Spend S0[s] minus discount d[s].
for s in suppliers:
    lp_problem += (S[s] == S0[s] - d[s], f"EffectiveSpend_{s}")

# 5.8 Rebate Activation Constraints:
# Activate rebate if awarded volume V[s] reaches rebate_threshold[s].
for s in suppliers:
    lp_problem += (V[s] >= rebate_threshold[s] * y[s], f"RebateThresholdLower_{s}")
    lp_problem += (V[s] <= (rebate_threshold[s] - epsilon) + U_volume[s] * y[s],
                   f"RebateThresholdUpper_{s}")

# 5.9 Rebate Linearization Constraints:
# Ensure that rebate[s] equals rebate_percentage[s] * S[s] if rebate is active, else 0.
for s in suppliers:
    lp_problem += (rebate[s] <= rebate_percentage[s] * S[s], f"RebateUpper1_{s}")
    lp_problem += (rebate[s] <= rebate_percentage[s] * U_spend[s] * y[s], f"RebateUpper2_{s}")
    lp_problem += (rebate[s] >= rebate_percentage[s] * S[s] - rebate_percentage[s] * U_spend[s] * (1 - y[s]),
                   f"RebateLower_{s}")

# --------------------------
# 6. Solve the Problem
# --------------------------
lp_problem.solve()

# --------------------------
# 7. Display the Results
# --------------------------
print("Status:", pulp.LpStatus[lp_problem.status])
print("Total Effective Cost: ${:.2f}".format(pulp.value(lp_problem.objective)))
print()

for s in suppliers:
    # Get raw values from the solver.
    base_spend = pulp.value(S0[s])
    discount_active = pulp.value(z[s])
    discount_amount_raw = pulp.value(d[s])
    effective_spend = pulp.value(S[s])
    total_volume = pulp.value(V[s])
    rebate_active = pulp.value(y[s])
    rebate_amount_raw = pulp.value(rebate[s])
    
    # Clean-up: if the binary flag is 0, then force the amount to 0.
    discount_amount = discount_amount_raw if discount_active >= 0.5 else 0.0
    rebate_amount = rebate_amount_raw if rebate_active >= 0.5 else 0.0

    # Optionally, you can also force the binary display to be 0 or 1.
    discount_active_display = 1 if discount_active >= 0.5 else 0
    rebate_active_display = 1 if rebate_active >= 0.5 else 0

    print(f"Supplier {s}:")
    print("  Base Spend (S0): ${:.2f}".format(base_spend))
    print("  Discount Active (z):", discount_active_display)
    print("  Discount Amount (d): ${:.2f}".format(discount_amount))
    print("  Effective Spend (S): ${:.2f}".format(effective_spend))
    print("  Total Awarded Volume (V): {:.2f} units".format(total_volume))
    print("  Discount Threshold: {} units, Discount Percentage: {:.2%}".format(
        discount_threshold[s], discount_percentage[s]))
    print("  Rebate Active (y):", rebate_active_display)
    print("  Rebate Amount: ${:.2f}".format(rebate_amount))
    print("  Rebate Threshold: {} units, Rebate Percentage: {:.2%}".format(
        rebate_threshold[s], rebate_percentage[s]))
    for j in items:
        print("    Award for {}: {:.2f} units".format(j, pulp.value(x[(s, j)])))
    print()
import pulp

# --------------------------
# 1. Problem Data
# --------------------------

# Sets of suppliers and items
suppliers = ['A', 'B']
items = ['item1', 'item2']

# Prices (in dollars per unit)
price = {('A','item1'): 50,
         ('A','item2'): 70,
         ('B','item1'): 60,
         ('B','item2'): 80}

# Capacities: maximum units that can be procured from a supplier for an item
capacity = {('A','item1'): 500,
            ('A','item2'): 400,
            ('B','item1'): 400,
            ('B','item2'): 300}

# Demand for each item (units)
demand = {'item1': 600,
          'item2': 500}

# Rebate structure: volume threshold (in units) and rebate percentage if threshold is met.
rebate_threshold = {'A': 600,  # Supplier A must get at least 600 units overall
                    'B': 500}  # Supplier B threshold
rebate_percentage = {'A': 0.10,  # 10% rebate
                     'B': 0.05}  # 5% rebate

# Calculate an upper bound on volume per supplier (sum of capacities for that supplier)
U_volume = {}
for s in suppliers:
    U_volume[s] = sum(capacity[(s, j)] for j in items)

# Calculate an upper bound on spend per supplier (maximum possible spend)
U_spend = {}
for s in suppliers:
    U_spend[s] = sum(price[(s, j)] * capacity[(s, j)] for j in items)

# A very small epsilon for rebate activation constraint
epsilon = 1e-6

# --------------------------
# 2. Define the MILP Problem
# --------------------------
lp_problem = pulp.LpProblem("Supplier_Sourcing_with_Rebates", pulp.LpMinimize)

# --------------------------
# 3. Decision Variables
# --------------------------

# x[s, j]: units ordered from supplier s for item j
x = {}
for s in suppliers:
    for j in items:
        x[(s,j)] = pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')

# For each supplier, define variables for total spend and total volume
S = {}   # Total spend from supplier s
V = {}   # Total volume (units) from supplier s
for s in suppliers:
    S[s] = pulp.LpVariable(f"Spend_{s}", lowBound=0, cat='Continuous')
    V[s] = pulp.LpVariable(f"Volume_{s}", lowBound=0, cat='Continuous')

# Binary variable y[s]: 1 if supplier s meets the rebate threshold, 0 otherwise.
y = {}
for s in suppliers:
    y[s] = pulp.LpVariable(f"RebateActive_{s}", cat='Binary')

# Rebate amount for each supplier
rebate = {}
for s in suppliers:
    rebate[s] = pulp.LpVariable(f"Rebate_{s}", lowBound=0, cat='Continuous')

# --------------------------
# 4. Objective Function
# --------------------------
# Minimize the total effective cost: total spend minus rebates
lp_problem += pulp.lpSum(S[s] - rebate[s] for s in suppliers), "Total_Effective_Cost"

# --------------------------
# 5. Constraints
# --------------------------

# 5.1 Demand Constraints: For each item, total procurement must equal demand.
for j in items:
    lp_problem += (pulp.lpSum(x[(s,j)] for s in suppliers) == demand[j], f"Demand_{j}")

# 5.2 Capacity Constraints: For each supplier and item.
for s in suppliers:
    for j in items:
        lp_problem += (x[(s,j)] <= capacity[(s,j)], f"Capacity_{s}_{j}")

# 5.3 Define Spend and Volume for each supplier.
for s in suppliers:
    lp_problem += (S[s] == pulp.lpSum(price[(s,j)] * x[(s,j)] for j in items), f"SpendCalc_{s}")
    lp_problem += (V[s] == pulp.lpSum(x[(s,j)] for j in items), f"VolumeCalc_{s}")

# 5.4 Rebate Activation Constraints: Tie the binary variable to the volume awarded.
for s in suppliers:
    # If y[s]=1, then V[s] must be at least the threshold; if y[s]=0, V[s] is forced below threshold.
    lp_problem += (V[s] >= rebate_threshold[s] * y[s], f"RebateThresholdLower_{s}")
    lp_problem += (V[s] <= (rebate_threshold[s] - epsilon) + U_volume[s] * y[s],
                   f"RebateThresholdUpper_{s}")

# 5.5 Rebate Linearization: Enforce rebate[s] = rebate_percentage[s] * S[s] when y[s]=1; else zero.
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
    print(f"Supplier {s}:")
    print("  Total Spend: ${:.2f}".format(pulp.value(S[s])))
    print("  Total Volume:", pulp.value(V[s]))
    print("  Rebate Active (y):", pulp.value(y[s]))
    print("  Rebate Amount: ${:.2f}".format(pulp.value(rebate[s])))
    for j in items:
        print(f"    Order for {j}: {pulp.value(x[(s,j)])} units")
    print()

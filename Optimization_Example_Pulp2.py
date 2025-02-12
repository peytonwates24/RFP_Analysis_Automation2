import pulp

# --------------------------
# 1. Problem Data
# --------------------------

# Define suppliers and items
suppliers = ['A', 'B']
items = ['item1', 'item2']

# Bid prices (in dollars per unit) for each supplier-item combination.
price = {('A', 'item1'): 50,
         ('A', 'item2'): 70,
         ('B', 'item1'): 60,
         ('B', 'item2'): 10}

# Supplier capacities (the maximum they can supply per item).
# Even if a supplier can bid high (say, 1,000 units), the actual award cannot exceed the demand.
capacity = {('A', 'item1'): 50000,
            ('A', 'item2'): 40000,
            ('B', 'item1'): 40000,
            ('B', 'item2'): 30000}

# Demand for each item.
demand = {'item1': 6000,
          'item2': 500}

# Rebate structure:
# Each supplier offers a rebate if the total awarded volume (across all items) reaches a threshold.
rebate_threshold = {'A': 6100,   # Supplier A must be awarded at least 600 units (across items) for a rebate.
                    'B': 500}   # Supplier B must be awarded at least 500 units for a rebate.
rebate_percentage = {'A': 0.40,  # 10% rebate for Supplier A.
                     'B': 0.05}  # 5% rebate for Supplier B.

# Calculate an upper bound on the total volume a supplier could possibly be awarded.
# This is used only in the rebate activation constraints as a big-M value.
U_volume = {}
for s in suppliers:
    U_volume[s] = sum(capacity[(s, j)] for j in items)

# Calculate an upper bound on spend per supplier (again, a big-M value for linearization).
U_spend = {}
for s in suppliers:
    U_spend[s] = sum(price[(s, j)] * capacity[(s, j)] for j in items)

# A very small epsilon to enforce a strict cutoff in the big-M constraints.
epsilon = 1e-6

# --------------------------
# 2. Define the MILP Problem
# --------------------------
lp_problem = pulp.LpProblem("Supplier_Sourcing_with_Rebates", pulp.LpMinimize)

# --------------------------
# 3. Decision Variables
# --------------------------

# x[s, j]: Number of units awarded from supplier s for item j.
x = {}
for s in suppliers:
    for j in items:
        x[(s, j)] = pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')

# For each supplier, define:
# S[s]: Total spend (cost) incurred from supplier s.
# V[s]: Total awarded volume (across all items) from supplier s.
S = {}
V = {}
for s in suppliers:
    S[s] = pulp.LpVariable(f"Spend_{s}", lowBound=0, cat='Continuous')
    V[s] = pulp.LpVariable(f"Volume_{s}", lowBound=0, cat='Continuous')

# Binary variable y[s]: equals 1 if supplier s’s awarded volume meets its rebate threshold; 0 otherwise.
y = {}
for s in suppliers:
    y[s] = pulp.LpVariable(f"RebateActive_{s}", cat='Binary')

# Rebate amount for each supplier.
rebate = {}
for s in suppliers:
    rebate[s] = pulp.LpVariable(f"Rebate_{s}", lowBound=0, cat='Continuous')

# --------------------------
# 4. Objective Function
# --------------------------
# Minimize the total effective cost: total spend minus rebates.
lp_problem += pulp.lpSum(S[s] - rebate[s] for s in suppliers), "Total_Effective_Cost"

# --------------------------
# 5. Constraints
# --------------------------

# 5.1 Demand Constraints:
# Ensure that for each item, the sum of awards from all suppliers equals the demand.
for j in items:
    lp_problem += (pulp.lpSum(x[(s, j)] for s in suppliers) == demand[j], f"Demand_{j}")

# 5.2 Independent Capacity Constraints:
# For each supplier and item, ensure that the awarded units do not exceed the supplier's capacity.
for s in suppliers:
    for j in items:
        lp_problem += (x[(s, j)] <= capacity[(s, j)], f"Capacity_{s}_{j}")

# 5.3 Define Total Spend and Total Awarded Volume for each supplier.
for s in suppliers:
    lp_problem += (S[s] == pulp.lpSum(price[(s, j)] * x[(s, j)] for j in items), f"SpendCalc_{s}")
    lp_problem += (V[s] == pulp.lpSum(x[(s, j)] for j in items), f"VolumeCalc_{s}")

# 5.4 Rebate Activation Constraints:
# The binary variable y[s] is set to 1 if the awarded volume V[s] reaches the supplier's threshold.
for s in suppliers:
    # When y[s] is 1, V[s] must be at least the threshold.
    lp_problem += (V[s] >= rebate_threshold[s] * y[s], f"RebateThresholdLower_{s}")
    # When y[s] is 0, V[s] is forced to be below the threshold, using U_volume[s] as a big-M constant.
    lp_problem += (V[s] <= (rebate_threshold[s] - epsilon) + U_volume[s] * y[s], f"RebateThresholdUpper_{s}")

# 5.5 Rebate Linearization Constraints:
# These constraints ensure that the rebate is:
# - Equal to rebate_percentage[s] * S[s] when y[s] = 1.
# - Zero when y[s] = 0.
for s in suppliers:
    lp_problem += (rebate[s] <= rebate_percentage[s] * S[s], f"RebateUpper1_{s}")
    lp_problem += (rebate[s] <= rebate_percentage[s] * U_spend[s] * y[s], f"RebateUpper2_{s}")
    lp_problem += (rebate[s] >= rebate_percentage[s] * S[s] - rebate_percentage[s] * U_spend[s] * (1 - y[s]), f"RebateLower_{s}")

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
    spend_value = pulp.value(S[s])
    volume_value = pulp.value(V[s])
    rebate_value = pulp.value(rebate[s])
    rebate_active = pulp.value(y[s])
    
    # If the rebate is activated, the effective lower bound on the rebate is rebate_percentage * Spend.
    rebate_lower_bound = rebate_percentage[s] * spend_value if rebate_active >= 0.99 else 0
    
    print(f"Supplier {s}:")
    print("  Total Spend: ${:.2f}".format(spend_value))
    print("  Total Awarded Volume: {} units".format(volume_value))
    print("  Rebate Active (y):", int(rebate_active))
    print("  Rebate Amount: ${:.2f}".format(rebate_value))
    print("  Rebate Threshold (Volume Lower Bound): {} units".format(rebate_threshold[s]))
    print("  Rebate Percentage: {:.2%}".format(rebate_percentage[s]))
    for j in items:
        print("    Award for {}: {} units".format(j, pulp.value(x[(s, j)])))
    print()

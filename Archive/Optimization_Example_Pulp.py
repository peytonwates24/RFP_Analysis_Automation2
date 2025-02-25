# Import PuLP library
import pulp

# 1. Define the Problem
# Create a Linear Programming problem instance.
# "Supplier_Optimization" is the name of the problem.
# pulp.LpMinimize indicates that this is a minimization problem.
lp_problem = pulp.LpProblem("Supplier_Optimization", pulp.LpMinimize)

# 2. Define the Decision Variables
# Define two decision variables:
# x_A: Units to order from Supplier A
# x_B: Units to order from Supplier B
# Both are continuous and must be non-negative.
x_A = pulp.LpVariable("Quantity_from_Supplier_A", lowBound=0, cat='Continuous')
x_B = pulp.LpVariable("Quantity_from_Supplier_B", lowBound=0, cat='Continuous')

# Total demand is 1000 units.
total_demand = 1000

# 3. Define the Objective Function
# Our goal is to minimize the total cost:
# Total Cost = 50 * (units from A) + 60 * (units from B)
lp_problem += 50 * x_A + 60 * x_B, "Total_Cost"

# 4. Add Constraints

# 4.1 Demand Constraint: The sum of orders must meet the total demand.
lp_problem += x_A + x_B == total_demand, "Demand_Constraint"

# 4.2 Supplier A Capacity: Cannot order more than 800 units from Supplier A.
lp_problem += x_A <= 800, "Supplier_A_Capacity"

# 4.3 Supplier B Capacity: Cannot order more than 600 units from Supplier B.
lp_problem += x_B <= 600, "Supplier_B_Capacity"

# 4.4 Diversification Constraint:
# At least 30% of the total order must come from Supplier B.
lp_problem += x_B >= 0.3 * total_demand, "Supplier_B_Minimum"

# 5. Solve the Problem
# Use PuLP's default solver to solve the problem.
lp_problem.solve()

# 6. Print the Results
# Print the status to check if an optimal solution was found.
print("Status:", pulp.LpStatus[lp_problem.status])

# Print the optimal (minimum) total cost.
print("Minimum Total Cost: $", pulp.value(lp_problem.objective))

# Print the optimal order quantities for each supplier.
print("Order from Supplier A:", pulp.value(x_A), "units")
print("Order from Supplier B:", pulp.value(x_B), "units")

import os
import uuid
import pandas as pd
import PySimpleGUI as sg
import pulp

#############################################
# DEFAULT DATA
#############################################

suppliers = ['A', 'B', 'C']

# Item Attributes (keyed by Bid ID)
default_item_attributes = {
    '1': {'BusinessUnit': 'A', 'Incumbent': 'A', 'Capacity Group': 'Widgets', 'Facility': 'Facility1'},
    '2': {'BusinessUnit': 'B', 'Incumbent': 'B', 'Capacity Group': 'Gadgets', 'Facility': 'Facility1'},
    '3': {'BusinessUnit': 'A', 'Incumbent': 'C', 'Capacity Group': 'Widgets', 'Facility': 'Facility1'},
    '4': {'BusinessUnit': 'A', 'Incumbent': 'C', 'Capacity Group': 'Widgets', 'Facility': 'Facility2'},
    '5': {'BusinessUnit': 'A', 'Incumbent': 'C', 'Capacity Group': 'Gadgets', 'Facility': 'Facility2'},
    '6': {'BusinessUnit': 'B', 'Incumbent': 'C', 'Capacity Group': 'Gadgets', 'Facility': 'Facility2'},
    '7': {'BusinessUnit': 'B', 'Incumbent': 'C', 'Capacity Group': 'Gadgets', 'Facility': 'Facility2'},
    '8': {'BusinessUnit': 'B', 'Incumbent': 'C', 'Capacity Group': 'Widgets', 'Facility': 'Facility3'},
    '9': {'BusinessUnit': 'A', 'Incumbent': 'C', 'Capacity Group': 'Widgets', 'Facility': 'Facility4'},
    '10': {'BusinessUnit': 'A', 'Incumbent': 'C', 'Capacity Group': 'Widgets', 'Facility': 'Facility5'}
}

default_price = {
    ('A', '1'): 50,  ('A', '2'): 70,  ('A', '3'): 55,
    ('B', '1'): 6,  ('B', '2'): 80,  ('B', '3'): 65,
    ('C', '1'): 55,  ('C', '2'): 75,  ('C', '3'): 60,
    ('A', '4'): 23,  ('A', '5'): 54,  ('A', '6'): 42,
    ('B', '4'): 75,  ('B', '5'): 34,  ('B', '6'): 24,
    ('C', '4'): 24,  ('C', '5'): 2,  ('C', '6'): 64,
    ('A', '7'): 232, ('A', '8'): 75,  ('A', '9'): 97,
    ('B', '7'): 53,  ('B', '8'): 13,  ('B', '9'): 56,
    ('C', '7'): 86,  ('C', '8'): 24,  ('C', '9'): 134,
    ('A', '10'): 64, ('B', '10'): 1300, ('C', '10'): 7500
}

default_demand = {
    '1': 700, '2': 9000, '3': 600, '4': 5670, '5': 45,
    '6': 242, '7': 664, '8': 24, '9': 232, '10': 13
}

default_rebate_tiers = {
    'A': [(0, 500, 0.0, None, None), (500, float('inf'), 0.10, None, None)],
    'B': [(0, 500, 0.0, None, None), (500, float('inf'), 0.05, "Capacity Group", "Gadgets")],
    'C': [(0, 700, 0.0, None, None), (700, float('inf'), 0.08, "Capacity Group", "Widgets")]
}

default_discount_tiers = {
    'A': [(0, 1000, 0.0, None, None), (1000, float('inf'), 0.01, "Capacity Group", "Widgets")],
    'B': [(0, 500, 0.0, None, None), (500, float('inf'), 0.03, "Capacity Group", "Gadgets")],
    'C': [(0, 500, 0.0, None, None), (500, float('inf'), 0.04, "Capacity Group", "Widgets")]
}

default_baseline_price = {
    '1': 100, '2': 156, '3': 423, '4': 453, '5': 342,
    '6': 653, '7': 432, '8': 456, '9': 234, '10': 231
}

default_per_item_capacity = {
    ('A', '1'): 5000, ('A', '2'): 4000, ('A', '3'): 3000,
    ('B', '1'): 4000, ('B', '2'): 8000, ('B', '3'): 6000,
    ('C', '1'): 3000, ('C', '2'): 5000, ('C', '3'): 7000
}

default_global_capacity_df = pd.DataFrame({
    "Supplier Name": ["A", "A", "B", "B", "C", "C"],
    "Capacity Group": ["Widgets", "Gadgets", "Widgets", "Gadgets", "Widgets", "Gadgets"],
    "Capacity": [100000, 90000, 12000, 11000, 150000, 300000]
})
default_global_capacity = {
    (str(row["Supplier Name"]).strip(), str(row["Capacity Group"]).strip()): row["Capacity"]
    for idx, row in default_global_capacity_df.iterrows()
}

def compute_U_volume(per_item_cap):
    total = {}
    for s in suppliers:
        tot = sum(per_item_cap.get((s, j), 0) for (sup, j) in per_item_cap.keys() if sup == s)
        total[s] = tot
    return total

default_U_volume = compute_U_volume(default_per_item_capacity)

def compute_U_spend(per_item_cap):
    total = {}
    for s in suppliers:
        tot = sum(default_price.get((s, j), 0) * default_per_item_capacity.get((s, j), 0)
                  for (sup, j) in default_per_item_capacity.keys() if sup == s)
        total[s] = tot
    return total

default_U_spend = compute_U_spend(default_per_item_capacity)

epsilon = 1e-6
small_value = 1e-3
M = 1e9

#############################################
# DEBUG FLAG (global)
#############################################
debug = True

#############################################
# NEW: SUPPLIER BID ATTRIBUTES (keyed by (supplier, bid))
#############################################
default_supplier_bid_attributes = {
    ('A', '1'): {"Milage": 400, "Origin Country": "Mexico"},
    ('A', '2'): {"Milage": 420, "Origin Country": "USA"},
    ('A', '3'): {"Milage": 410, "Origin Country": "USA"},
    ('A', '4'): {"Milage": 430, "Origin Country": "USA"},
    ('A', '5'): {"Milage": 450, "Origin Country": "Canada"},
    ('A', '6'): {"Milage": 460, "Origin Country": "USA"},
    ('A', '7'): {"Milage": 470, "Origin Country": "Canada"},
    ('A', '8'): {"Milage": 480, "Origin Country": "USA"},
    ('A', '9'): {"Milage": 490, "Origin Country": "USA"},
    ('A', '10'): {"Milage": 500, "Origin Country": "Mexico"},

    ('B', '1'): {"Milage": 499, "Origin Country": "Canada"},
    ('B', '2'): {"Milage": 610, "Origin Country": "Canada"},
    ('B', '3'): {"Milage": 620, "Origin Country": "Canada"},
    ('B', '4'): {"Milage": 630, "Origin Country": "Mexico"},
    ('B', '5'): {"Milage": 640, "Origin Country": "Canada"},
    ('B', '6'): {"Milage": 650, "Origin Country": "Canada"},
    ('B', '7'): {"Milage": 660, "Origin Country": "Canada"},
    ('B', '8'): {"Milage": 670, "Origin Country": "Mexico"},
    ('B', '9'): {"Milage": 680, "Origin Country": "Canada"},
    ('B', '10'): {"Milage": 690, "Origin Country": "Canada"},

    ('C', '1'): {"Milage": 500, "Origin Country": "USA"},
    ('C', '2'): {"Milage": 510, "Origin Country": "USA"},
    ('C', '3'): {"Milage": 520, "Origin Country": "USA"},
    ('C', '4'): {"Milage": 530, "Origin Country": "USA"},
    ('C', '5'): {"Milage": 540, "Origin Country": "Mexico"},
    ('C', '6'): {"Milage": 550, "Origin Country": "USA"},
    ('C', '7'): {"Milage": 560, "Origin Country": "USA"},
    ('C', '8'): {"Milage": 570, "Origin Country": "Mexico"},
    ('C', '9'): {"Milage": 580, "Origin Country": "Mexico"},
    ('C', '10'): {"Milage": 590, "Origin Country": "USA"}
}
default_bid_grouping_options = ["Milage", "Origin Country"]

#############################################
# UPDATE GROUPING SCOPE FUNCTION (for item attributes)
#############################################
def update_grouping_scope(grouping, item_attr_data):
    if grouping == "Bid ID":
        vals = sorted(list(item_attr_data.keys()))
    else:
        vals = sorted({str(item_attr_data[j].get(grouping, "")).strip() 
                       for j in item_attr_data if str(item_attr_data[j].get(grouping, "")).strip() != ""})
    return ["Apply to all items individually"] + vals

#############################################
# OPTIMIZATION MODEL FUNCTION
#############################################
def run_optimization(use_global, capacity_data, demand_data, item_attr_data, price_data,
                     rebate_tiers, discount_tiers, baseline_price_data, rules=[]):
    global debug
    items_dynamic = list(demand_data.keys())
    # Create transition variables for non-incumbent suppliers.
    T = {}
    for j in items_dynamic:
        incumbent = item_attr_data[j].get("Incumbent")
        for s in suppliers:
            if s != incumbent:
                T[(j, s)] = pulp.LpVariable(f"T_{j}_{s}", cat='Binary')
    U_volume = default_U_volume
    U_spend = default_U_spend

    lp_problem = pulp.LpProblem("Sourcing_with_MultiTier_Rebates_Discounts", pulp.LpMinimize)

    # Decision variables.
    x = {(s, j): pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')
         for s in suppliers for j in items_dynamic}
    S0 = {s: pulp.LpVariable(f"S0_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    S = {s: pulp.LpVariable(f"S_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    V = {s: pulp.LpVariable(f"V_{s}", lowBound=0, cat='Continuous') for s in suppliers}

    # Transition constraints.
    for j in items_dynamic:
        for s in suppliers:
            if (j, s) in T:
                lp_problem += x[(s, j)] <= demand_data[j] * T[(j, s)], f"Transition_{j}_{s}"

    # Discount tier constraints.
    z_discount = {}
    for s in suppliers:
        tiers = discount_tiers.get(s, default_discount_tiers[s])
        z_discount[s] = {k: pulp.LpVariable(f"z_discount_{s}_{k}", cat='Binary') for k in range(len(tiers))}
        lp_problem += pulp.lpSum(z_discount[s][k] for k in range(len(tiers))) == 1, f"DiscountTierSelect_{s}"
    # Rebate tier constraints.
    y_rebate = {}
    for s in suppliers:
        tiers = rebate_tiers.get(s, default_rebate_tiers[s])
        y_rebate[s] = {k: pulp.LpVariable(f"y_rebate_{s}_{k}", cat='Binary') for k in range(len(tiers))}
        lp_problem += pulp.lpSum(y_rebate[s][k] for k in range(len(tiers))) == 1, f"RebateTierSelect_{s}"

    d = {s: pulp.LpVariable(f"d_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    rebate_var = {s: pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous') for s in suppliers}

    lp_problem += pulp.lpSum(S[s] - rebate_var[s] for s in suppliers), "Total_Effective_Cost"

    for j in items_dynamic:
        lp_problem += pulp.lpSum(x[(s, j)] for s in suppliers) == demand_data[j], f"Demand_{j}"

    if use_global:
        supplier_capacity_groups = {}
        all_groups = {item_attr_data[j].get("Capacity Group") for j in items_dynamic if item_attr_data[j].get("Capacity Group") is not None}
        for s in suppliers:
            supplier_capacity_groups[s] = {g: [] for g in all_groups}
            for j in items_dynamic:
                group = item_attr_data[j].get("Capacity Group")
                if group is not None:
                    supplier_capacity_groups[s][group].append(j)
        for s in suppliers:
            for group, item_list in supplier_capacity_groups[s].items():
                cap = capacity_data.get((s, group), 1e9)
                lp_problem += pulp.lpSum(x[(s, j)] for j in item_list) <= cap, f"GlobalCapacity_{s}_{group}"
    else:
        for s in suppliers:
            for j in items_dynamic:
                cap = capacity_data.get((s, j), 1e9)
                lp_problem += x[(s, j)] <= cap, f"PerItemCapacity_{s}_{j}"

    for s in suppliers:
        lp_problem += S0[s] == pulp.lpSum(price_data[(s, j)] * x[(s, j)] for j in items_dynamic), f"BaseSpend_{s}"
        lp_problem += V[s] == pulp.lpSum(x[(s, j)] for j in items_dynamic), f"Volume_{s}"

    for s in suppliers:
        tiers = discount_tiers.get(s, default_discount_tiers[s])
        M_discount = U_spend[s]
        for k, tier in enumerate(tiers):
            Dmin, Dmax, Dperc, scope_attr, scope_value = tier
            if scope_attr is None or scope_value is None:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic)
            else:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic if item_attr_data[j].get(scope_attr) == scope_value)
            lp_problem += vol_expr >= Dmin * z_discount[s][k], f"DiscountTierMin_{s}_{k}"
            if Dmax < float('inf'):
                lp_problem += vol_expr <= Dmax + M_discount*(1 - z_discount[s][k]), f"DiscountTierMax_{s}_{k}"
            lp_problem += d[s] >= Dperc * S0[s] - M_discount*(1 - z_discount[s][k]), f"DiscountTierLower_{s}_{k}"
            lp_problem += d[s] <= Dperc * S0[s] + M_discount*(1 - z_discount[s][k]), f"DiscountTierUpper_{s}_{k}"

    for s in suppliers:
        lp_problem += S[s] == S0[s] - d[s], f"EffectiveSpend_{s}"

    for s in suppliers:
        tiers = rebate_tiers.get(s, default_rebate_tiers[s])
        M_rebate = U_spend[s]
        for k, tier in enumerate(tiers):
            Rmin, Rmax, Rperc, scope_attr, scope_value = tier
            if scope_attr is None or scope_value is None:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic)
            else:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic if item_attr_data[j].get(scope_attr) == scope_value)
            lp_problem += vol_expr >= Rmin * y_rebate[s][k], f"RebateTierMin_{s}_{k}"
            if Rmax < float('inf'):
                lp_problem += vol_expr <= Rmax + M_rebate*(1 - y_rebate[s][k]), f"RebateTierMax_{s}_{k}"
            lp_problem += rebate_var[s] >= Rperc * S[s] - M_rebate*(1 - y_rebate[s][k]), f"RebateTierLower_{s}_{k}"
            lp_problem += rebate_var[s] <= Rperc * S[s] + M_rebate*(1 - y_rebate[s][k]), f"RebateTierUpper_{s}_{k}"

    lowest_cost_supplier = {}
    second_lowest_cost_supplier = {}
    for j in items_dynamic:
        prices = []
        for s in suppliers:
            if (s, j) in price_data:
                prices.append((price_data[(s, j)], s))
        if prices:
            prices.sort(key=lambda x: x[0])
            lowest_cost_supplier[j] = prices[0][1]
            second_lowest_cost_supplier[j] = prices[1][1] if len(prices) > 1 else prices[0][1]

    #############################################
    # CUSTOM RULES PROCESSING
    #############################################
    for r_idx, rule in enumerate(rules):
        if rule["rule_type"] == "# of suppliers":
            if rule["grouping"] == "Bid ID" and rule["operator"] == "Exactly" and rule["rule_input"] == "1":
                if rule["grouping_scope"] == "Apply to all items individually":
                    bids = sorted(list(item_attr_data.keys()))
                else:
                    bids = [rule["grouping_scope"]]
                for j in bids:
                    w = {}
                    for s in suppliers:
                        w[(s, j)] = pulp.LpVariable(f"w_{r_idx}_{j}_{s}", cat='Binary')
                        lp_problem += x[(s, j)] <= M * w[(s, j)], f"RuleSupplierIndicator_{r_idx}_{j}_{s}"
                        lp_problem += x[(s, j)] >= small_value * w[(s, j)], f"RuleSupplierIndicatorLB_{r_idx}_{j}_{s}"
                    lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) == 1, f"RuleSingleSupplier_{r_idx}_{j}"
                    if debug:
                        print(f"DEBUG: Enforcing exactly one supplier for Bid {j} via rule {r_idx}")
            else:
                if rule["grouping"] == "All" or rule["grouping_scope"] == "All":
                    items_group = items_dynamic
                elif rule["grouping_scope"] == "Apply to all items individually":
                    if rule["grouping"] == "Bid ID":
                        items_group = sorted(list(item_attr_data.keys()))
                    else:
                        items_group = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip()
                                               for j in item_attr_data if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""})
                else:
                    items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
                try:
                    supplier_target = int(rule["rule_input"])
                except:
                    continue
                operator = rule["operator"]
                w = {}
                for s in suppliers:
                    w[s] = pulp.LpVariable(f"w_{r_idx}_{s}", cat='Binary')
                    lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) <= M * w[s], f"SupplierIndicator_{r_idx}_{s}"
                    lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) >= small_value * w[s], f"SupplierIndicatorLB_{r_idx}_{s}"
                total_suppliers = pulp.lpSum(w[s] for s in suppliers)
                if operator == "At least":
                    lp_problem += total_suppliers >= supplier_target, f"Rule_{r_idx}"
                elif operator == "At most":
                    lp_problem += total_suppliers <= supplier_target, f"Rule_{r_idx}"
                elif operator == "Exactly":
                    lp_problem += total_suppliers == supplier_target, f"Rule_{r_idx}"
        elif rule["rule_type"] == "% of Volume Awarded":
            if rule["supplier_scope"] == "New Suppliers" and rule["grouping"] == "All":
                try:
                    percentage = float(rule["rule_input"]) / 100.0
                except:
                    continue
                total_new_suppliers_vol = pulp.lpSum(
                    pulp.lpSum(x[(s, j)] for s in suppliers if s != item_attr_data[j].get("Incumbent"))
                    for j in items_dynamic
                )
                lp_problem += total_new_suppliers_vol <= percentage * sum(demand_data[j] for j in items_dynamic), f"Rule_{r_idx}_AggregateNewSuppliers"
                if debug:
                    print(f"DEBUG: Enforcing aggregate new suppliers volume <= {percentage * sum(demand_data[j] for j in items_dynamic):.2f}")
            else:
                if rule["grouping"] == "All" or rule["grouping_scope"] == "All":
                    items_group = items_dynamic
                elif rule["grouping_scope"] == "Apply to all items individually":
                    if rule["grouping"] == "Bid ID":
                        items_group = sorted(list(item_attr_data.keys()))
                    else:
                        items_group = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip()
                                               for j in item_attr_data if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""})
                else:
                    items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
                try:
                    percentage = float(rule["rule_input"]) / 100.0
                except:
                    continue
                operator = rule["operator"]
                if rule["supplier_scope"] in ["Lowest cost supplier", "Second Lowest Cost Supplier", "Incumbent"]:
                    if rule["supplier_scope"] == "Lowest cost supplier":
                        lhs = pulp.lpSum(x[(lowest_cost_supplier[j], j)] for j in items_group)
                    elif rule["supplier_scope"] == "Second Lowest Cost Supplier":
                        lhs = pulp.lpSum(x[(second_lowest_cost_supplier[j], j)] for j in items_group)
                    elif rule["supplier_scope"] == "Incumbent":
                        lhs = pulp.lpSum(x[(item_attr_data[j].get("Incumbent"), j)] for j in items_group)
                else:
                    lhs = pulp.lpSum(x[(rule["supplier_scope"], j)] for j in items_group)
                total_vol = pulp.lpSum(x[(s, j)] for s in suppliers for j in items_group)
                if operator == "At least":
                    lp_problem += lhs >= percentage * total_vol, f"Rule_{r_idx}"
                elif operator == "At most":
                    lp_problem += lhs <= percentage * total_vol, f"Rule_{r_idx}"
                elif operator == "Exactly":
                    lp_problem += lhs == percentage * total_vol, f"Rule_{r_idx}"
        elif rule["rule_type"] == "# of transitions":
            if rule["grouping"] == "All" or rule["grouping_scope"] == "All":
                items_group = items_dynamic
            elif rule["grouping_scope"] == "Apply to all items individually":
                if rule["grouping"] == "Bid ID":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip()
                                           for j in item_attr_data if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""})
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            try:
                transitions_target = int(rule["rule_input"])
            except:
                continue
            operator = rule["operator"]
            total_transitions = pulp.lpSum(T[(j, s)] for (j, s) in T if j in items_group)
            if operator == "At least":
                lp_problem += total_transitions >= transitions_target, f"Rule_{r_idx}"
            elif operator == "At most":
                lp_problem += total_transitions <= transitions_target, f"Rule_{r_idx}"
            elif operator == "Exactly":
                lp_problem += total_transitions == transitions_target, f"Rule_{r_idx}"
            if debug:
                print(f"DEBUG: Enforcing total transitions {operator} {transitions_target} over items {items_group}")
        elif rule["rule_type"] == "Bid Exclusions":
            # Process Bid Exclusions rule.
            if rule["grouping"] == "All" or rule["grouping_scope"] in ["All", "Apply to all items individually"]:
                items_group = items_dynamic
            else:
                items_group = [j for j in items_dynamic if str(default_item_attributes[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            bid_group = rule.get("bid_grouping", None)
            if bid_group is None:
                continue
            # Determine if the bid attribute is numeric using the first available value.
            is_numeric = True
            for j in items_group:
                for s in suppliers:
                    val = default_supplier_bid_attributes.get((s, j), {}).get(bid_group, None)
                    if val is not None:
                        try:
                            float(val)
                        except:
                            is_numeric = False
                        break
                break
            # For each bid and each supplier, check the bid attribute value.
            for j in items_group:
                for s in suppliers:
                    bid_val = default_supplier_bid_attributes.get((s, j), {}).get(bid_group, None)
                    if bid_val is None:
                        continue
                    exclude = False
                    if is_numeric:
                        try:
                            bid_val_num = float(bid_val)
                            threshold = float(rule["rule_input"])
                        except:
                            continue
                        op = rule["operator"]
                        if op == "At most" and bid_val_num > threshold:
                            exclude = True
                        elif op == "At least" and bid_val_num < threshold:
                            exclude = True
                        elif op == "Exactly" and bid_val_num != threshold:
                            exclude = True
                    else:
                        if bid_val.strip() == rule.get("bid_exclusion_value", "").strip():
                            exclude = True
                    if exclude:
                        lp_problem += x[(s, j)] == 0, f"BidExclusion_{r_idx}_{j}_{s}"
                        if debug:
                            print(f"DEBUG: Excluding Bid {j} for supplier {s} due to Bid Exclusions rule {r_idx} on {bid_group} with value {bid_val}")
    if debug:
        constraint_names = list(lp_problem.constraints.keys())
        duplicates = set([n for n in constraint_names if constraint_names.count(n) > 1])
        if duplicates:
            print("DEBUG: Duplicate constraint names found:", duplicates)
        else:
            print("DEBUG: No duplicate constraint names.")
        print("DEBUG: Total constraints added:", len(constraint_names))
    
    lp_problem.solve()
    model_status = pulp.LpStatus[lp_problem.status]
    
    feasibility_notes = ""
    if model_status == "Infeasible":
        feasibility_notes += "Model is infeasible. Likely causes include:\n"
        feasibility_notes += " - Insufficient supplier capacity relative to demand.\n"
        feasibility_notes += " - Custom rule constraints conflicting with overall volume/demand.\n"
        for j in items_dynamic:
            if use_global:
                group = default_item_attributes[j].get("Capacity Group", "Unknown")
                total_cap = sum(capacity_data.get((s, group), 0) for s in suppliers)
                feasibility_notes += f"  Bid {j}: demand = {demand_data[j]}, capacity for group {group} = {total_cap}\n"
            else:
                total_cap = sum(capacity_data.get((s, j), 0) for s in suppliers)
                feasibility_notes += f"  Bid {j}: demand = {demand_data[j]}, capacity = {total_cap}\n"
        feasibility_notes += "Please review supplier capacities, demand, and custom rule constraints."
    else:
        feasibility_notes = "Model is optimal."
    
    excel_rows = []
    letter_list = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    for idx, j in enumerate(items_dynamic, start=1):
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
            active_discount = 0
            for k, tier in enumerate(discount_tiers.get(s, default_discount_tiers[s])):
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
                group = default_item_attributes[j].get("Capacity Group", "Unknown")
                awarded_capacity = capacity_data.get((s, group), 0)
            else:
                awarded_capacity = capacity_data.get((s, j), 0)
            active_rebate = 0
            for k, tier in enumerate(rebate_tiers.get(s, default_rebate_tiers[s])):
                if pulp.value(y_rebate[s][k]) is not None and pulp.value(y_rebate[s][k]) >= 0.5:
                    active_rebate = tier[2]
                    break
            rebate_savings = awarded_spend * active_rebate
            facility_val = default_item_attributes[j].get("Facility", "")
            row = {
                "Bid ID": idx,
                "Capacity Group": default_item_attributes[j].get("Capacity Group", "") if use_global else "",
                "Bid ID Split": bid_split,
                "Facility": facility_val,
                "Incumbent": default_item_attributes[j].get("Incumbent", ""),
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
    temp_lp_file = os.path.join(downloads_folder, "temp_model.lp")
    lp_problem.writeLP(temp_lp_file)
    with open(temp_lp_file, "r") as f:
        lp_text = f.read()
    df_lp = pd.DataFrame({"LP Model": [lp_text]})
    
    output_file = os.path.join(downloads_folder, "optimization_results.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_results.to_excel(writer, sheet_name="Results", index=False)
        df_feasibility.to_excel(writer, sheet_name="Feasibility Notes", index=False)
        df_lp.to_excel(writer, sheet_name="LP Model", index=False)
    
    return output_file, feasibility_notes, model_status

#############################################
# GUI LAYOUT AND EVENT LOOP
#############################################
default_grouping_options = ["All", "Bid ID"] + list(default_item_attributes[next(iter(default_item_attributes))].keys())
default_supplier_scope_options = suppliers + ["New Suppliers", "Lowest cost supplier", "Second Lowest Cost Supplier", "Incumbent"]

layout = [
    [sg.Text("Select Capacity Input Type:")],
    [sg.Radio("Global Capacity", "CAP_TYPE", key="-GLOBAL-", default=True),
     sg.Radio("Per Item Capacity", "CAP_TYPE", key="-PERITEM-")],
    [sg.Frame("Custom Rules", [
         [sg.Text("Rule Type:"), sg.Combo(["% of Volume Awarded", "# of Volume Awarded", "# of transitions", "# of suppliers", "Supplier Exclusion", "Bid Exclusions"],
                                             key="-RULETYPE-", size=(30, 1), enable_events=True)],
         [sg.Text("Operator:"), sg.Combo(["At least", "At most", "Exactly"], key="-RULEOP-", size=(30, 1), default_value="At least")],
         [sg.Text("Rule Input:"), sg.Input(key="-RULEINPUT-", size=(10, 1))],
         [sg.Text("Grouping:"), sg.Combo(values=default_grouping_options, key="-GROUPING-", size=(30, 1), enable_events=True, default_value=default_grouping_options[0])],
         [sg.Text("Apply to all items individually:"), sg.Combo(values=[], key="-GROUPSCOPE-", size=(30, 1))],
         [sg.Text("Supplier Scope:"), sg.Combo(values=default_supplier_scope_options, key="-SUPPSCOPE-", size=(30, 1))],
         # New fields for Bid Exclusions:
         [sg.Text("Bid Grouping:"), sg.Combo(values=default_bid_grouping_options, key="-BIDGROUPING-", size=(30, 1), enable_events=True, visible=False)],
         [sg.Text("Bid Exclusion Value:"), sg.Combo(values=[], key="-BIDEXCLUSIONVALUE-", size=(30, 1), visible=False)],
         [sg.Button("Add Rule", key="-ADDRULE-"), sg.Button("Clear Rules", key="-CLEARRULES-")],
         [sg.Multiline("", size=(60, 5), key="-RULELIST-")]
    ])],
    [sg.Button("Run Optimization")],
    [sg.Button("Exit")]
]

window = sg.Window("Sourcing Optimization", layout)
rules_list = []

def rule_to_text(rule):
    if rule["rule_type"] == "% of Volume Awarded":
        supplier_text = f"{rule['supplier_scope']}"
        grouping_text = "all items" if rule["grouping"] == "All" else rule["grouping_scope"]
        return f"% Vol: {rule['operator']} {rule['rule_input']}% of {grouping_text} awarded to {supplier_text}"
    elif rule["rule_type"] == "# of Volume Awarded":
        supplier_text = f"{rule['supplier_scope']}"
        grouping_text = "all items" if rule["grouping"] == "All" else rule["grouping_scope"]
        return f"# Vol: {rule['operator']} {rule['rule_input']} units of {grouping_text} awarded to {supplier_text}"
    elif rule["rule_type"] == "# of transitions":
        grouping_text = "all items" if rule["grouping"] == "All" else rule["grouping_scope"]
        return f"# Transitions: {rule['operator']} {rule['rule_input']} transitions in {grouping_text}"
    elif rule["rule_type"] == "# of suppliers":
        grouping_text = "all items" if rule["grouping"] == "All" else rule["grouping_scope"]
        return f"# Suppliers: {rule['operator']} {rule['rule_input']} unique suppliers in {grouping_text}"
    elif rule["rule_type"] == "Supplier Exclusion":
        grouping_text = "all items" if rule["grouping"] == "All" else rule["grouping_scope"]
        return f"Exclude {rule['supplier_scope']} from {grouping_text}"
    elif rule["rule_type"] == "Bid Exclusions":
        if rule["bid_grouping"] in ["Milage"]:
            return f"Bid Exclusions on {rule['bid_grouping']}: {rule['operator']} {rule['rule_input']}"
        else:
            return f"Bid Exclusions on {rule['bid_grouping']}: exclude '{rule['bid_exclusion_value']}'"
    else:
        return str(rule)

while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    elif event == "-RULETYPE-":
        if values["-RULETYPE-"] == "Supplier Exclusion":
            window["-RULEOP-"].update(value="", disabled=True)
            window["-RULEINPUT-"].update(value="", disabled=True)
            window["-SUPPSCOPE-"].update(disabled=False)
            window["-BIDGROUPING-"].update(visible=False)
            window["-BIDEXCLUSIONVALUE-"].update(visible=False)
        elif values["-RULETYPE-"] == "# of transitions":
            window["-RULEOP-"].update(disabled=False)
            window["-RULEINPUT-"].update(disabled=False)
            window["-SUPPSCOPE-"].update(value="", disabled=True)
            window["-BIDGROUPING-"].update(visible=False)
            window["-BIDEXCLUSIONVALUE-"].update(visible=False)
        elif values["-RULETYPE-"] == "# of suppliers":
            window["-RULEOP-"].update(disabled=False)
            window["-RULEINPUT-"].update(disabled=False)
            window["-SUPPSCOPE-"].update(value="", disabled=True)
            window["-BIDGROUPING-"].update(visible=False)
            window["-BIDEXCLUSIONVALUE-"].update(visible=False)
        elif values["-RULETYPE-"] in ["% of Volume Awarded", "# of Volume Awarded"]:
            window["-RULEOP-"].update(disabled=False)
            window["-RULEINPUT-"].update(disabled=False)
            window["-SUPPSCOPE-"].update(visible=True, disabled=False)
            window["-BIDGROUPING-"].update(visible=False)
            window["-BIDEXCLUSIONVALUE-"].update(visible=False)
        elif values["-RULETYPE-"] == "Bid Exclusions":
            window["-SUPPSCOPE-"].update(value="", disabled=True)
            window["-BIDGROUPING-"].update(visible=True)
            window["-BIDEXCLUSIONVALUE-"].update(visible=False)
            window["-RULEOP-"].update(visible=True, disabled=False)
            window["-RULEINPUT-"].update(visible=True, disabled=False)
    elif event == "-GROUPING-":
        grouping = values["-GROUPING-"]
        if grouping == "All":
            window["-GROUPSCOPE-"].update(value="", values=[], disabled=True)
        else:
            window["-GROUPSCOPE-"].update(disabled=False)
            scope_vals = update_grouping_scope(grouping, default_item_attributes)
            window["-GROUPSCOPE-"].update(values=scope_vals, value=scope_vals[0] if scope_vals else "")
    elif event == "-BIDGROUPING-":
        bid_group = values["-BIDGROUPING-"]
        if bid_group == "Milage":
            window["-BIDEXCLUSIONVALUE-"].update(visible=False)
            window["-RULEOP-"].update(visible=True, disabled=False)
            window["-RULEINPUT-"].update(visible=True, disabled=False)
        else:
            window["-RULEOP-"].update(value="", visible=False, disabled=True)
            window["-RULEINPUT-"].update(value="", visible=False, disabled=True)
            # Collect unique bid attribute values over all supplier-bid pairs
            unique_vals = sorted({str(val.get(bid_group, "")).strip() for key, val in default_supplier_bid_attributes.items()})
            window["-BIDEXCLUSIONVALUE-"].update(values=unique_vals, visible=True)
            if unique_vals:
                window["-BIDEXCLUSIONVALUE-"].update(value=unique_vals[0])
    elif event == "-ADDRULE-":
        try:
            if values["-RULETYPE-"] in ["% of Volume Awarded", "# of Volume Awarded"]:
                rule_input = float(values["-RULEINPUT-"])
            elif values["-RULETYPE-"] in ["# of transitions", "# of suppliers"]:
                rule_input = int(values["-RULEINPUT-"])
            else:
                rule_input = None
        except:
            sg.popup_error("Please enter a valid number for rule input.")
            continue
        rule = {
            "rule_type": values["-RULETYPE-"],
            "operator": values["-RULEOP-"],
            "rule_input": values["-RULEINPUT-"],
            "grouping": values["-GROUPING-"],
            "grouping_scope": values["-GROUPSCOPE-"],
            "supplier_scope": values["-SUPPSCOPE-"]
        }
        if values["-RULETYPE-"] == "Bid Exclusions":
            rule["bid_grouping"] = values["-BIDGROUPING-"]
            rule["bid_exclusion_value"] = values["-BIDEXCLUSIONVALUE-"]
        rules_list.append(rule)
        window["-RULELIST-"].update("\n".join([f"{i+1}. " + rule_to_text(r) for i, r in enumerate(rules_list)]))
    elif event == "-CLEARRULES-":
        rules_list = []
        window["-RULELIST-"].update("")
    elif event == "Run Optimization":
        use_global = values["-GLOBAL-"]
        if use_global:
            cap_dict = default_global_capacity
        else:
            cap_dict = default_per_item_capacity
        demand_dict = default_demand
        item_attr_dict = default_item_attributes
        price_dict = default_price
        rebate_tiers = default_rebate_tiers
        discount_tiers = default_discount_tiers
        baseline_dict = default_baseline_price

        try:
            output_file, feasibility_notes, model_status = run_optimization(use_global, cap_dict,
                                                                             demand_dict, item_attr_dict,
                                                                             price_dict, rebate_tiers,
                                                                             discount_tiers, baseline_dict,
                                                                             rules_list)
        except KeyError as e:
            sg.popup_error("Model may be infeasible due to custom rule constraints, capacity, or demand issues.\n"
                           "Please review your custom rules and input data.\nError details: " + str(e))
            continue

        sg.popup(f"Model Status: {model_status}\nResults written to:\n{output_file}\n\nFeasibility Notes:\n{feasibility_notes}")

window.close()

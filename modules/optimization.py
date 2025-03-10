import os
import pandas as pd
import pulp

# Global constant for Big-M constraints.
M = 1e9

#############################################
# REQUIRED COLUMNS for Excel Validation
#############################################
# For "Item Attributes", only "Bid ID" and "Incumbent" are required.
# For "Supplier Bid Attributes", only "Supplier Name" and "Bid ID" are required.
REQUIRED_COLUMNS = {
    "Item Attributes": ["Bid ID", "Incumbent"],
    "Price": ["Supplier Name", "Bid ID", "Price"],
    "Demand": ["Bid ID", "Demand"],
    "Baseline Price": ["Bid ID", "Baseline Price"],
    "Capacity": ["Supplier Name", "Capacity Scope", "Scope Value", "Capacity"],
    "Rebate Tiers": ["Supplier Name", "Min", "Max", "Percentage", "Scope Attribute", "Scope Value"],
    "Discount Tiers": ["Supplier Name", "Min", "Max", "Percentage", "Scope Attribute", "Scope Value"],
    "Supplier Bid Attributes": ["Supplier Name", "Bid ID"]
}

#############################################
# Helper Functions for Excel Loading & Validation
#############################################
def load_excel_sheets(uploaded_file):
    """
    Load all sheets from an uploaded Excel file into a dictionary of DataFrames.
    """
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for sheet in xls.sheet_names:
        sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
    return sheets

def validate_sheet(df, sheet_name):
    """
    Check that the DataFrame contains all required columns for the specified sheet.
    Returns a list of missing columns.
    """
    required = REQUIRED_COLUMNS.get(sheet_name, [])
    missing = [col for col in required if col not in df.columns]
    return missing

#############################################
# Data Conversion Helper Functions
#############################################
def df_to_dict_item_attributes(df):
    """
    Convert the "Item Attributes" sheet into a dictionary keyed by Bid ID.
    Only "Bid ID" and "Incumbent" are required; all other columns are read in.
    """
    d = {}
    for _, row in df.iterrows():
        bid = str(row["Bid ID"]).strip()
        d[bid] = row.to_dict()
        d[bid].pop("Bid ID", None)
    return d

def df_to_dict_price(df):
    """
    Convert the "Price" sheet into a dictionary keyed by (Supplier Name, Bid ID) with price values.
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        bid = str(row["Bid ID"]).strip()
        d[(supplier, bid)] = row["Price"]
    return d

def df_to_dict_demand(df):
    """
    Convert the "Demand" sheet into a dictionary keyed by Bid ID with demand values.
    """
    d = {}
    for _, row in df.iterrows():
        bid = str(row["Bid ID"]).strip()
        d[bid] = row["Demand"]
    return d

def df_to_dict_baseline_price(df):
    """
    Convert the "Baseline Price" sheet into a dictionary keyed by Bid ID with baseline prices.
    """
    d = {}
    for _, row in df.iterrows():
        bid = str(row["Bid ID"]).strip()
        d[bid] = row["Baseline Price"]
    return d

def df_to_dict_capacity(df):
    """
    Convert the "Capacity" sheet into a dictionary keyed by (Supplier Name, Capacity Scope, Scope Value)
    with capacity values.
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        cap_scope = str(row["Capacity Scope"]).strip()
        scope_value = str(row["Scope Value"]).strip()
        d[(supplier, cap_scope, scope_value)] = row["Capacity"]
    return d

def df_to_dict_tiers(df):
    """
    Convert a tiers sheet (either "Rebate Tiers" or "Discount Tiers") into a dictionary keyed by Supplier Name.
    Each value is a list of tuples: (Min, Max, Percentage, Scope Attribute, Scope Value)
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        tier = (row["Min"], row["Max"], row["Percentage"], row.get("Scope Attribute"), row.get("Scope Value"))
        if supplier in d:
            d[supplier].append(tier)
        else:
            d[supplier] = [tier]
    return d

def df_to_dict_supplier_bid_attributes(df):
    """
    Convert the "Supplier Bid Attributes" sheet into a dictionary keyed by (Supplier Name, Bid ID)
    with all bid attribute information (all columns besides Supplier Name and Bid ID).
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        bid = str(row["Bid ID"]).strip()
        attr = row.to_dict()
        attr.pop("Supplier Name", None)
        attr.pop("Bid ID", None)
        d[(supplier, bid)] = attr
    return d

#############################################
# Helper for Custom Rule Text Representation
#############################################
def rule_to_text(rule):
    """
    Return a string representation of a custom rule.
    """
    if rule["rule_type"] == "% of Volume Awarded":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"% Vol: {rule['operator']} {rule['rule_input']}% of {grouping_text} awarded to {rule['supplier_scope']}"
    elif rule["rule_type"] == "# of Volume Awarded":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"# Vol: {rule['operator']} {rule['rule_input']} units of {grouping_text} awarded to {rule['supplier_scope']}"
    elif rule["rule_type"] == "# of transitions":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"# Transitions: {rule['operator']} {rule['rule_input']} transitions in {grouping_text}"
    elif rule["rule_type"] == "# of suppliers":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"# Suppliers: {rule['operator']} {rule['rule_input']} unique suppliers in {grouping_text}"
    elif rule["rule_type"] == "Supplier Exclusion":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"Exclude {rule['supplier_scope']} from {grouping_text}"
    elif rule["rule_type"] == "Bid Exclusions":
        bid_grouping = rule.get("bid_grouping", "Not specified")
        if is_bid_attribute_numeric(bid_grouping, {}):
            return f"Bid Exclusions on {bid_grouping}: {rule['operator']} {rule['rule_input']}"
        else:
            bid_exclusion_value = rule.get("bid_exclusion_value", "Not specified")
            return f"Bid Exclusions on {bid_grouping}: exclude '{bid_exclusion_value}'"
    elif rule["rule_type"] == "# Minimum volume awarded":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"# Min Vol: at least {rule['rule_input']} units in {grouping_text}"
    elif rule["rule_type"] == "% Minimum volume awarded":
        grouping_text = "all items" if rule["grouping"] == "All" or not rule["grouping_scope"] else rule["grouping_scope"]
        return f"% Min Vol: at least {rule['rule_input']}% in {grouping_text}"
    else:
        return str(rule)

#############################################
# Helper: Determine if a bid attribute is numeric.
#############################################
def is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict):
    """
    Determine if the specified bid attribute (bid_group) is numeric.
    It checks the first non-None occurrence in supplier_bid_attr_dict.
    """
    for key, attr in supplier_bid_attr_dict.items():
        if bid_group in attr and attr[bid_group] is not None:
            try:
                float(attr[bid_group])
                return True
            except:
                return False
    return False

#############################################
# Main Optimization Function
#############################################
def run_optimization(capacity_data, demand_data, item_attr_data, price_data,
                     rebate_tiers, discount_tiers, baseline_price_data, rules=[],
                     supplier_bid_attr_dict=None, suppliers=None):
    """
    Run the sourcing optimization model.

    Parameters:
      - capacity_data: dict keyed by (Supplier Name, Capacity Scope, Scope Value) with capacities.
      - demand_data: dict keyed by Bid ID with demand values.
      - item_attr_data: dict keyed by Bid ID with item attributes.
      - price_data: dict keyed by (Supplier Name, Bid ID) with prices.
      - rebate_tiers: dict keyed by Supplier; each value is a list of tuples (Min, Max, Percentage, Scope Attribute, Scope Value).
      - discount_tiers: dict keyed by Supplier; each value is a list of tuples (Min, Max, Percentage, Scope Attribute, Scope Value).
      - baseline_price_data: dict keyed by Bid ID with baseline prices.
      - rules: list of custom rule dictionaries.
      - supplier_bid_attr_dict: dict keyed by (Supplier Name, Bid ID) with supplier bid attributes.
      - suppliers: list of supplier names.

    Returns:
      - output_file (str): path to an Excel file with optimization results.
      - feasibility_notes (str): feasibility information.
      - model_status (str): PuLP model status.
    """
    if supplier_bid_attr_dict is None:
        raise ValueError("supplier_bid_attr_dict must be provided from the 'Supplier Bid Attributes' sheet.")
    if suppliers is None:
        raise ValueError("suppliers must be provided (extracted from the 'Price' sheet).")
    
    items_dynamic = list(demand_data.keys())
    
    # Create transition variables (for non-incumbent supplier transitions).
    T = {}
    for j in items_dynamic:
        incumbent = item_attr_data[j].get("Incumbent")
        for s in suppliers:
            if s != incumbent:
                T[(j, s)] = pulp.LpVariable(f"T_{j}_{s}", cat='Binary')
    
    lp_problem = pulp.LpProblem("Sourcing_with_MultiTier_Rebates_Discounts", pulp.LpMinimize)
    
    # Decision variables.
    x = {(s, j): pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')
         for s in suppliers for j in items_dynamic}
    S0 = {s: pulp.LpVariable(f"S0_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    S = {s: pulp.LpVariable(f"S_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    V = {s: pulp.LpVariable(f"V_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    
    # Demand constraints.
    for j in items_dynamic:
        lp_problem += pulp.lpSum(x[(s, j)] for s in suppliers) == demand_data[j], f"Demand_{j}"
    
    # Transition constraints.
    for j in items_dynamic:
        for s in suppliers:
            if (j, s) in T:
                lp_problem += x[(s, j)] <= demand_data[j] * T[(j, s)], f"Transition_{j}_{s}"
    
    # Capacity constraints.
    for (s, cap_scope, scope_value), cap in capacity_data.items():
        if cap_scope == "Bid ID":
            items_group = [scope_value] if scope_value in item_attr_data else []
        else:
            items_group = [j for j in items_dynamic if str(item_attr_data[j].get(cap_scope, "")).strip() == str(scope_value).strip()]
        if items_group:
            lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) <= cap, f"Capacity_{s}_{cap_scope}_{scope_value}"
    
    # Base spend and volume.
    for s in suppliers:
        lp_problem += S0[s] == pulp.lpSum(price_data[(s, j)] * x[(s, j)] for j in items_dynamic), f"BaseSpend_{s}"
        lp_problem += V[s] == pulp.lpSum(x[(s, j)] for j in items_dynamic), f"Volume_{s}"
    
    # Compute supplier-specific Big-M values.
    max_price_val = max(price_data.values())
    U_spend = {}
    for s in suppliers:
        total_cap = sum(cap for ((sup, _, _), cap) in capacity_data.items() if sup == s)
        U_spend[s] = total_cap * max_price_val
    
    # Discount tier constraints.
    z_discount = {}
    for s in suppliers:
        tiers = discount_tiers.get(s, [])
        z_discount[s] = {k: pulp.LpVariable(f"z_discount_{s}_{k}", cat='Binary') for k in range(len(tiers))}
        lp_problem += pulp.lpSum(z_discount[s][k] for k in range(len(tiers))) == 1, f"DiscountTierSelect_{s}"
    
    # Rebate tier constraints.
    y_rebate = {}
    for s in suppliers:
        tiers = rebate_tiers.get(s, [])
        y_rebate[s] = {k: pulp.LpVariable(f"y_rebate_{s}_{k}", cat='Binary') for k in range(len(tiers))}
        lp_problem += pulp.lpSum(y_rebate[s][k] for k in range(len(tiers))) == 1, f"RebateTierSelect_{s}"
    
    d = {s: pulp.LpVariable(f"d_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    rebate_var = {s: pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    
    lp_problem += pulp.lpSum(S[s] - rebate_var[s] for s in suppliers), "Total_Effective_Cost"
    
    # Discount Tiers constraints.
    for s in suppliers:
        tiers = discount_tiers.get(s, [])
        M_discount = U_spend[s] if s in U_spend else M
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
    
    # Rebate Tiers constraints.
    for s in suppliers:
        tiers = rebate_tiers.get(s, [])
        M_rebate = U_spend[s] if s in U_spend else M
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
    
    # Compute lowest cost suppliers per bid.
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
        # "# of suppliers" rule.
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
                        lp_problem += pulp.lpSum(x[(s, j)] for j in [j]) <= M * w[(s, j)], f"RuleSupplierIndicator_{r_idx}_{j}_{s}"
                        lp_problem += pulp.lpSum(x[(s, j)] for j in [j]) >= 1e-3 * w[(s, j)], f"RuleSupplierIndicatorLB_{r_idx}_{j}_{s}"
                    lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) == 1, f"RuleSingleSupplier_{r_idx}_{j}"
            else:
                if rule["grouping"] == "All" or not rule["grouping_scope"]:
                    items_group = items_dynamic
                elif rule["grouping"] == "Bid ID":
                    items_group = [rule["grouping_scope"]]
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
                    lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) >= 1e-3 * w[s], f"SupplierIndicatorLB_{r_idx}_{s}"
                total_suppliers = pulp.lpSum(w[s] for s in suppliers)
                if operator == "At least":
                    lp_problem += total_suppliers >= supplier_target, f"Rule_{r_idx}"
                elif operator == "At most":
                    lp_problem += total_suppliers <= supplier_target, f"Rule_{r_idx}"
                elif operator == "Exactly":
                    lp_problem += total_suppliers == supplier_target, f"Rule_{r_idx}"
        # "% of Volume Awarded" rule.
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
            else:
                if rule["grouping"] == "All" or not rule["grouping_scope"]:
                    items_group = items_dynamic
                elif rule["grouping"] == "Bid ID":
                    items_group = [rule["grouping_scope"]]
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
        # "# of Volume Awarded" rule.
        elif rule["rule_type"] == "# of Volume Awarded":
            if rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            elif rule["grouping"] == "Bid ID":
                items_group = [rule["grouping_scope"]]
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            try:
                volume_target = float(rule["rule_input"])
            except:
                continue
            operator = rule["operator"]
            lhs = pulp.lpSum(x[(rule["supplier_scope"], j)] for j in items_group)
            if operator == "At least":
                lp_problem += lhs >= volume_target, f"Rule_{r_idx}"
            elif operator == "At most":
                lp_problem += lhs <= volume_target, f"Rule_{r_idx}"
            elif operator == "Exactly":
                lp_problem += lhs == volume_target, f"Rule_{r_idx}"
        # "# of transitions" rule.
        elif rule["rule_type"] == "# of transitions":
            if rule["grouping"] == "All" or not rule["grouping_scope"]:
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
        # "Bid Exclusions" rule.
        elif rule["rule_type"] == "Bid Exclusions":
            if rule["grouping"] == "Bid ID":
                items_group = [rule["grouping_scope"]]
            elif rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            bid_group = rule.get("bid_grouping", None)
            if bid_group is None:
                continue
            is_numeric = is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict)
            for j in items_group:
                for s in suppliers:
                    bid_val = supplier_bid_attr_dict.get((s, j), {}).get(bid_group, None)
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
        # "Supplier Exclusion" rule.
        elif rule["rule_type"] == "Supplier Exclusion":
            if rule["grouping"] == "Bid ID":
                items_group = [rule["grouping_scope"]]
            elif rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            for j in items_group:
                lp_problem += x[(rule["supplier_scope"], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
        # "# Minimum volume awarded" rule.
        elif rule["rule_type"] == "# Minimum volume awarded":
            if rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            elif rule["grouping"] == "Bid ID":
                items_group = [rule["grouping_scope"]]
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            for s in suppliers:
                y = pulp.LpVariable(f"MinVol_{r_idx}_{s}", cat='Binary')
                lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) <= M * y, f"MinVol_UB_{r_idx}_{s}"
                lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) >= float(rule["rule_input"]) * y, f"MinVol_LB_{r_idx}_{s}"
        # "% Minimum volume awarded" rule.
        elif rule["rule_type"] == "% Minimum volume awarded":
            if rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            elif rule["grouping"] == "Bid ID":
                items_group = [rule["grouping_scope"]]
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            group_demand = sum(demand_data[j] for j in items_group)
            for s in suppliers:
                y = pulp.LpVariable(f"MinVolPct_{r_idx}_{s}", cat='Binary')
                lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) <= M * y, f"MinVolPct_UB_{r_idx}_{s}"
                lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) >= (float(rule["rule_input"]) / 100.0) * group_demand * y, f"MinVolPct_LB_{r_idx}_{s}"
    
    # Debug output.
    constraint_names = list(lp_problem.constraints.keys())
    duplicates = set([n for n in constraint_names if constraint_names.count(n) > 1])
    if duplicates:
        print("DEBUG: Duplicate constraint names found:", duplicates)
    print("DEBUG: Total constraints added:", len(constraint_names))
    
    # Solve the model.
    solver = pulp.PULP_CBC_CMD(msg=False, gapRel=0, gapAbs=0)
    lp_problem.solve(solver)
    model_status = pulp.LpStatus[lp_problem.status]
    
    feasibility_notes = ""
    if model_status == "Infeasible":
        feasibility_notes += "Model is infeasible. Likely causes include:\n"
        feasibility_notes += " - Insufficient supplier capacity relative to demand.\n"
        feasibility_notes += " - Custom rule constraints conflicting with overall volume/demand.\n"
        for j in items_dynamic:
            cap_note = ""
            for (s, cap_scope, scope_value), cap in capacity_data.items():
                if cap_scope == "Bid ID" and scope_value == j:
                    cap_note += f"Supplier {s} (Bid ID capacity): {cap}; "
                elif cap_scope != "Bid ID" and str(item_attr_data[j].get(cap_scope, "")).strip() == str(scope_value).strip():
                    cap_note += f"Supplier {s} ({cap_scope}={scope_value} capacity): {cap}; "
            feasibility_notes += f"  Bid {j}: demand = {demand_data[j]}, capacities: {cap_note}\n"
        feasibility_notes += "Please review supplier capacities, demand, and custom rule constraints."
    else:
        feasibility_notes = "Model is optimal."
    
    # Prepare results.
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
            for k, tier in enumerate(discount_tiers.get(s, [])):
                if pulp.value(z_discount[s][k]) is not None and pulp.value(z_discount[s][k]) >= 0.5:
                    active_discount = tier[2]
                    break
            discount_pct = active_discount
            discounted_price = orig_price * (1 - discount_pct)
            awarded_spend = discounted_price * award_val
            base_price = baseline_price_data[j]
            baseline_spend = base_price * award_val
            baseline_savings = baseline_spend - awarded_spend
            active_rebate = 0
            for k, tier in enumerate(rebate_tiers.get(s, [])):
                if pulp.value(y_rebate[s][k]) is not None and pulp.value(y_rebate[s][k]) >= 0.5:
                    active_rebate = tier[2]
                    break
            rebate_savings = awarded_spend * active_rebate
            facility_val = item_attr_data[j].get("Facility", "")
            row = {
                "Bid ID": idx,
                "Bid ID Split": bid_split,
                "Facility": facility_val,
                "Incumbent": item_attr_data[j].get("Incumbent", ""),
                "Baseline Price": base_price,
                "Baseline Spend": baseline_spend,
                "Awarded Supplier": s,
                "Original Awarded Supplier Price": orig_price,
                "Percentage Volume Discount": f"{discount_pct*100:.0f}%" if discount_pct else "0%",
                "Discounted Awarded Supplier Price": discounted_price,
                "Awarded Supplier Spend": awarded_spend,
                "Awarded Volume": award_val,
                "Baseline Savings": baseline_savings,
                "Rebate %": f"{active_rebate*100:.0f}%" if active_rebate else "0%",
                "Rebate Savings": rebate_savings
            }
            excel_rows.append(row)
    
    df_results = pd.DataFrame(excel_rows)
    cols = ["Bid ID", "Bid ID Split", "Facility", "Incumbent", "Baseline Price", "Baseline Spend",
            "Awarded Supplier", "Original Awarded Supplier Price", "Percentage Volume Discount",
            "Discounted Awarded Supplier Price", "Awarded Supplier Spend", "Awarded Volume",
            "Baseline Savings", "Rebate %", "Rebate Savings"]
    df_results = df_results[cols]
    
    df_feasibility = pd.DataFrame({"Feasibility Notes": [feasibility_notes]})
    temp_lp_file = os.path.join(os.getcwd(), "temp_model.lp")
    lp_problem.writeLP(temp_lp_file)
    with open(temp_lp_file, "r") as f:
        lp_text = f.read()
    df_lp = pd.DataFrame({"LP Model": [lp_text]})
    
    output_file = os.path.join(os.getcwd(), "optimization_results.xlsx")
    capacity_df = pd.DataFrame([
        {"Supplier Name": s, "Capacity Scope": cs, "Scope Value": sv, "Capacity": cap}
        for (s, cs, sv), cap in capacity_data.items()
    ])
    
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_results.to_excel(writer, sheet_name="Results", index=False)
        df_feasibility.to_excel(writer, sheet_name="Feasibility Notes", index=False)
        df_lp.to_excel(writer, sheet_name="LP Model", index=False)
        capacity_df.to_excel(writer, sheet_name="Capacity", index=False)
    
    return output_file, feasibility_notes, model_status

if __name__ == "__main__":
    print("Optimization module loaded. Please call run_optimization() from your Streamlit app.")

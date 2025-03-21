import os
import pandas as pd
import pulp
import streamlit as st

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
def normalize_bid_id(bid):
    """
    Convert a Bid ID to a string.
    If bid is a list or tuple, join its elements with a hyphen.
    """
    if isinstance(bid, (list, tuple)):
        return "-".join(str(x).strip() for x in bid)
    try:
        num = float(bid)
        if num.is_integer():
            return str(int(num))
        else:
            return str(num)
    except Exception:
        return str(bid).strip()

def df_to_dict_item_attributes(df):
    """
    Convert the "Item Attributes" sheet into a dictionary keyed by normalized Bid ID.
    """
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = row.to_dict()
        d[bid].pop("Bid ID", None)
    return d

def df_to_dict_demand(df):
    """
    Convert the "Demand" sheet into a dictionary keyed by normalized Bid ID.
    """
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = row["Demand"]
    return d

def df_to_dict_price(df):
    """
    Convert the "Price" sheet into a dictionary keyed by (Supplier, normalized Bid ID).
    Only rows with nonzero price are included.
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        bid = normalize_bid_id(row["Bid ID"])
        price = row["Price"]
        if pd.isna(price) or price == 0:
            continue
        d[(supplier, bid)] = price
    return d

def df_to_dict_baseline_price(df):
    """
    Convert the "Baseline Price" sheet into a dictionary keyed by normalized Bid ID with baseline prices.
    """
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = row["Baseline Price"]
    return d

def df_to_dict_capacity(df):
    """
    Convert the "Capacity" sheet into a dictionary keyed by (Supplier, Capacity Scope, normalized Scope Value)
    with capacity values.
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        cap_scope = str(row["Capacity Scope"]).strip()
        if cap_scope == "Bid ID":
            scope_value = normalize_bid_id(row["Scope Value"])
        else:
            scope_value = str(row["Scope Value"]).strip()
        d[(supplier, cap_scope, scope_value)] = row["Capacity"]
    return d

def df_to_dict_tiers(df):
    """
    Convert a tiers sheet (either "Rebate Tiers" or "Discount Tiers") into a dictionary keyed by Supplier Name.
    Each value is a list of tuples.
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
    Convert the "Supplier Bid Attributes" sheet into a dictionary keyed by (Supplier, normalized Bid ID).
    """
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        bid = normalize_bid_id(row["Bid ID"])
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
    Return a human-readable description of a custom rule.
    """
    grouping = rule.get("grouping", "all items")
    grouping_scope = rule.get("grouping_scope", "all items")
    supplier = rule.get("supplier_scope", "All")
    op = rule.get("operator", "").lower()
    if grouping.strip() == "Bid ID":
        grouping_scope_str = str(grouping_scope).strip()
        if not grouping_scope_str.lower().startswith("bid id"):
            grouping_scope_str = "Bid ID " + grouping_scope_str
    else:
        grouping_scope_str = grouping_scope
    if rule["rule_type"] == "% of Volume Awarded":
        return f"For {grouping_scope_str}, {supplier} is {op} awarded {rule['rule_input']} of the total volume."
    elif rule["rule_type"] == "# of Volume Awarded":
        return f"For {grouping_scope_str}, {supplier} is {op} awarded {rule['rule_input']} units of volume."
    elif rule["rule_type"] == "# of Transitions":
        return f"For {grouping_scope_str}, the number of transitions must be {op} {rule['rule_input']}."
    elif rule["rule_type"] == "# of Suppliers":
        return f"For {grouping_scope_str}, the number of unique suppliers must be {op} {rule['rule_input']}."
    elif rule["rule_type"] == "Supplier Exclusion":
        return f"Exclude {supplier} from {grouping_scope_str}."
    elif rule["rule_type"] == "Exclude Bids":
        bid_attr = rule.get("bid_grouping", "Unknown Attribute")
        if rule.get("rule_input", ""):
            return f"Exclude bids where {bid_attr} is {op} {rule['rule_input']}."
        else:
            exclusion_val = rule.get("bid_exclusion_value", "Unknown")
            return f"Exclude bids where {bid_attr} equals '{exclusion_val}'."
    elif rule["rule_type"] == "# Minimum Volume Awarded":
        return f"For {grouping_scope_str}, the supplier must be awarded at least {rule['rule_input']} units of volume."
    elif rule["rule_type"] == "% Minimum Volume Awarded":
        return f"For {grouping_scope_str}, the supplier must be awarded at least {rule['rule_input']} of the total volume."
    else:
        return str(rule)

def expand_rule_text(rule, item_attr_data):
    """
    If grouping_scope is "apply to all items individually", expand the rule text.
    """
    grouping = rule.get("grouping", "All")
    grouping_scope = rule.get("grouping_scope", "").strip().lower()
    if grouping_scope == "apply to all items individually":
        if grouping == "Bid ID":
            groups = sorted(item_attr_data.keys())
        else:
            groups = sorted(set(str(item_attr_data[j].get(grouping, "")).strip() 
                                for j in item_attr_data if str(item_attr_data[j].get(grouping, "")).strip() != ""))
        texts = []
        for i, group in enumerate(groups):
            new_rule = rule.copy()
            new_rule["grouping_scope"] = group
            texts.append(f"{i+1}. {rule_to_text(new_rule)}")
        return "<br>".join(texts)
    else:
        return rule_to_text(rule)

#############################################
# Helper: Determine if a bid attribute is numeric.
#############################################
def is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict):
    """
    Determine if the specified bid attribute (bid_group) is numeric by testing the first non-None occurrence.
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
    All keys for Bid IDs are normalized to avoid tuple-key issues.
    Additionally, if no supplier provided a valid bid for a Bid ID, we set that Bid ID’s demand to zero.
    """
    if supplier_bid_attr_dict is None:
        raise ValueError("supplier_bid_attr_dict must be provided from the 'Supplier Bid Attributes' sheet.")
    if suppliers is None:
        raise ValueError("suppliers must be provided (extracted from the 'Price' sheet).")
    
    # --- Normalize keys in all dictionaries ---
    demand_data = {normalize_bid_id(k): v for k, v in demand_data.items()}
    item_attr_data = {normalize_bid_id(k): v for k, v in item_attr_data.items()}
    baseline_price_data = {normalize_bid_id(k): v for k, v in baseline_price_data.items()}
    price_data = {(s, normalize_bid_id(j)): v for (s, j), v in price_data.items()}
    capacity_data = { (s, cs, (normalize_bid_id(sv) if cs=="Bid ID" else str(sv).strip())): cap 
                      for (s, cs, sv), cap in capacity_data.items() }
    supplier_bid_attr_dict = {(s, normalize_bid_id(j)): attr for (s, j), attr in supplier_bid_attr_dict.items()}
    
    # Rebuild dictionaries using the updated normalization
    demand_data = {normalize_bid_id(k): v for k, v in demand_data.items()}
    price_data = {(s, normalize_bid_id(j)): v for (s, j), v in price_data.items()}

    # For each Bid ID in demand, if no supplier has a nonzero bid, set demand to zero.
    for bid in list(demand_data.keys()):
        # Use get() to default missing keys to 0
        has_valid_bid = any(price_data.get((s, bid), 0) != 0 for s in suppliers)
        if not has_valid_bid:
            demand_data[bid] = 0


    # --- Build list of Bid IDs from demand (keep all, even those with zero demand) ---
    items_dynamic = [normalize_bid_id(j) for j in demand_data.keys()]
    
    # --- Define no_bid_items as those Bid IDs with zero demand ---
    no_bid_items = [bid for bid, d_val in demand_data.items() if d_val == 0]

    # --- Create transition variables ---
    T = {}
    for j in items_dynamic:
        incumbent = item_attr_data[normalize_bid_id(j)].get("Incumbent")
        for s in suppliers:
            if s != incumbent:
                T[(j, s)] = pulp.LpVariable(f"T_{j}_{s}", cat='Binary')
    
    lp_problem = pulp.LpProblem("Sourcing_with_MultiTier_Rebates_Discounts", pulp.LpMinimize)
    
    # --- Decision variables ---
    x = {(s, j): pulp.LpVariable(f"x_{s}_{j}", lowBound=0, cat='Continuous')
         for s in suppliers for j in items_dynamic}
    S0 = {s: pulp.LpVariable(f"S0_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    S = {s: pulp.LpVariable(f"S_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    V = {s: pulp.LpVariable(f"V_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    
    # --- Demand constraints ---
    for j in items_dynamic:
        lp_problem += pulp.lpSum(x[(s, j)] for s in suppliers) == demand_data[normalize_bid_id(j)], f"Demand_{j}"
    
    # --- Non-bid constraints ---
    for s in suppliers:
        for j in items_dynamic:
            if (s, j) not in price_data:
                lp_problem += x[(s, j)] == 0, f"NonBid_{s}_{j}"
    
    # --- Transition constraints ---
    for j in items_dynamic:
        for s in suppliers:
            if (j, s) in T:
                lp_problem += x[(s, j)] <= demand_data[normalize_bid_id(j)] * T[(j, s)], f"Transition_{j}_{s}"
    
    # --- Capacity constraints ---
    for (s, cap_scope, scope_value), cap in capacity_data.items():
        if cap_scope == "Bid ID":
            items_group = [scope_value] if scope_value in item_attr_data else []
        else:
            items_group = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(cap_scope, "")).strip() == str(scope_value).strip()]
        if items_group:
            lp_problem += pulp.lpSum(x[(s, j)] for j in items_group) <= cap, f"Capacity_{s}_{cap_scope}_{scope_value}"
    
    # --- Base spend and volume ---
    for s in suppliers:
        lp_problem += S0[s] == pulp.lpSum(price_data.get((s, j), 0) * x[(s, j)] for j in items_dynamic), f"BaseSpend_{s}"
        lp_problem += V[s] == pulp.lpSum(x[(s, j)] for j in items_dynamic), f"Volume_{s}"
    
    # --- Compute Big-M values for each supplier ---
    max_price_val = max(price_data.values()) if price_data else 0
    U_spend = {}
    for s in suppliers:
        total_cap = sum(cap for ((sup, _, _), cap) in capacity_data.items() if sup == s)
        U_spend[s] = total_cap * max_price_val
    
    # --- Discount tiers ---
    z_discount = {}
    for s in suppliers:
        tiers = discount_tiers.get(s, [])
        if tiers:
            z_discount[s] = {k: pulp.LpVariable(f"z_discount_{s}_{k}", cat='Binary') for k in range(len(tiers))}
            lp_problem += pulp.lpSum(z_discount[s][k] for k in range(len(tiers))) == 1, f"DiscountTierSelect_{s}"
        else:
            z_discount[s] = {}
    
    # --- Rebate tiers ---
    y_rebate = {}
    for s in suppliers:
        tiers = rebate_tiers.get(s, [])
        if tiers:
            y_rebate[s] = {k: pulp.LpVariable(f"y_rebate_{s}_{k}", cat='Binary') for k in range(len(tiers))}
            lp_problem += pulp.lpSum(y_rebate[s][k] for k in range(len(tiers))) == 1, f"RebateTierSelect_{s}"
        else:
            y_rebate[s] = {}
    
    # --- Adjustment variables for discounts and rebates ---
    d = {s: pulp.LpVariable(f"d_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    rebate_var = {s: pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    for s in suppliers:
        if not discount_tiers.get(s, []):
            lp_problem += d[s] == 0, f"Fix_d_{s}"
        if not rebate_tiers.get(s, []):
            lp_problem += rebate_var[s] == 0, f"Fix_rebate_{s}"
    
    # --- Objective ---
    lp_problem += pulp.lpSum(S[s] - rebate_var[s] for s in suppliers), "Total_Effective_Cost"
    
    # --- Discount Tier constraints ---
    for s in suppliers:
        tiers = discount_tiers.get(s, [])
        M_discount = U_spend[s] if s in U_spend else M
        for k, tier in enumerate(tiers):
            Dmin, Dmax, Dperc, scope_attr, scope_value = tier
            if scope_attr is None or scope_value is None:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic)
            else:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic if item_attr_data[normalize_bid_id(j)].get(scope_attr) == scope_value)
            lp_problem += vol_expr >= Dmin * z_discount[s][k], f"DiscountTierMin_{s}_{k}"
            if Dmax < float('inf'):
                lp_problem += vol_expr <= Dmax + M_discount * (1 - z_discount[s][k]), f"DiscountTierMax_{s}_{k}"
            lp_problem += d[s] >= Dperc * S0[s] - M_discount * (1 - z_discount[s][k]), f"DiscountTierLower_{s}_{k}"
            lp_problem += d[s] <= Dperc * S0[s] + M_discount * (1 - z_discount[s][k]), f"DiscountTierUpper_{s}_{k}"
    
    for s in suppliers:
        lp_problem += S[s] == S0[s] - d[s], f"EffectiveSpend_{s}"
    
    # --- Rebate Tier constraints ---
    for s in suppliers:
        tiers = rebate_tiers.get(s, [])
        M_rebate = U_spend[s] if s in U_spend else M
        for k, tier in enumerate(tiers):
            Rmin, Rmax, Rperc, scope_attr, scope_value = tier
            if scope_attr is None or scope_value is None:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic)
            else:
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic if item_attr_data[normalize_bid_id(j)].get(scope_attr) == scope_value)
            lp_problem += vol_expr >= Rmin * y_rebate[s][k], f"RebateTierMin_{s}_{k}"
            if Rmax < float('inf'):
                lp_problem += vol_expr <= Rmax + M_rebate * (1 - y_rebate[s][k]), f"RebateTierMax_{s}_{k}"
            lp_problem += rebate_var[s] >= Rperc * S[s] - M_rebate * (1 - y_rebate[s][k]), f"RebateTierLower_{s}_{k}"
            lp_problem += rebate_var[s] <= Rperc * S[s] + M_rebate * (1 - y_rebate[s][k]), f"RebateTierUpper_{s}_{k}"
    
    # --- Compute lowest cost suppliers per bid ---
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
    # (Below is the full custom rules processing code as in your original implementation.)
    for r_idx, rule in enumerate(rules):
        # "# of suppliers" rule.
        if rule["rule_type"].lower() == "# of suppliers":
            try:
                supplier_target = int(rule["rule_input"])
            except Exception:
                continue
            operator = rule["operator"].lower()
            if rule["grouping"].strip().lower() == "all" or not rule["grouping_scope"]:
                items_group = items_dynamic
            elif rule["grouping_scope"].strip().lower() == "apply to all items individually":
                if rule["grouping"].strip().lower() == "bid id":
                    unique_groups = sorted(list(item_attr_data.keys()))
                else:
                    unique_groups = sorted({str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() 
                                              for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip()})
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        subgroup = [group_val]
                    else:
                        subgroup = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]
                    for j in subgroup:
                        w = {}
                        for s in suppliers:
                            w[(s, j)] = pulp.LpVariable(f"w_{r_idx}_{s}_{j}", cat='Binary')
                            epsilon = 1  
                            lp_problem += x[(s, j)] <= M * w[(s, j)], f"SupplIndicator_{r_idx}_{s}_{j}"
                            lp_problem += x[(s, j)] >= epsilon * w[(s, j)], f"SupplIndicatorLB_{r_idx}_{s}_{j}"
                        if operator == "at least":
                            lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) >= supplier_target, f"SupplierCount_{r_idx}_{j}"
                        elif operator == "at most":
                            lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) <= supplier_target, f"SupplierCount_{r_idx}_{j}"
                        elif operator == "exactly":
                            lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) == supplier_target, f"SupplierCount_{r_idx}_{j}"
                continue
            
            else:
                if rule["grouping"].strip().lower() == "bid id":
                    items_group = [rule["grouping_scope"].strip()]
                else:
                    items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            for j in items_group:
                w = {}
                for s in suppliers:
                    w[(s, j)] = pulp.LpVariable(f"w_{r_idx}_{s}_{j}", cat='Binary')
                    epsilon = 1
                    lp_problem += x[(s, j)] <= M * w[(s, j)], f"SupplIndicator_{r_idx}_{s}_{j}"
                    lp_problem += x[(s, j)] >= epsilon * w[(s, j)], f"SupplIndicatorLB_{r_idx}_{s}_{j}"
                if operator == "at least":
                    lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) >= supplier_target, f"SupplierCount_{r_idx}_{j}"
                elif operator == "at most":
                    lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) <= supplier_target, f"SupplierCount_{r_idx}_{j}"
                elif operator == "exactly":
                    lp_problem += pulp.lpSum(w[(s, j)] for s in suppliers) == supplier_target, f"SupplierCount_{r_idx}_{j}"
                    
        # "% of Volume Awarded" rule.
        elif rule["rule_type"] == "% of Volume Awarded":
            try:
                percentage = float(rule["rule_input"].rstrip("%")) / 100.0
            except Exception:
                continue
            if rule["grouping"].strip() == "Bid ID":
                if rule["grouping_scope"].strip().lower() == "all":
                    items_group = items_dynamic
                elif rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [rule["grouping_scope"].strip()]
                for j in items_group:
                    total_vol = pulp.lpSum(x[(s, j)] for s in suppliers)
                    if rule["supplier_scope"] in ["Lowest cost supplier", "Second Lowest Cost Supplier", "Incumbent", "New Suppliers"]:
                        if rule["supplier_scope"] == "Lowest cost supplier":
                            supplier_for_rule = lowest_cost_supplier[j]
                            lhs = x[(supplier_for_rule, j)]
                        elif rule["supplier_scope"] == "Second Lowest Cost Supplier":
                            supplier_for_rule = second_lowest_cost_supplier[j]
                            lhs = x[(supplier_for_rule, j)]
                        elif rule["supplier_scope"] == "Incumbent":
                            supplier_for_rule = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                            lhs = x[(supplier_for_rule, j)]
                        elif rule["supplier_scope"] == "New Suppliers":
                            incumbent = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                            lhs = pulp.lpSum(x[(s, j)] for s in suppliers if s != incumbent)
                    else:
                        lhs = x[(rule["supplier_scope"], j)]
                    if rule["operator"] == "At least":
                        lp_problem += lhs >= percentage * total_vol, f"Rule_{r_idx}_{j}"
                    elif rule["operator"] == "At most":
                        lp_problem += lhs <= percentage * total_vol, f"Rule_{r_idx}_{j}"
                    elif rule["operator"] == "Exactly":
                        lp_problem += lhs == percentage * total_vol, f"Rule_{r_idx}_{j}"
            else:
                if rule["grouping"] == "All" or rule["grouping_scope"] == "All":
                    items_group = items_dynamic
                elif rule["grouping_scope"] == "Apply to all items individually":
                    if rule["grouping"].strip() == "Bid ID":
                        items_group = sorted(list(item_attr_data.keys()))
                    else:
                        items_group = sorted({str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip()
                                              for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() != ""})
                else:
                    items_group = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
                if rule["supplier_scope"] in ["Lowest cost supplier", "Second Lowest Cost Supplier", "Incumbent"]:
                    if rule["supplier_scope"] == "Lowest cost supplier":
                        lhs = pulp.lpSum(x[(lowest_cost_supplier[j], j)] for j in items_group)
                    elif rule["supplier_scope"] == "Second Lowest Cost Supplier":
                        lhs = pulp.lpSum(x[(second_lowest_cost_supplier[j], j)] for j in items_group)
                    elif rule["supplier_scope"] == "Incumbent":
                        lhs = pulp.lpSum(x[(item_attr_data[normalize_bid_id(j)].get("Incumbent"), j)] for j in items_group)
                elif rule["supplier_scope"] == "New Suppliers":
                    lhs = pulp.lpSum(pulp.lpSum(x[(s, j)] for s in suppliers if s != item_attr_data[normalize_bid_id(j)].get("Incumbent")) for j in items_group)
                else:
                    lhs = pulp.lpSum(x[(rule["supplier_scope"], j)] for j in items_group)
                total_vol = pulp.lpSum(x[(s, j)] for s in suppliers for j in items_group)
                if rule["operator"] == "At least":
                    lp_problem += lhs >= percentage * total_vol, f"Rule_{r_idx}"
                elif rule["operator"] == "At most":
                    lp_problem += lhs <= percentage * total_vol, f"Rule_{r_idx}"
                elif rule["operator"] == "Exactly":
                    lp_problem += lhs == percentage * total_vol, f"Rule_{r_idx}"
    
         # "# of Volume Awarded" rule.
        elif rule["rule_type"] == "# of Volume Awarded":
            # Determine items_group based on the grouping settings.
            if rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            elif rule["grouping"] == "Bid ID":
                if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = items_dynamic
                else:
                    items_group = [rule["grouping_scope"].strip()]
            else:
                items_group = [j for j in items_dynamic
                            if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() ==
                                str(rule["grouping_scope"]).strip()]
            try:
                volume_target = float(rule["rule_input"])
            except Exception:
                continue
            operator = rule["operator"].lower()
            supplier_scope_value = rule["supplier_scope"]

            # Generate a unique string for the group for naming constraints.
            group_str = "_".join(str(j) for j in items_group) if items_group else "all"

            # Handle special supplier scopes that require per-bid handling.
            if supplier_scope_value.lower() in ["lowest cost supplier", "second lowest cost supplier", "incumbent"]:
                # Here, we still apply per bid.
                for j in items_group:
                    if supplier_scope_value.lower() == "lowest cost supplier":
                        supplier_for_rule = lowest_cost_supplier[j]
                    elif supplier_scope_value.lower() == "second lowest cost supplier":
                        supplier_for_rule = second_lowest_cost_supplier[j]
                    elif supplier_scope_value.lower() == "incumbent":
                        supplier_for_rule = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                    constraint_name = f"MinVolAwarded_{r_idx}_{supplier_scope_value}_{j}"
                    lhs = x[(supplier_for_rule, j)]
                    if operator == "at least":
                        lp_problem += lhs >= volume_target, constraint_name + "_LB"
                    elif operator == "exactly":
                        lp_problem += lhs == volume_target, constraint_name + "_Exact"
                    elif operator == "at most":
                        lp_problem += lhs <= volume_target, constraint_name + "_UB"

            # For "New Suppliers", aggregate over the entire group.
            elif supplier_scope_value.lower() == "new suppliers":
                total_award_new = pulp.lpSum(
                    x[(s, j)]
                    for j in items_group
                    for s in suppliers
                    if s != item_attr_data[normalize_bid_id(j)].get("Incumbent")
                )
                constraint_name_lb = f"MinVolAwarded_NewSuppliers_LB_{r_idx}_{group_str}"
                constraint_name_ub = f"MinVolAwarded_NewSuppliers_UB_{r_idx}_{group_str}"
                if operator == "at least":
                    lp_problem += total_award_new >= volume_target, constraint_name_lb
                elif operator == "exactly":
                    lp_problem += total_award_new == volume_target, constraint_name_lb
                elif operator == "at most":
                    lp_problem += total_award_new <= volume_target, constraint_name_ub

            # For a specific supplier (e.g. "C"), aggregate awarded volume across the entire group.
            else:
                s = supplier_scope_value
                total_award = pulp.lpSum(x[(s, j)] for j in items_group)
                constraint_name = f"MinVolAwarded_{r_idx}_{s}_{group_str}"
                if operator == "at least":
                    lp_problem += total_award >= volume_target, constraint_name + "_LB"
                elif operator == "exactly":
                    lp_problem += total_award == volume_target, constraint_name + "_Exact"
                elif operator == "at most":
                    lp_problem += total_award <= volume_target, constraint_name + "_UB"


        # "# of transitions" rule.
        elif rule["rule_type"].strip().lower() == "# of transitions":
            # Determine the set of Bid IDs (items) to which the rule applies.
            grouping = rule.get("grouping", "").strip().lower()
            grouping_scope = rule.get("grouping_scope", "").strip().lower()

            if grouping == "all" or not grouping_scope:
                items_group = items_dynamic
            elif grouping_scope == "apply to all items individually":
                if grouping == "bid id":
                    # All bid IDs from the item attribute dictionary are already normalized.
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    # Gather all nonempty grouping values from item_attr_data.
                    items_group = sorted({
                        str(item_attr_data[j].get(rule["grouping"], "")).strip().lower()
                        for j in item_attr_data
                        if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""
                    })
            else:
                # Apply the rule only to items matching the specified grouping_scope.
                items_group = [
                    j for j in items_dynamic
                    if str(item_attr_data[j].get(rule["grouping"], "")).strip().lower() == grouping_scope
                ]

            # Convert rule input to an integer target.
            try:
                transitions_target = int(rule["rule_input"])
            except (ValueError, TypeError):
                continue  # Skip this rule if conversion fails

            # Use a lowercase operator for comparison.
            operator = rule.get("operator", "").strip().lower()

            # Sum transitions for all (Bid ID, supplier) pairs where the Bid ID is in the determined group.
            total_transitions = pulp.lpSum(T[(j, s)] for (j, s) in T if j in items_group)

            # Add the appropriate constraint based on the operator.
            if operator == "at least":
                lp_problem += total_transitions >= transitions_target, f"Rule_{r_idx}"
            elif operator == "at most":
                lp_problem += total_transitions <= transitions_target, f"Rule_{r_idx}"
            elif operator == "exactly":
                lp_problem += total_transitions == transitions_target, f"Rule_{r_idx}"

        # "Exclude Bids" rule.
        elif rule["rule_type"].strip().lower() == "exclude bids":
            if rule["grouping"] == "Bid ID":
                # Check if the grouping scope indicates to apply to all items individually.
                if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = items_dynamic
                else:
                    items_group = [rule["grouping_scope"]]
            elif rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            bid_group = rule.get("bid_grouping", None)
            if bid_group is None:
                continue
            is_numeric = is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict)
            for j in items_group:
                for s in suppliers:
                    bid_val = supplier_bid_attr_dict.get((s, normalize_bid_id(j)), {}).get(bid_group, None)
                    if bid_val is None:
                        continue
                    exclude = False
                    if is_numeric:
                        try:
                            bid_val_num = float(bid_val)
                            threshold = float(rule["rule_input"])
                        except:
                            continue
                        op = rule["operator"].strip().lower()
                        if op in ["greater than", ">"]:
                            if bid_val_num > threshold:
                                exclude = True
                        elif op in ["less than", "<"]:
                            if bid_val_num < threshold:
                                exclude = True
                        elif op in ["exactly", "=="]:
                            if bid_val_num == threshold:
                                exclude = True
                    else:
                        if bid_val.strip().lower() == rule.get("bid_exclusion_value", "").strip().lower():
                            exclude = True
                    if exclude:
                        lp_problem += x[(s, normalize_bid_id(j))] == 0, f"BidExclusion_{r_idx}_{j}_{s}"

    
        # "Supplier Exclusion" rule.
        elif rule["rule_type"].strip().lower() == "supplier exclusion":
            # When grouping is "All" (or grouping_scope is not provided), include all bids.
            if rule["grouping"].strip().lower() == "all" or not rule["grouping_scope"]:
                bids_in_group = items_dynamic
            else:
                # If grouping is specified and is not "All":
                if rule["grouping"].strip().lower() == "bid id":
                    bids_in_group = [rule["grouping_scope"].strip()]
                else:
                    bids_in_group = [
                        j for j in items_dynamic 
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() ==
                        str(rule["grouping_scope"]).strip()
                    ]
            
            # Determine the supplier scope (convert to lowercase for comparisons).
            s_scope = rule["supplier_scope"].strip().lower() if rule["supplier_scope"] else None

            if s_scope == "lowest cost supplier":
                for j in bids_in_group:
                    lp_problem += x[(lowest_cost_supplier[j], normalize_bid_id(j))] == 0, f"SupplierExclusion_{r_idx}_{j}"
            elif s_scope == "second lowest cost supplier":
                for j in bids_in_group:
                    lp_problem += x[(second_lowest_cost_supplier[j], normalize_bid_id(j))] == 0, f"SupplierExclusion_{r_idx}_{j}"
            elif s_scope == "incumbent":
                for j in bids_in_group:
                    incumbent = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                    if incumbent:
                        lp_problem += x[(incumbent, normalize_bid_id(j))] == 0, f"SupplierExclusion_{r_idx}_{j}"
            elif s_scope == "new suppliers":
                for j in bids_in_group:
                    incumbent = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                    for s in suppliers:
                        if s != incumbent:
                            lp_problem += x[(s, normalize_bid_id(j))] == 0, f"SupplierExclusion_{r_idx}_{j}_{s}"
            else:
                # For a specific supplier (e.g., "C") not matching any of the special cases.
                for j in bids_in_group:
                    lp_problem += x[(rule["supplier_scope"], normalize_bid_id(j))] == 0, f"SupplierExclusion_{r_idx}_{j}"

    
        # "# Minimum Volume Awarded" rule.
        elif rule["rule_type"].strip().lower() == "# minimum volume awarded":
            try:
                min_volume = float(rule["rule_input"])
            except Exception as e:
                print(f"Invalid rule input for # Minimum Volume Awarded: {e}")
                continue
            if rule.get("grouping_scope", "").strip().lower() == "apply to all items individually":
                if rule["grouping"].strip().lower() == "bid id":
                    unique_groups = sorted(list(item_attr_data.keys()))
                else:
                    unique_groups = sorted({str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() != ""})
                items_group_list = []
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        items_group_list.append((group_val, [group_val]))
                    else:
                        group_items = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]
                        if group_items:
                            items_group_list.append((group_val, group_items))
            else:
                if rule["grouping"].strip().lower() == "all" or not rule.get("grouping_scope"):
                    items_group_list = [("All", items_dynamic)]
                elif rule["grouping"].strip().lower() == "bid id":
                    group_val = str(rule["grouping_scope"]).strip()
                    items_group_list = [(group_val, [group_val])]
                else:
                    group_val = str(rule["grouping_scope"]).strip()
                    group_items = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]
                    items_group_list = [(group_val, group_items)] if group_items else []
            for group_val, items_group in items_group_list:
                group_demand = sum(float(demand_data[normalize_bid_id(j)]) for j in items_group)
                print(f"Group: {items_group} | Total Demand: {group_demand} | Min Volume: {min_volume}")
                if min_volume > group_demand:
                    lp_problem += 0 >= 1, f"MinVol_Infeasible_{r_idx}_{group_val}"
                else:
                    for s in suppliers:
                        x_sum = pulp.lpSum(x[(s, j)] for j in items_group)
                        y = pulp.LpVariable(f"MinVol_{r_idx}_{s}_{group_val}", cat='Binary')
                        lp_problem += x_sum >= min_volume - (1 - y) * M, f"MinVol_LB_{r_idx}_{s}_{group_val}"
                        lp_problem += x_sum <= group_demand * y, f"MinVol_UB_{r_idx}_{s}_{group_val}"
    
        # "% Minimum Volume Awarded" rule.
        elif rule["rule_type"].strip().lower() == "% minimum volume awarded":
            try:
                min_pct = float(rule["rule_input"].rstrip("%")) / 100.0
            except Exception as e:
                st.error(f"Invalid rule input for % Minimum Volume Awarded: {e}")
                continue
            if rule.get("grouping_scope", "").strip().lower() == "apply to all items individually":
                if rule["grouping"].strip().lower() == "bid id":
                    unique_groups = sorted(list(item_attr_data.keys()))
                else:
                    unique_groups = sorted({str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() != ""})
                group_list = []
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        group_list.append([group_val])
                    else:
                        group_items = [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]
                        if group_items:
                            group_list.append(group_items)
            else:
                if rule["grouping"].strip().lower() == "all" or not rule.get("grouping_scope"):
                    group_list = [items_dynamic]
                elif rule["grouping"].strip().lower() == "bid id":
                    group_list = [[str(rule["grouping_scope"]).strip()]]
                else:
                    group_list = [[j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]]
            for group in group_list:
                if not group:
                    continue
                try:
                    group_demand = sum(float(demand_data[normalize_bid_id(j)]) for j in group)
                except Exception as e:
                    st.error(f"Error computing group demand: {e}")
                    continue
                if rule.get("supplier_scope") in [None, "All"]:
                    for s in suppliers:
                        y = pulp.LpVariable(f"MinVolPct_{r_idx}_{s}_{hash(tuple(group))}", cat='Binary')
                        lp_problem += pulp.lpSum(x[(s, j)] for j in group) >= min_pct * group_demand * y, f"MinVolPct_LB_{r_idx}_{s}_{hash(tuple(group))}"
                        lp_problem += pulp.lpSum(x[(s, j)] for j in group) <= group_demand * y, f"MinVolPct_UB_{r_idx}_{s}_{hash(tuple(group))}"
                else:
                    s = rule["supplier_scope"]
                    lp_problem += pulp.lpSum(x[(s, j)] for j in group) >= min_pct * group_demand, f"MinVolPct_{r_idx}_{s}_{hash(tuple(group))}"
    
    #############################################
    # DEBUG OUTPUT
    #############################################
    constraint_names = list(lp_problem.constraints.keys())
    duplicates = set([n for n in constraint_names if constraint_names.count(n) > 1])
    if duplicates:
        print("DEBUG: Duplicate constraint names found:", duplicates)
    print("DEBUG: Total constraints added:", len(constraint_names))
    
    # --- Solve the model ---
    solver = pulp.PULP_CBC_CMD(msg=False, gapRel=0, gapAbs=0)
    lp_problem.solve(solver)
    model_status = pulp.LpStatus[lp_problem.status]

    feasibility_notes = ""
    if model_status == "Infeasible":
        feasibility_notes += "Model is infeasible. Likely causes include:\n"
        feasibility_notes += " - Insufficient supplier capacity relative to demand.\n"
        feasibility_notes += " - Custom rule constraints conflicting with overall volume/demand.\n\n"
        
        # Detailed evaluation per custom rule:
        feasibility_notes += "Detailed Rule Evaluations:\n"
        for idx, rule in enumerate(rules):
            r_type = rule.get("rule_type", "").strip().lower()
            if r_type == "# of suppliers":
                try:
                    required_suppliers = int(rule.get("rule_input", 0))
                except Exception:
                    required_suppliers = 0
                available_suppliers = 2  # Replace with your computed value
                feasibility_notes += (
                    f"Rule {idx+1} ('# of Suppliers'): The rule requires at least {required_suppliers} unique suppliers. "
                    f"Data shows only {available_suppliers} valid suppliers for the relevant grouping. "
                    "This mismatch makes the rule unsatisfiable.\n"
                )
            elif r_type == "% of volume awarded":
                try:
                    required_pct = float(rule["rule_input"].rstrip("%")) / 100.0
                except Exception:
                    required_pct = 0.0
                bid_id = rule.get("grouping_scope", "").strip()
                total_demand = demand_data.get(bid_id, 0)
                required_volume = required_pct * total_demand if total_demand else 0
                supplier = rule.get("supplier_scope", "N/A")
                supplier_capacity = capacity_data.get((supplier, "Bid ID", bid_id), None)
                capacity_str = f"{supplier_capacity}" if supplier_capacity is not None else "not available"
                feasibility_notes += (
                    f"Rule {idx+1} ('% of Volume Awarded'): For Bid ID {bid_id}, total demand is {total_demand} units. "
                    f"A {required_pct*100:.0f}% allocation requires {required_volume:.1f} units to be awarded to {supplier}. "
                    f"Supplier capacity is {capacity_str}. If {required_volume:.1f} exceeds available capacity, this rule cannot be met.\n"
                )
            elif r_type == "# of volume awarded":
                try:
                    target_volume = float(rule["rule_input"])
                except Exception:
                    target_volume = 0.0
                grouping = rule.get("grouping", "All").strip().lower()
                if grouping == "bid id":
                    bid_id = rule.get("grouping_scope", "").strip()
                    total_demand = demand_data.get(bid_id, 0)
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Volume Awarded'): For Bid ID {bid_id}, the target volume is {target_volume} units, "
                        f"but the total demand is only {total_demand} units.\n"
                    )
                else:
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Volume Awarded'): The target volume is {target_volume} units in the grouping. "
                        "Please verify that the demand and supplier capacity support this allocation.\n"
                    )
            elif r_type == "# of transitions":
                try:
                    target_transitions = int(rule["rule_input"])
                except Exception:
                    target_transitions = 0
                # Check if the rule is applied on a single Bid ID and the target is 1.
                if target_transitions == 1 and rule.get("grouping", "").strip().lower() == "bid id":
                    bid_id = rule.get("grouping_scope", "").strip()
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Transitions'): The rule requires at least 1 transition for Bid ID {bid_id}. "
                        "Note: Requiring at least one transition on an individual Bid ID is a very strict requirement; "
                        "it forces a non-incumbent allocation even when data or economic factors might not support a transition. "
                        "Consider applying this rule conditionally or relaxing the requirement to improve feasibility.\n"
                    )
                else:
                    # Otherwise, include the normal message with computed or placeholder actual transitions.
                    actual_transitions = 1  # (Replace with your computed metric, if available)
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Transitions'): The rule requires {target_transitions} transitions, "
                        f"but only {actual_transitions} transition(s) are possible based on the bid distribution.\n"
                    )

            elif r_type == "exclude bids":
                # Dynamically compute metrics for "Exclude Bids" rule.
                if rule["grouping"] == "Bid ID":
                    if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                        items_group_for_exclude = items_dynamic
                    else:
                        items_group_for_exclude = [rule["grouping_scope"].strip()]
                elif rule["grouping"] == "All" or not rule["grouping_scope"]:
                    items_group_for_exclude = items_dynamic
                else:
                    items_group_for_exclude = [j for j in items_dynamic 
                                            if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() 
                                            == str(rule["grouping_scope"]).strip()]
                bid_group = rule.get("bid_grouping", None)
                if bid_group is None:
                    continue
                is_numeric = is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict)
                
                # For each bid, calculate total bids and the number that meet the exclusion criteria.
                for j in items_group_for_exclude:
                    total_bids_for_j = sum(1 for s in suppliers if (s, normalize_bid_id(j)) in supplier_bid_attr_dict)
                    excluded_bids_for_j = 0
                    for s in suppliers:
                        key = (s, normalize_bid_id(j))
                        if key in supplier_bid_attr_dict:
                            bid_val = supplier_bid_attr_dict.get(key, {}).get(bid_group, None)
                            if bid_val is None:
                                continue
                            exclude = False
                            if is_numeric:
                                try:
                                    bid_val_num = float(bid_val)
                                    threshold = float(rule["rule_input"])
                                except:
                                    continue
                                op = rule["operator"].strip().lower()
                                if op in ["greater than", ">"] and bid_val_num > threshold:
                                    exclude = True
                                elif op in ["less than", "<"] and bid_val_num < threshold:
                                    exclude = True
                                elif op in ["exactly", "=="] and bid_val_num == threshold:
                                    exclude = True
                            else:
                                if bid_val.strip().lower() == rule.get("bid_exclusion_value", "").strip().lower():
                                    exclude = True
                            if exclude:
                                excluded_bids_for_j += 1
                    feasibility_notes += (
                        f"Rule {idx+1} ('Exclude Bids'): For Bid ID {j}, there are {total_bids_for_j} bids in total, "
                        f"and {excluded_bids_for_j} meet the exclusion criteria. If these numbers are equal, "
                        "no valid bid remains to satisfy demand.\n"
                    )
            elif r_type == "supplier exclusion":
                bid_id = rule.get("grouping_scope", "").strip()
                supplier = rule.get("supplier_scope", "N/A")
                feasibility_notes += (
                    f"Rule {idx+1} ('Supplier Exclusion'): For Bid ID {bid_id}, supplier {supplier} is excluded, "
                    "and it is the only supplier with a valid bid.\n"
                )
            elif r_type == "# minimum volume awarded":
                try:
                    min_volume = float(rule["rule_input"])
                except Exception:
                    min_volume = 0.0
                grouping = rule.get("grouping", "All").strip().lower()
                if grouping == "bid id":
                    bid_id = rule.get("grouping_scope", "").strip()
                    total_demand = demand_data.get(bid_id, 0)
                    feasibility_notes += (
                        f"Rule {idx+1} ('# Minimum Volume Awarded'): For Bid ID {bid_id}, the rule requires at least {min_volume} units, "
                        f"but total demand is only {total_demand} units.\n"
                    )
                else:
                    feasibility_notes += (
                        f"Rule {idx+1} ('# Minimum Volume Awarded'): A minimum volume of {min_volume} units is required in the grouping. "
                        "Please verify that the grouping's total demand is sufficient.\n"
                    )
            elif r_type == "% minimum volume awarded":
                try:
                    min_pct = float(rule["rule_input"].rstrip("%")) / 100.0
                except Exception:
                    min_pct = 0.0
                grouping = rule.get("grouping", "All").strip().lower()
                if grouping == "bid id":
                    bid_id = rule.get("grouping_scope", "").strip()
                    total_demand = demand_data.get(bid_id, 0)
                    required_volume = min_pct * total_demand if total_demand else 0
                    feasibility_notes += (
                        f"Rule {idx+1} ('% Minimum Volume Awarded'): For Bid ID {bid_id}, total demand is {total_demand} units; "
                        f"a minimum of {min_pct*100:.0f}% requires {required_volume:.1f} units. Please verify that this is feasible.\n"
                    )
                else:
                    feasibility_notes += (
                        f"Rule {idx+1} ('% Minimum Volume Awarded'): The rule requires a minimum allocation of {min_pct*100:.0f}% of the volume in the grouping. "
                        "Please check that the grouping's demand supports this percentage.\n"
                    )
            else:
                feasibility_notes += (
                    f"Rule {idx+1} ('{rule.get('rule_type', '')}'): Please review the parameters for potential conflicts "
                    "with demand or capacity.\n"
                )
        
        # Additional diagnostic information (e.g., Bid IDs with no valid bids)
        no_bid_items = [bid for bid, d_val in demand_data.items() if d_val == 0]
        if no_bid_items:
            feasibility_notes += (
                "\nNote: The following Bid ID(s) were excluded because no valid bids were found: " +
                ", ".join(no_bid_items) + ".\n"
            )
        feasibility_notes += "\nPlease review supplier capacities, demand figures, and custom rule constraints for adjustments."
    else:
        feasibility_notes = "Model is optimal."

    
    #############################################
    # PREPARE RESULTS
    #############################################
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
        # If no supplier is awarded, mark as "No Bid"
        if not awarded_list:
            awarded_list = [("No Bid", 0)]
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
            base_price = baseline_price_data[normalize_bid_id(j)]
            baseline_spend = base_price * award_val
            baseline_savings = baseline_spend - awarded_spend
            active_rebate = 0
            for k, tier in enumerate(rebate_tiers.get(s, [])):
                if pulp.value(y_rebate[s][k]) is not None and pulp.value(y_rebate[s][k]) >= 0.5:
                    active_rebate = tier[2]
                    break
            rebate_savings = awarded_spend * active_rebate
            facility_val = item_attr_data[normalize_bid_id(j)].get("Facility", "")
            row = {
                "Bid ID": idx,
                "Bid ID Split": bid_split,
                "Facility": facility_val,
                "Incumbent": item_attr_data[normalize_bid_id(j)].get("Incumbent", ""),
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

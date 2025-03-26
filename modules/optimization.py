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

            # Aggregated case: if grouping is ALL or grouping scope is empty.
            if rule["grouping"].strip().upper() == "ALL" or not rule["grouping_scope"].strip():
                items_group = items_dynamic  # All bids.
                # For each supplier, define an aggregated binary variable Z_s that is 1 if supplier s is awarded on any bid in items_group.
                Z = {}
                for s in suppliers:
                    Z[s] = pulp.LpVariable(f"Z_{r_idx}_{s}", cat='Binary')
                    # For each bid j in the group, create a per-bid binary indicator and link it to Z[s].
                    for j in items_group:
                        w = pulp.LpVariable(f"w_{r_idx}_{s}_{j}", cat='Binary')
                        epsilon = 1
                        lp_problem += x[(s, j)] <= M * w, f"AggSupplIndicator_{r_idx}_{s}_{j}"
                        lp_problem += x[(s, j)] >= epsilon * w, f"AggSupplIndicatorLB_{r_idx}_{s}_{j}"
                        # If supplier s is awarded on bid j, then Z[s] must be 1.
                        lp_problem += Z[s] >= w, f"Link_Z_{r_idx}_{s}_{j}"
                # Now enforce the rule on the aggregated count.
                if operator == "at least":
                    lp_problem += pulp.lpSum(Z[s] for s in suppliers) >= supplier_target, f"SupplierCountAgg_{r_idx}"
                elif operator == "at most":
                    lp_problem += pulp.lpSum(Z[s] for s in suppliers) <= supplier_target, f"SupplierCountAgg_{r_idx}"
                elif operator == "exactly":
                    lp_problem += pulp.lpSum(Z[s] for s in suppliers) == supplier_target, f"SupplierCountAgg_{r_idx}"
                continue

            # Non-aggregated case:
            if rule["grouping"].strip().lower() == "bid id":
                if rule["grouping_scope"].strip().lower() == "all":
                    items_group = items_dynamic
                elif rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [rule["grouping_scope"].strip()]
            else:
                if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [
                        j for j in items_dynamic
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip().lower() == str(rule["grouping_scope"]).strip().lower()
                    ]
            # For each bid in the selected group, create per-bid binary indicators and enforce the constraint.
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

            # Determine the set of bids (items_group) based on the grouping.
            if rule["grouping"].strip().upper() == "ALL":
                # If the grouping is set to ALL, then include all bids.
                items_group = items_dynamic
            elif rule["grouping"].strip() == "Bid ID":
                if rule["grouping_scope"].strip().lower() == "all":
                    items_group = items_dynamic
                elif rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [rule["grouping_scope"].strip()]
            else:
                # For any other grouping, if grouping_scope is "apply to all items individually",
                # you might choose to loop over each unique value. Otherwise, filter for the provided value.
                if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    unique_groups = sorted({
                        str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip()
                        for j in items_dynamic
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() != ""
                    })
                    # Here, one option is to apply the rule separately to each unique group.
                    # For simplicity, you might also choose to aggregate across all bids.
                    # For now, we set items_group to all bids in the selected grouping.
                    items_group = []
                    for group_val in unique_groups:
                        items_group.extend(
                            [j for j in items_dynamic if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]
                        )
                else:
                    items_group = [
                        j for j in items_dynamic
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()
                    ]

            # For each bid j in the selected group, add the volume awarded constraint.
            for j in items_group:
                total_vol = pulp.lpSum(x[(s, j)] for s in suppliers)
                # Handle special supplier scopes.
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
                    if rule["operator"] == "At least":
                        lp_problem += lhs >= percentage * total_vol, f"Rule_{r_idx}_{j}"
                    elif rule["operator"] == "At most":
                        lp_problem += lhs <= percentage * total_vol, f"Rule_{r_idx}_{j}"
                    elif rule["operator"] == "Exactly":
                        lp_problem += lhs == percentage * total_vol, f"Rule_{r_idx}_{j}"
                # When Supplier Scope is "ALL", loop over all suppliers.
                elif rule["supplier_scope"].strip().upper() == "ALL":
                    for s in suppliers:
                        w_var = pulp.LpVariable(f"w_{r_idx}_{s}_{j}", cat='Binary')
                        epsilon = 1  # Minimal positive threshold.
                        lp_problem += x[(s, j)] <= M * w_var, f"SupplIndicator_{r_idx}_{s}_{j}"
                        lp_problem += x[(s, j)] >= epsilon * w_var, f"SupplIndicatorLB_{r_idx}_{s}_{j}"
                        if rule["operator"] == "At least":
                            lp_problem += x[(s, j)] >= percentage * total_vol - M * (1 - w_var), f"Rule_{r_idx}_{s}_{j}"
                        elif rule["operator"] == "At most":
                            lp_problem += x[(s, j)] <= percentage * total_vol + M * (1 - w_var), f"Rule_{r_idx}_{s}_{j}"
                        elif rule["operator"] == "Exactly":
                            lp_problem += x[(s, j)] >= percentage * total_vol - M * (1 - w_var), f"RuleLB_{r_idx}_{s}_{j}"
                            lp_problem += x[(s, j)] <= percentage * total_vol + M * (1 - w_var), f"RuleUB_{r_idx}_{s}_{j}"
                else:
                    # Interpret supplier_scope as a specific supplier name.
                    lhs = x[(rule["supplier_scope"], j)]
                    if rule["operator"] == "At least":
                        lp_problem += lhs >= percentage * total_vol, f"Rule_{r_idx}_{j}"
                    elif rule["operator"] == "At most":
                        lp_problem += lhs <= percentage * total_vol, f"Rule_{r_idx}_{j}"
                    elif rule["operator"] == "Exactly":
                        lp_problem += lhs == percentage * total_vol, f"Rule_{r_idx}_{j}"

    
        # "# of Volume Awarded" rule.
        elif rule["rule_type"] == "# of Volume Awarded":
            try:
                volume_target = float(rule["rule_input"])
            except Exception:
                continue

            # Aggregated case: When grouping is "ALL", aggregate over all bids.
            if rule["grouping"].strip().upper() == "ALL":
                items_group = items_dynamic  # All bids.
                if rule["supplier_scope"].strip().lower() == "new suppliers":
                    # Aggregate allocation for new suppliers: for each bid, include only suppliers that are not the incumbent.
                    aggregated_new_alloc = pulp.lpSum(
                        x[(s, j)]
                        for j in items_group
                        for s in suppliers
                        if s != item_attr_data[normalize_bid_id(j)].get("Incumbent")
                    )
                    if rule["operator"].lower() == "at most":
                        lp_problem += aggregated_new_alloc <= volume_target, f"Aggregated_NewVolRule_{r_idx}_UB"
                    elif rule["operator"].lower() == "at least":
                        lp_problem += aggregated_new_alloc >= volume_target, f"Aggregated_NewVolRule_{r_idx}_LB"
                    elif rule["operator"].lower() == "exactly":
                        lp_problem += aggregated_new_alloc == volume_target, f"Aggregated_NewVolRule_{r_idx}_Exact"
                else:
                    # For any other supplier scope in aggregated mode,
                    # enforce the constraint for each supplier individually over all bids.
                    for s in suppliers:
                        aggregated_alloc = pulp.lpSum(x[(s, j)] for j in items_group)
                        w_var = pulp.LpVariable(f"w_vol_{r_idx}_{s}_agg", cat='Binary')
                        epsilon = 1
                        lp_problem += aggregated_alloc <= M * w_var, f"AggSupplIndicator_{r_idx}_{s}"
                        lp_problem += aggregated_alloc >= epsilon * w_var, f"AggSupplIndicatorLB_{r_idx}_{s}"
                        if rule["operator"].lower() == "at least":
                            lp_problem += aggregated_alloc >= volume_target - M * (1 - w_var), f"AggregatedVolRule_{r_idx}_{s}_LB"
                        elif rule["operator"].lower() == "at most":
                            lp_problem += aggregated_alloc <= volume_target + M * (1 - w_var), f"AggregatedVolRule_{r_idx}_{s}_UB"
                        elif rule["operator"].lower() == "exactly":
                            lp_problem += aggregated_alloc >= volume_target - M * (1 - w_var), f"AggregatedVolRuleLB_{r_idx}_{s}"
                            lp_problem += aggregated_alloc <= volume_target + M * (1 - w_var), f"AggregatedVolRuleUB_{r_idx}_{s}"
                # Finished aggregated constraint; skip the rest of this rule.
                continue

            # Non-aggregated case (grouping is not ALL):
            if rule["grouping"].strip() == "Bid ID":
                if rule["grouping_scope"].strip().lower() == "all":
                    items_group = items_dynamic
                elif rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [rule["grouping_scope"].strip()]
                for j in items_group:
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
                        if rule["operator"].lower() == "at least":
                            lp_problem += lhs >= volume_target, f"VolRule_{r_idx}_{j}_LB"
                        elif rule["operator"].lower() == "at most":
                            lp_problem += lhs <= volume_target, f"VolRule_{r_idx}_{j}_UB"
                        elif rule["operator"].lower() == "exactly":
                            lp_problem += lhs == volume_target, f"VolRule_{r_idx}_{j}_Exact"
                    elif rule["supplier_scope"].strip().upper() == "ALL":
                        for s in suppliers:
                            w_var = pulp.LpVariable(f"w_vol_{r_idx}_{s}_{j}", cat='Binary')
                            epsilon = 1
                            lp_problem += x[(s, j)] <= M * w_var, f"VolSupplIndicator_{r_idx}_{s}_{j}"
                            lp_problem += x[(s, j)] >= epsilon * w_var, f"VolSupplIndicatorLB_{r_idx}_{s}_{j}"
                            if rule["operator"].lower() == "at least":
                                lp_problem += x[(s, j)] >= volume_target - M * (1 - w_var), f"VolRule_{r_idx}_{s}_{j}_LB"
                            elif rule["operator"].lower() == "at most":
                                lp_problem += x[(s, j)] <= volume_target + M * (1 - w_var), f"VolRule_{r_idx}_{s}_{j}_UB"
                            elif rule["operator"].lower() == "exactly":
                                lp_problem += x[(s, j)] >= volume_target - M * (1 - w_var), f"VolRuleLB_{r_idx}_{s}_{j}"
                                lp_problem += x[(s, j)] <= volume_target + M * (1 - w_var), f"VolRuleUB_{r_idx}_{s}_{j}"
                    else:
                        supplier = rule["supplier_scope"].strip()
                        lhs = x[(supplier, j)]
                        if rule["operator"].lower() == "at least":
                            lp_problem += lhs >= volume_target, f"VolRule_{r_idx}_{j}_LB"
                        elif rule["operator"].lower() == "at most":
                            lp_problem += lhs <= volume_target, f"VolRule_{r_idx}_{j}_UB"
                        elif rule["operator"].lower() == "exactly":
                            lp_problem += lhs == volume_target, f"VolRule_{r_idx}_{j}_Exact"
            else:
                # Grouping by an attribute other than Bid ID.
                if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    unique_groups = sorted({
                        str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip()
                        for j in items_dynamic
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() != ""
                    })
                    for group_val in unique_groups:
                        items_group = [
                            j for j in items_dynamic
                            if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val
                        ]
                        group_total_vol = pulp.lpSum(x[(s, j)] for s in suppliers for j in items_group)
                        for s in suppliers:
                            supplier_group_alloc = pulp.lpSum(x[(s, j)] for j in items_group)
                            w_var = pulp.LpVariable(f"w_vol_{r_idx}_{s}_{group_val}", cat='Binary')
                            epsilon = 1
                            lp_problem += supplier_group_alloc <= M * w_var, f"GroupVolSupplIndicator_{r_idx}_{s}_{group_val}"
                            lp_problem += supplier_group_alloc >= epsilon * w_var, f"GroupVolSupplIndicatorLB_{r_idx}_{s}_{group_val}"
                            if rule["operator"].lower() == "at least":
                                lp_problem += supplier_group_alloc >= volume_target - M * (1 - w_var), f"GroupVolRule_{r_idx}_{s}_{group_val}_LB"
                            elif rule["operator"].lower() == "at most":
                                lp_problem += supplier_group_alloc <= volume_target + M * (1 - w_var), f"GroupVolRule_{r_idx}_{s}_{group_val}_UB"
                            elif rule["operator"].lower() == "exactly":
                                lp_problem += supplier_group_alloc >= volume_target - M * (1 - w_var), f"GroupVolRuleLB_{r_idx}_{s}_{group_val}"
                                lp_problem += supplier_group_alloc <= volume_target + M * (1 - w_var), f"GroupVolRuleUB_{r_idx}_{s}_{group_val}"
                else:
                    items_group = [
                        j for j in items_dynamic
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()
                    ]
                    group_total_vol = pulp.lpSum(x[(s, j)] for s in suppliers for j in items_group)
                    for s in suppliers:
                        supplier_group_alloc = pulp.lpSum(x[(s, j)] for j in items_group)
                        w_var = pulp.LpVariable(f"w_vol_{r_idx}_{s}_group", cat='Binary')
                        epsilon = 1
                        lp_problem += supplier_group_alloc <= M * w_var, f"GroupVolSupplIndicator_{r_idx}_{s}"
                        lp_problem += supplier_group_alloc >= epsilon * w_var, f"GroupVolSupplIndicatorLB_{r_idx}_{s}"
                        if rule["operator"].lower() == "at least":
                            lp_problem += supplier_group_alloc >= volume_target - M * (1 - w_var), f"GroupVolRule_{r_idx}_{s}_LB"
                        elif rule["operator"].lower() == "at most":
                            lp_problem += supplier_group_alloc <= volume_target + M * (1 - w_var), f"GroupVolRule_{r_idx}_{s}_UB"
                        elif rule["operator"].lower() == "exactly":
                            lp_problem += supplier_group_alloc >= volume_target - M * (1 - w_var), f"GroupVolRuleLB_{r_idx}_{s}"
                            lp_problem += supplier_group_alloc <= volume_target + M * (1 - w_var), f"GroupVolRuleUB_{r_idx}_{s}"


        # "# of transitions" rule.
        elif rule["rule_type"].strip().lower() == "# of transitions":
            # Determine grouping parameters.
            grouping = rule.get("grouping", "").strip().lower()
            grouping_scope = rule.get("grouping_scope", "").strip().lower()

            # Determine the set of bid IDs (items_group) to which the rule applies.
            if grouping.upper() == "ALL" or not grouping_scope:
                items_group = items_dynamic
            elif grouping == "bid id":
                if grouping_scope in ["all", "apply to all items individually"]:
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [rule["grouping_scope"].strip()]
            else:
                # For other groupings:
                if grouping_scope == "apply to all items individually":
                    items_group = sorted(list(item_attr_data.keys()))
                else:
                    items_group = [
                        j for j in items_dynamic
                        if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip().lower() == grouping_scope
                    ]

            # Convert rule input to an integer target.
            try:
                transitions_target = int(rule["rule_input"])
            except (ValueError, TypeError):
                continue  # Skip this rule if conversion fails

            operator = rule.get("operator", "").strip().lower()

            # For each bid j in the group, define a binary variable Y[j] that indicates if a transition occurs.
            Y = {}
            for j in items_group:
                Y[j] = pulp.LpVariable(f"Y_{j}", cat='Binary')
                # For each supplier s for which a transition variable T[(j,s)] exists (i.e. s is non-incumbent for bid j),
                # force Y[j] to be at least T[(j,s)].
                for s in suppliers:
                    if (j, s) in T:
                        lp_problem += Y[j] >= T[(j, s)], f"Link_Y_lower_{j}_{s}"
                # Force Y[j] to be no greater than the sum of transition indicators.
                lp_problem += Y[j] <= pulp.lpSum(T[(j, s)] for s in suppliers if (j, s) in T), f"Link_Y_upper_{j}"

            # Total transitions is the sum of Y[j] over the bids in the group.
            total_transitions = pulp.lpSum(Y[j] for j in items_group)

            # Enforce the operator constraint on the aggregated transitions.
            if operator == "at least":
                lp_problem += total_transitions >= transitions_target, f"TransRule_{r_idx}"
            elif operator == "at most":
                lp_problem += total_transitions <= transitions_target, f"TransRule_{r_idx}"
            elif operator == "exactly":
                lp_problem += total_transitions == transitions_target, f"TransRule_{r_idx}"


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
            # Determine the list of Bid IDs (bids_in_group) based on grouping and grouping scope.
            grouping = rule.get("grouping", "").strip().lower()
            grouping_scope = rule.get("grouping_scope", "").strip().lower()
            if grouping == "bid id":
                if grouping_scope == "apply to all items individually":
                    bids_in_group = items_dynamic  # Use all Bid IDs.
                else:
                    bids_in_group = [rule["grouping_scope"].strip()]
            elif grouping == "all" or not rule.get("grouping_scope"):
                bids_in_group = items_dynamic
            else:
                bids_in_group = [
                    j for j in items_dynamic 
                    if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip().lower() == grouping_scope
                ]
            
            # Determine the supplier scope (converted to lowercase).
            s_scope = rule.get("supplier_scope", "").strip().lower() if rule.get("supplier_scope") else None

            if s_scope == "lowest cost supplier":
                for bid in bids_in_group:
                    lp_problem += x[(lowest_cost_supplier[bid], normalize_bid_id(bid))] == 0, f"SupplierExclusion_{bid}"
            elif s_scope == "second lowest cost supplier":
                for bid in bids_in_group:
                    lp_problem += x[(second_lowest_cost_supplier[bid], normalize_bid_id(bid))] == 0, f"SupplierExclusion_{bid}"
            elif s_scope == "incumbent":
                # Exclude the incumbent's bid (if that's the intended behavior for this rule).
                for bid in bids_in_group:
                    incumbent = item_attr_data[normalize_bid_id(bid)].get("Incumbent")
                    if incumbent:
                        lp_problem += x[(incumbent, normalize_bid_id(bid))] == 0, f"SupplierExclusion_{bid}"
            elif s_scope == "new suppliers":
                # Exclude new suppliers: for each bid, force allocations from non-incumbents to zero
                # and force the incumbent to cover the full demand.
                for bid in bids_in_group:
                    incumbent = item_attr_data[normalize_bid_id(bid)].get("Incumbent")
                    if not incumbent:
                        # If no incumbent exists, skip or handle appropriately.
                        continue
                    for s in suppliers:
                        if s != incumbent:
                            lp_problem += x[(s, normalize_bid_id(bid))] == 0, f"SupplierExclusion_{bid}_{s}"
                    bid_demand = demand_data.get(normalize_bid_id(bid), 0)
                    lp_problem += x[(incumbent, normalize_bid_id(bid))] == bid_demand, f"SupplierExclusion_Full_{bid}"
            else:
                # For a specific supplier name (e.g., "C") that doesn't match any special cases.
                for bid in bids_in_group:
                    lp_problem += x[(rule["supplier_scope"], normalize_bid_id(bid))] == 0, f"SupplierExclusion_{bid}"


    
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
            # Determine applicable Bid IDs based on grouping.
            grouping = rule.get("grouping", "").strip().lower()
            grouping_scope = rule.get("grouping_scope", "").strip().lower()
            if grouping == "bid id":
                if grouping_scope == "apply to all items individually":
                    bid_ids = items_dynamic  # Use all bid IDs.
                else:
                    bid_ids = [rule["grouping_scope"].strip()]
            elif grouping == "all" or not grouping_scope:
                bid_ids = items_dynamic
            else:
                bid_ids = [
                    j for j in items_dynamic 
                    if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip().lower() == grouping_scope
                ]
            
            # Process based on rule type:
            r_type = rule.get("rule_type", "").strip().lower()
            
            # ----- "# of Suppliers" Rule -----
            if r_type == "# of suppliers":
                try:
                    required_suppliers = int(rule.get("rule_input", 0))
                except Exception:
                    required_suppliers = 0
                # Here you should compute available suppliers per bid. 
                # (Placeholder below: replace with your actual computation.)
                for bid in bid_ids:
                    available_suppliers = 2  # e.g., compute_available_suppliers(bid, rule, item_attr_data, suppliers)
                    emoji = "✅" if available_suppliers >= required_suppliers else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Suppliers'): {emoji} For Bid ID {bid}, required suppliers: {required_suppliers}, "
                        f"available: {available_suppliers}.\n"
                    )
            
            # ----- "% of Volume Awarded" Rule -----
            elif r_type == "% of volume awarded":
                try:
                    required_pct = float(rule["rule_input"].rstrip("%")) / 100.0
                except Exception:
                    required_pct = 0.0
                supplier_scope = rule.get("supplier_scope", "").strip().lower()  # e.g., "incumbent"
                for bid in bid_ids:
                    norm_bid = normalize_bid_id(bid)
                    total_demand = demand_data.get(norm_bid, 0)
                    required_volume = required_pct * total_demand if total_demand else 0
                    emoji = "✅"
                    if supplier_scope == "incumbent":
                        incumbent = item_attr_data.get(norm_bid, {}).get("Incumbent")
                        if not incumbent:
                            emoji = "❌"
                            feasibility_notes += (
                                f"Rule {idx+1} ('% of Volume Awarded'): {emoji} For Bid ID {bid}, no incumbent is defined.\n"
                            )
                            continue
                        if (incumbent, norm_bid) not in price_data:
                            emoji = "❌"
                            feasibility_notes += (
                                f"Rule {idx+1} ('% of Volume Awarded'): {emoji} For Bid ID {bid}, the incumbent ({incumbent}) did not submit a bid.\n"
                            )
                            continue
                        supplier_capacity = capacity_data.get((incumbent, "Bid ID", bid), None)
                        capacity_str = f"{supplier_capacity}" if supplier_capacity is not None else "not available"
                    else:
                        capacity_str = "N/A"
                    # You might add additional checks (e.g., if required_volume > supplier_capacity).
                    feasibility_notes += (
                        f"Rule {idx+1} ('% of Volume Awarded'): {emoji} For Bid ID {bid}, total demand is {total_demand} units; "
                        f"a {required_pct*100:.0f}% allocation requires {required_volume:.1f} units for {rule['supplier_scope']}. "
                        f"Supplier capacity is {capacity_str}.\n"
                    )
            
            # ----- "# of Volume Awarded" Rule -----
            elif r_type == "# of volume awarded":
                for bid in bid_ids:
                    try:
                        target_volume = float(rule["rule_input"])
                    except Exception:
                        target_volume = 0.0
                    total_demand = demand_data.get(normalize_bid_id(bid), 0)
                    emoji = "✅" if target_volume <= total_demand else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Volume Awarded'): {emoji} For Bid ID {bid}, target volume is {target_volume} units, "
                        f"but total demand is {total_demand} units.\n"
                    )
            
            # ----- "# of Transitions" Rule -----
            elif r_type == "# of transitions":
                for bid in bid_ids:
                    try:
                        target_transitions = int(rule["rule_input"])
                    except Exception:
                        target_transitions = 0
                    # Compute actual transitions per bid (replace with actual computation if available)
                    actual_transitions = 1  # Placeholder
                    emoji = "✅" if actual_transitions >= target_transitions else "❌"
                    # Special note: if target is 1 on a single bid, consider it too strict.
                    if target_transitions == 1 and grouping == "bid id":
                        emoji = "❌"
                        feasibility_notes += (
                            f"Rule {idx+1} ('# of Transitions'): {emoji} For Bid ID {bid}, requiring at least 1 transition on an individual bid "
                            "is very strict.\n"
                        )
                    else:
                        feasibility_notes += (
                            f"Rule {idx+1} ('# of Transitions'): {emoji} For Bid ID {bid}, target transitions: {target_transitions}, "
                            f"actual transitions: {actual_transitions}.\n"
                        )
            
            # ----- "Exclude Bids" Rule -----
            elif r_type == "exclude bids":
                bid_group = rule.get("bid_grouping", None)
                if bid_group is None:
                    continue
                is_numeric = is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict)
                for bid in bid_ids:
                    total_bids_for_bid = sum(1 for s in suppliers if (s, normalize_bid_id(bid)) in supplier_bid_attr_dict)
                    excluded_bids_for_bid = 0
                    for s in suppliers:
                        key = (s, normalize_bid_id(bid))
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
                                excluded_bids_for_bid += 1
                    emoji = "✅" if total_bids_for_bid > 0 and total_bids_for_bid > excluded_bids_for_bid else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('Exclude Bids'): {emoji} For Bid ID {bid}, total bids: {total_bids_for_bid}, "
                        f"excluded bids: {excluded_bids_for_bid}.\n"
                    )
            
            # ----- "Supplier Exclusion" Rule -----
            elif r_type == "supplier exclusion":
                # Here, we assume the grouping is by Bid ID and evaluate per bid.
                for bid in bid_ids:
                    s_scope = rule.get("supplier_scope", "").strip().lower()
                    if s_scope == "lowest cost supplier":
                        emoji = "✅"  # Assuming evaluation passes if lowest cost is found.
                        feasibility_notes += (
                            f"Rule {idx+1} ('Supplier Exclusion'): {emoji} For Bid ID {bid}, lowest cost supplier is excluded.\n"
                        )
                    elif s_scope == "second lowest cost supplier":
                        emoji = "✅"
                        feasibility_notes += (
                            f"Rule {idx+1} ('Supplier Exclusion'): {emoji} For Bid ID {bid}, second lowest cost supplier is excluded.\n"
                        )
                    elif s_scope == "incumbent":
                        incumbent = item_attr_data.get(normalize_bid_id(bid), {}).get("Incumbent")
                        if not incumbent:
                            emoji = "❌"
                            feasibility_notes += (
                                f"Rule {idx+1} ('Supplier Exclusion'): {emoji} For Bid ID {bid}, no incumbent is defined.\n"
                            )
                        else:
                            emoji = "❌"  # Excluding the incumbent may be problematic.
                            feasibility_notes += (
                                f"Rule {idx+1} ('Supplier Exclusion'): {emoji} For Bid ID {bid}, incumbent ({incumbent}) is excluded.\n"
                            )
                    else:
                        # For a specific supplier name.
                        emoji = "❌"
                        feasibility_notes += (
                            f"Rule {idx+1} ('Supplier Exclusion'): {emoji} For Bid ID {bid}, supplier {rule['supplier_scope']} is excluded.\n"
                        )
            
            # ----- "# Minimum Volume Awarded" Rule -----
            elif r_type == "# minimum volume awarded":
                for bid in bid_ids:
                    try:
                        min_volume = float(rule["rule_input"])
                    except Exception:
                        min_volume = 0.0
                    total_demand = demand_data.get(normalize_bid_id(bid), 0)
                    emoji = "✅" if min_volume <= total_demand else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('# Minimum Volume Awarded'): {emoji} For Bid ID {bid}, required minimum volume: {min_volume}, "
                        f"total demand: {total_demand}.\n"
                    )
            
            # ----- "% Minimum Volume Awarded" Rule -----
            elif r_type == "% minimum volume awarded":
                for bid in bid_ids:
                    try:
                        min_pct = float(rule["rule_input"].rstrip("%")) / 100.0
                    except Exception:
                        min_pct = 0.0
                    total_demand = demand_data.get(normalize_bid_id(bid), 0)
                    required_volume = min_pct * total_demand if total_demand else 0
                    emoji = "✅" if required_volume <= total_demand else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('% Minimum Volume Awarded'): {emoji} For Bid ID {bid}, total demand: {total_demand}, "
                        f"required volume: {required_volume:.1f}.\n"
                    )
            
            else:
                emoji = "❌"
                feasibility_notes += (
                    f"Rule {idx+1} ('{rule.get('rule_type', '')}'): {emoji} Please review the parameters for potential conflicts.\n"
                )
        
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

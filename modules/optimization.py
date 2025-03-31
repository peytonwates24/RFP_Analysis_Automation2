import os
import pandas as pd
import pulp
import streamlit as st

# Global constant for Big-M constraints.
M = 1e9
# Tolerance for per-bid constraints (if needed)
TOL = 0.000

#############################################
# REQUIRED COLUMNS for Excel Validation
#############################################
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
# Helper Functions (unchanged)
#############################################
def load_excel_sheets(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for sheet in xls.sheet_names:
        sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
    return sheets

def validate_sheet(df, sheet_name):
    required = REQUIRED_COLUMNS.get(sheet_name, [])
    missing = [col for col in required if col not in df.columns]
    return missing

def normalize_bid_id(bid):
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
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = row.to_dict()
        d[bid].pop("Bid ID", None)
    return d

def df_to_dict_demand(df):
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = row["Demand"]
    return d

def df_to_dict_price(df):
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
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = row["Baseline Price"]
    return d

def df_to_dict_capacity(df):
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
    operator = rule.get("operator", "").capitalize()
    rule_input = rule.get("rule_input", "")
    grouping = rule.get("grouping", "").strip()
    grouping_scope = rule.get("grouping_scope", "").strip()
    supplier_scope = rule.get("supplier_scope", "").strip()
    
    if grouping.upper() == "ALL" or not grouping:
        grouping_text = "ALL Groupings"
    else:
        grouping_text = grouping

    if not grouping_scope or grouping_scope.upper() == "ALL":
        grouping_scope_text = "ALL Groupings"
    else:
        grouping_scope_text = grouping_scope

    rule_type = rule.get("rule_type", "").lower()

    if rule_type == "% of volume awarded":
        if grouping.upper() == "ALL":
            return f"{operator} {rule_input}% of ALL Groupings is awarded to {supplier_scope}."
        else:
            return f"{operator} {rule_input}% of {grouping} in {grouping_scope_text} is awarded to {supplier_scope}."
    elif rule_type == "# of volume awarded":
        if grouping.upper() == "ALL":
            return f"{operator} {rule_input} units awarded across ALL items to {supplier_scope}."
        else:
            return f"{operator} {rule_input} units awarded in {grouping_scope_text} of {grouping} to {supplier_scope}."
    else:
        return str(rule)

def expand_rule_text(rule, item_attr_data):
    grouping_scope_lower = rule.get("grouping_scope", "").strip().lower()
    if grouping_scope_lower == "apply to all items individually":
        grouping = rule.get("grouping", "").strip().lower()
        if grouping == "bid id":
            groups = sorted(item_attr_data.keys())
        else:
            groups = sorted({
                str(item_attr_data[k].get(rule.get("grouping", ""), "")).strip()
                for k in item_attr_data if str(item_attr_data[k].get(rule.get("grouping", ""), "")).strip() != ""
            })
        texts = []
        for i, group in enumerate(groups):
            new_rule = rule.copy()
            new_rule["grouping_scope"] = group
            texts.append(f"{i+1}. {rule_to_text(new_rule)}")
        return "<br>".join(texts)
    else:
        return rule_to_text(rule)

def is_bid_attribute_numeric(bid_group, supplier_bid_attr_dict):
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
    
    # For each Bid ID in demand, if no supplier has a nonzero bid, set its demand to zero.
    for bid in list(demand_data.keys()):
        has_valid_bid = any(price_data.get((s, bid), 0) != 0 for s in suppliers)
        if not has_valid_bid:
            demand_data[bid] = 0

    # --- Build list of Bid IDs (items_dynamic) ---
    items_dynamic = [normalize_bid_id(j) for j in demand_data.keys()]
    
    # --- Define no_bid_items as those with zero demand ---
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
            scope_attr_str = str(scope_attr) if scope_attr is not None else ""
            scope_value_str = str(scope_value) if scope_value is not None else ""
            if (not scope_attr_str.strip()) or (scope_attr_str.strip().upper() == "ALL") or (not scope_value_str.strip()) or (scope_value_str.strip().upper() == "ALL"):
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic)
            else:
                vol_expr = pulp.lpSum(
                    x[(s, j)] 
                    for j in items_dynamic 
                    if item_attr_data[normalize_bid_id(j)].get(scope_attr) == scope_value
                )
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
            scope_attr_str = str(scope_attr) if scope_attr is not None else ""
            scope_value_str = str(scope_value) if scope_value is not None else ""
            if (not scope_attr_str.strip()) or (scope_attr_str.strip().upper() == "ALL") or (not scope_value_str.strip()) or (scope_value_str.strip().upper() == "ALL"):
                vol_expr = pulp.lpSum(x[(s, j)] for j in items_dynamic)
            else:
                vol_expr = pulp.lpSum(
                    x[(s, j)]
                    for j in items_dynamic
                    if item_attr_data[normalize_bid_id(j)].get(scope_attr) == scope_value
                )
            lp_problem += vol_expr >= Rmin * y_rebate[s][k], f"RebateTierMin_{s}_{k}"
            if Rmax < float('inf'):
                lp_problem += vol_expr <= Rmax + M_rebate * (1 - y_rebate[s][k]), f"RebateTierMax_{s}_{k}"
            lp_problem += rebate_var[s] >= Rperc * S[s] - M_rebate * (1 - y_rebate[s][k]), f"RebateTierLower_{s}_{k}"
            lp_problem += rebate_var[s] <= Rperc * S[s] + M_rebate * (1 - y_rebate[s][k]), f"RebateTierUpper_{s}_{k}"
    
    # --- Compute lowest cost and second lowest cost supplier per bid ---
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
        # "# of suppliers" rule (unchanged)
        if rule["rule_type"].lower() == "# of suppliers":
            try:
                supplier_target = int(rule["rule_input"])
            except Exception:
                continue
            operator = rule["operator"].lower()
            if rule["grouping"].strip().upper() == "ALL" or not rule["grouping_scope"].strip():
                items_group = items_dynamic
                Z = {}
                for s in suppliers:
                    Z[s] = pulp.LpVariable(f"Z_{r_idx}_{s}", cat='Binary')
                    for j in items_group:
                        w = pulp.LpVariable(f"w_{r_idx}_{s}_{j}", cat='Binary')
                        epsilon = 1
                        lp_problem += x[(s, j)] <= M * w, f"AggSupplIndicator_{r_idx}_{s}_{j}"
                        lp_problem += x[(s, j)] >= epsilon * w, f"AggSupplIndicatorLB_{r_idx}_{s}_{j}"
                        lp_problem += Z[s] >= w, f"Link_Z_{r_idx}_{s}_{j}"
                if operator == "at least":
                    lp_problem += pulp.lpSum(Z[s] for s in suppliers) >= supplier_target, f"SupplierCountAgg_{r_idx}"
                elif operator == "at most":
                    lp_problem += pulp.lpSum(Z[s] for s in suppliers) <= supplier_target, f"SupplierCountAgg_{r_idx}"
                elif operator == "exactly":
                    lp_problem += pulp.lpSum(Z[s] for s in suppliers) == supplier_target, f"SupplierCountAgg_{r_idx}"
                continue

        # --------------------------------------------------------------------------
        # Revised "% of Volume Awarded" rule.
        elif rule["rule_type"].lower() == "% of volume awarded":
            try:
                percentage = float(rule["rule_input"].rstrip("%")) / 100.0
            except Exception:
                continue

            scope = rule["supplier_scope"].strip().lower()
            operator = rule["operator"].strip().lower()

            # Determine grouping:
            if rule["grouping"].strip().upper() == "ALL":
                group_items = items_dynamic
            elif rule["grouping"].strip().lower() == "bid id":
                group_items = [normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                group_val = rule["grouping_scope"].strip()
                group_items = [j for j in items_dynamic
                               if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]

            if rule["grouping"].strip().upper() == "ALL" or len(group_items) > 1:
                aggregated_total = pulp.lpSum(x[(s, j)] for j in group_items for s in suppliers)
                if scope == "incumbent":
                    aggregated_vol = pulp.lpSum(
                        x[(item_attr_data[normalize_bid_id(j)].get("Incumbent"), j)]
                        for j in group_items if item_attr_data[normalize_bid_id(j)].get("Incumbent") is not None
                    )
                elif scope == "new suppliers":
                    aggregated_vol = pulp.lpSum(
                        pulp.lpSum(x[(s, j)] for s in suppliers if s != item_attr_data[normalize_bid_id(j)].get("Incumbent"))
                        for j in group_items
                    )
                elif scope == "lowest cost supplier":
                    aggregated_vol = pulp.lpSum(x[(lowest_cost_supplier[j], j)] for j in group_items)
                elif scope == "second lowest cost supplier":
                    aggregated_vol = pulp.lpSum(x[(second_lowest_cost_supplier[j], j)] for j in group_items)
                elif scope == "all":
                    for s in suppliers:
                        vol_s = pulp.lpSum(x[(s, j)] for j in group_items)
                        y_s = pulp.LpVariable(f"y_{r_idx}_{s}", cat='Binary')
                        epsilon = 1e-3
                        lp_problem += vol_s <= M * y_s, f"Active_{r_idx}_{s}"
                        lp_problem += vol_s >= epsilon * y_s, f"MinActive_{r_idx}_{s}"
                        if operator == "at least":
                            lp_problem += vol_s >= percentage * aggregated_total - M*(1 - y_s), f"%VolAwarded_Agg_{r_idx}_{s}_LB"
                        elif operator == "at most":
                            lp_problem += vol_s <= percentage * aggregated_total + M*(1 - y_s), f"%VolAwarded_Agg_{r_idx}_{s}_UB"
                        elif operator == "exactly":
                            lp_problem += vol_s >= percentage * aggregated_total - M*(1 - y_s), f"%VolAwarded_Agg_{r_idx}_{s}_EQ_LB"
                            lp_problem += vol_s <= percentage * aggregated_total + M*(1 - y_s), f"%VolAwarded_Agg_{r_idx}_{s}_EQ_UB"
                    continue
                else:
                    supplier = rule["supplier_scope"].strip()
                    aggregated_vol = pulp.lpSum(x[(supplier, j)] for j in group_items)
                if operator == "at least":
                    lp_problem += aggregated_vol >= percentage * aggregated_total, f"%VolAwarded_Agg_{r_idx}_{scope}_LB"
                elif operator == "at most":
                    lp_problem += aggregated_vol <= percentage * aggregated_total, f"%VolAwarded_Agg_{r_idx}_{scope}_UB"
                elif operator == "exactly":
                    lp_problem += aggregated_vol >= percentage * aggregated_total, f"%VolAwarded_Agg_{r_idx}_{scope}_EQ_LB"
                    lp_problem += aggregated_vol <= percentage * aggregated_total, f"%VolAwarded_Agg_{r_idx}_{scope}_EQ_UB"
            else:
                for j in group_items:
                    total_vol = pulp.lpSum(x[(s, j)] for s in suppliers)
                    if scope == "lowest cost supplier":
                        supplier_for_rule = lowest_cost_supplier[j]
                        lhs = x[(supplier_for_rule, j)]
                    elif scope == "second lowest cost supplier":
                        supplier_for_rule = second_lowest_cost_supplier[j]
                        lhs = x[(supplier_for_rule, j)]
                    elif scope == "incumbent":
                        supplier_for_rule = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                        if supplier_for_rule is None:
                            raise ValueError(("Incumbent", j))
                        lhs = x[(supplier_for_rule, j)]
                    elif scope == "new suppliers":
                        incumbent = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                        lhs = pulp.lpSum(x[(s, j)] for s in suppliers if s != incumbent)
                    elif scope == "all":
                        for s in suppliers:
                            w_var = pulp.LpVariable(f"w_%Vol_{r_idx}_{s}_{j}", cat='Binary')
                            epsilon = 1e-3
                            lp_problem += x[(s, j)] <= M * w_var, f"%VolAwarded_{r_idx}_{s}_{j}_Indicator_UB"
                            lp_problem += x[(s, j)] >= epsilon * w_var, f"%VolAwarded_{r_idx}_{s}_{j}_Indicator_LB"
                            if operator == "at least":
                                lp_problem += x[(s, j)] >= percentage * total_vol - M * (1 - w_var), f"%VolAwarded_{r_idx}_{s}_{j}_LB"
                            elif operator == "at most":
                                lp_problem += x[(s, j)] <= percentage * total_vol + M * (1 - w_var), f"%VolAwarded_{r_idx}_{s}_{j}_UB"
                            elif operator == "exactly":
                                lp_problem += x[(s, j)] >= percentage * total_vol - M * (1 - w_var), f"%VolAwarded_{r_idx}_{s}_{j}_EQ_LB"
                                lp_problem += x[(s, j)] <= percentage * total_vol + M * (1 - w_var), f"%VolAwarded_{r_idx}_{s}_{j}_EQ_UB"
                        continue
                    else:
                        supplier = rule["supplier_scope"].strip()
                        lhs = x[(supplier, j)]
                    if operator == "at least":
                        lp_problem += lhs >= percentage * total_vol, f"%VolAwarded_{r_idx}_{j}_LB"
                    elif operator == "at most":
                        lp_problem += lhs <= percentage * total_vol, f"%VolAwarded_{r_idx}_{j}_UB"
                    elif operator == "exactly":
                        lp_problem += lhs >= percentage * total_vol, f"%VolAwarded_{r_idx}_{j}_EQ_LB"
                        lp_problem += lhs <= percentage * total_vol, f"%VolAwarded_{r_idx}_{j}_EQ_UB"
                        
        # --------------------------------------------------------------------------
        # NEW: Revised "# of Volume Awarded" rule.
        elif rule["rule_type"].lower() == "# of volume awarded":
            try:
                volume_target = float(rule["rule_input"])
            except Exception:
                continue

            scope = rule["supplier_scope"].strip().lower()
            operator = rule["operator"].strip().lower()

            # Determine grouping:
            if rule["grouping"].strip().upper() == "ALL":
                group_items = items_dynamic
            elif rule["grouping"].strip().lower() == "bid id":
                group_items = [normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                group_val = rule["grouping_scope"].strip()
                group_items = [j for j in items_dynamic
                               if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip() == group_val]

            # Aggregated formulation: if grouping is ALL or the group contains more than one bid.
            if rule["grouping"].strip().upper() == "ALL" or len(group_items) > 1:
                if scope == "incumbent":
                    aggregated_vol = pulp.lpSum(
                        x[(item_attr_data[normalize_bid_id(j)].get("Incumbent"), j)]
                        for j in group_items if item_attr_data[normalize_bid_id(j)].get("Incumbent") is not None
                    )
                elif scope == "new suppliers":
                    aggregated_vol = pulp.lpSum(
                        pulp.lpSum(x[(s, j)] for s in suppliers if s != item_attr_data[normalize_bid_id(j)].get("Incumbent"))
                        for j in group_items
                    )
                elif scope == "lowest cost supplier":
                    aggregated_vol = pulp.lpSum(x[(lowest_cost_supplier[j], j)] for j in group_items)
                elif scope == "second lowest cost supplier":
                    aggregated_vol = pulp.lpSum(x[(second_lowest_cost_supplier[j], j)] for j in group_items)
                elif scope == "all":
                    # For "All", enforce the rule for every supplier individually.
                    for s in suppliers:
                        vol_s = pulp.lpSum(x[(s, j)] for j in group_items)
                        y_s = pulp.LpVariable(f"y_vol_{r_idx}_{s}", cat='Binary')
                        epsilon = 1e-3
                        lp_problem += vol_s <= M * y_s, f"VolActive_{r_idx}_{s}"
                        lp_problem += vol_s >= epsilon * y_s, f"VolMinActive_{r_idx}_{s}"
                        if operator == "at least":
                            lp_problem += vol_s >= volume_target - M*(1 - y_s), f"VolAwarded_Agg_{r_idx}_{s}_LB"
                        elif operator == "at most":
                            lp_problem += vol_s <= volume_target + M*(1 - y_s), f"VolAwarded_Agg_{r_idx}_{s}_UB"
                        elif operator == "exactly":
                            lp_problem += vol_s >= volume_target - M*(1 - y_s), f"VolAwarded_Agg_{r_idx}_{s}_EQ_LB"
                            lp_problem += vol_s <= volume_target + M*(1 - y_s), f"VolAwarded_Agg_{r_idx}_{s}_EQ_UB"
                    continue
                else:
                    supplier = rule["supplier_scope"].strip()
                    aggregated_vol = pulp.lpSum(x[(supplier, j)] for j in group_items)
                if operator == "at least":
                    lp_problem += aggregated_vol >= volume_target, f"VolAwarded_Agg_{r_idx}_{scope}_LB"
                elif operator == "at most":
                    lp_problem += aggregated_vol <= volume_target, f"VolAwarded_Agg_{r_idx}_{scope}_UB"
                elif operator == "exactly":
                    lp_problem += aggregated_vol >= volume_target, f"VolAwarded_Agg_{r_idx}_{scope}_EQ_LB"
                    lp_problem += aggregated_vol <= volume_target, f"VolAwarded_Agg_{r_idx}_{scope}_EQ_UB"
            else:
                # For a single bid, apply constraint per bid.
                for j in group_items:
                    if scope == "lowest cost supplier":
                        supplier_for_rule = lowest_cost_supplier[j]
                        lhs = x[(supplier_for_rule, j)]
                    elif scope == "second lowest cost supplier":
                        supplier_for_rule = second_lowest_cost_supplier[j]
                        lhs = x[(supplier_for_rule, j)]
                    elif scope == "incumbent":
                        supplier_for_rule = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                        if supplier_for_rule is None:
                            raise ValueError(("Incumbent", j))
                        lhs = x[(supplier_for_rule, j)]
                    elif scope == "new suppliers":
                        incumbent = item_attr_data[normalize_bid_id(j)].get("Incumbent")
                        lhs = pulp.lpSum(x[(s, j)] for s in suppliers if s != incumbent)
                    elif scope == "all":
                        for s in suppliers:
                            w_var = pulp.LpVariable(f"w_vol_{r_idx}_{s}_{j}", cat='Binary')
                            epsilon = 1e-3
                            lp_problem += x[(s, j)] <= M * w_var, f"VolAwarded_{r_idx}_{s}_{j}_Indicator_UB"
                            lp_problem += x[(s, j)] >= epsilon * w_var, f"VolAwarded_{r_idx}_{s}_{j}_Indicator_LB"
                            if operator == "at least":
                                lp_problem += x[(s, j)] >= volume_target - M*(1 - w_var), f"VolAwarded_{r_idx}_{s}_{j}_LB"
                            elif operator == "at most":
                                lp_problem += x[(s, j)] <= volume_target + M*(1 - w_var), f"VolAwarded_{r_idx}_{s}_{j}_UB"
                            elif operator == "exactly":
                                lp_problem += x[(s, j)] >= volume_target - M*(1 - w_var), f"VolAwarded_{r_idx}_{s}_{j}_EQ_LB"
                                lp_problem += x[(s, j)] <= volume_target + M*(1 - w_var), f"VolAwarded_{r_idx}_{s}_{j}_EQ_UB"
                        continue
                    else:
                        supplier = rule["supplier_scope"].strip()
                        lhs = x[(supplier, j)]
                    if operator == "at least":
                        lp_problem += lhs >= volume_target, f"VolAwarded_{r_idx}_{j}_LB"
                    elif operator == "at most":
                        lp_problem += lhs <= volume_target, f"VolAwarded_{r_idx}_{j}_UB"
                    elif operator == "exactly":
                        lp_problem += lhs >= volume_target, f"VolAwarded_{r_idx}_{j}_EQ_LB"
                        lp_problem += lhs <= volume_target, f"VolAwarded_{r_idx}_{j}_EQ_UB"

        # Other custom rule branches remain unchanged...
    
    #############################################
    # DEBUG OUTPUT
    #############################################
    constraint_names = list(lp_problem.constraints.keys())
    duplicates = set([n for n in constraint_names if constraint_names.count(n) > 1])
    if duplicates:
        print("DEBUG: Duplicate constraint names found:", duplicates)
    print("DEBUG: Total constraints added:", len(constraint_names))
    
    solver = pulp.PULP_CBC_CMD(msg=False, gapRel=0, gapAbs=0)
    lp_problem.solve(solver)
    model_status = pulp.LpStatus[lp_problem.status]

    feasibility_notes = ""
    if model_status == "Infeasible":
        feasibility_notes += "Model is infeasible. Likely causes include:\n"
        feasibility_notes += " - Insufficient supplier capacity relative to demand.\n"
        feasibility_notes += " - Custom rule constraints conflicting with overall volume/demand.\n\n"
        feasibility_notes += "Detailed Rule Evaluations:\n"
        for idx, rule in enumerate(rules):
            grouping = rule.get("grouping", "").strip().lower()
            grouping_scope = rule.get("grouping_scope", "").strip().lower()
            if grouping == "bid id":
                if grouping_scope == "apply to all items individually":
                    bid_ids = items_dynamic
                else:
                    bid_ids = [rule["grouping_scope"].strip()]
            elif grouping == "all" or not rule.get("grouping_scope"):
                bid_ids = items_dynamic
            else:
                bid_ids = [
                    j for j in items_dynamic 
                    if str(item_attr_data[normalize_bid_id(j)].get(rule["grouping"], "")).strip().lower() == grouping_scope
                ]
            r_type = rule.get("rule_type", "").strip().lower()
            if r_type == "# of suppliers":
                try:
                    required_suppliers = int(rule.get("rule_input", 0))
                except Exception:
                    required_suppliers = 0
                for bid in bid_ids:
                    available_suppliers = 2  # Placeholder for actual computation.
                    emoji = "✅" if available_suppliers >= required_suppliers else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Suppliers'): {emoji} For Bid ID {bid}, required suppliers: {required_suppliers}, "
                        f"available: {available_suppliers}.\n"
                    )
            elif r_type == "% of volume awarded":
                try:
                    required_pct = float(rule["rule_input"].rstrip("%")) / 100.0
                except Exception:
                    required_pct = 0.0
                for bid in bid_ids:
                    norm_bid = normalize_bid_id(bid)
                    total_demand = demand_data.get(norm_bid, 0)
                    required_volume = required_pct * total_demand if total_demand else 0
                    emoji = "✅"
                    if rule["supplier_scope"].strip().lower() == "incumbent":
                        incumbent = item_attr_data.get(norm_bid, {}).get("Incumbent")
                        if not incumbent or (incumbent, norm_bid) not in price_data:
                            emoji = "❌"
                            feasibility_notes += (
                                f"Rule {idx+1} ('% of Volume Awarded'): {emoji} For Bid ID {bid}, incumbent issue.\n"
                            )
                            continue
                    feasibility_notes += (
                        f"Rule {idx+1} ('% of Volume Awarded'): {emoji} For Bid ID {bid}, total demand: {total_demand} units; "
                        f"a {required_pct*100:.0f}% allocation requires {required_volume:.1f} units for {rule['supplier_scope']}.\n"
                    )
            elif r_type == "# of volume awarded":
                try:
                    vol_target = float(rule["rule_input"])
                except Exception:
                    vol_target = 0.0
                for bid in bid_ids:
                    norm_bid = normalize_bid_id(bid)
                    total_demand = demand_data.get(norm_bid, 0)
                    emoji = "✅" if vol_target <= total_demand else "❌"
                    feasibility_notes += (
                        f"Rule {idx+1} ('# of Volume Awarded'): {emoji} For Bid ID {bid}, total demand: {total_demand} units; "
                        f"a target volume of {vol_target:.1f} units is specified for {rule['supplier_scope']}.\n"
                    )
            else:
                emoji = "❌"
                feasibility_notes += (
                    f"Rule {idx+1} ('{rule.get('rule_type', '')}'): {emoji} Please review the parameters.\n"
                )
        no_bid_items = [bid for bid, d_val in demand_data.items() if d_val == 0]
        if no_bid_items:
            feasibility_notes += (
                "\nNote: The following Bid ID(s) were excluded because no valid bids were found: " +
                ", ".join(no_bid_items) + ".\n"
            )
        feasibility_notes += "\nReview supplier capacities, demand, and custom rule constraints for adjustments."
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

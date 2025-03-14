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
    try:
        # Convert to float first in case bid is a numeric type.
        num = float(bid)
        # If the number is an integer, convert it to an integer to avoid a trailing .0
        if num.is_integer():
            return str(int(num))
        else:
            return str(num)
    except Exception:
        return str(bid).strip()

def df_to_dict_item_attributes(df):
    """
    Convert the "Item Attributes" sheet into a dictionary keyed by Bid ID.
    Only "Bid ID" and "Incumbent" are required; all other columns are read in.
    """
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
    """
    Convert the "Baseline Price" sheet into a dictionary keyed by Bid ID with baseline prices.
    """
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])  # Use normalization here
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
        bid = normalize_bid_id(row["Bid ID"])  # Use normalization here
        attr = row.to_dict()
        attr.pop("Supplier Name", None)
        attr.pop("Bid ID", None)
        d[(supplier, bid)] = attr
    return d

#############################################
# Helper for Custom Rule Text Representation
#############################################
def rule_to_text(rule):
    # Retrieve grouping and supplier information.
    grouping = rule.get("grouping", "all items")
    grouping_scope = rule.get("grouping_scope", "all items")
    supplier = rule.get("supplier_scope")
    if supplier is None:
        supplier = "All"
    op = rule.get("operator", "").lower()

    # If grouping is "Bid ID", ensure we output "Bid ID <value>".
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
    If the grouping scope is set to "Apply to all items individually",
    this function expands the rule text for each unique grouping value.
    Otherwise, it just returns the standard rule text.
    """
    grouping = rule.get("grouping", "All")
    grouping_scope = rule.get("grouping_scope", "").strip().lower()
    if grouping_scope == "apply to all items individually":
        # If grouping is Bid ID, use all Bid IDs;
        # otherwise, gather unique grouping values from the item attributes.
        if grouping == "Bid ID":
            groups = sorted(item_attr_data.keys())
        else:
            groups = sorted(
                set(
                    str(item_attr_data[j].get(grouping, "")).strip()
                    for j in item_attr_data
                    if str(item_attr_data[j].get(grouping, "")).strip() != ""
                )
            )
        texts = []
        for i, group in enumerate(groups):
            new_rule = rule.copy()
            new_rule["grouping_scope"] = group
            texts.append(f"{i+1}. {rule_to_text(new_rule)}")
        # Return each expanded rule on its own line using a HTML line break.
        return "<br>".join(texts)
    else:
        return rule_to_text(rule)

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
    
    # Force volume to 0 for supplier-bid pairs that did not submit a valid bid.
    for s in suppliers:
        for j in items_dynamic:
            if (s, j) not in price_data:
                lp_problem += x[(s, j)] == 0, f"NonBid_{s}_{j}"
    
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
        lp_problem += S0[s] == pulp.lpSum(price_data.get((s, j), 0) * x[(s, j)] for j in items_dynamic), f"BaseSpend_{s}"
        lp_problem += V[s] == pulp.lpSum(x[(s, j)] for j in items_dynamic), f"Volume_{s}"
    
    # Compute supplier-specific Big-M values.
    max_price_val = max(price_data.values()) if price_data else 0
    U_spend = {}
    for s in suppliers:
        total_cap = sum(cap for ((sup, _, _), cap) in capacity_data.items() if sup == s)
        U_spend[s] = total_cap * max_price_val
    

    # Discount tiers
    z_discount = {}
    for s in suppliers:
        tiers = discount_tiers.get(s, [])
        if tiers:  # Only add tier variables and constraint if there are tiers
            z_discount[s] = {k: pulp.LpVariable(f"z_discount_{s}_{k}", cat='Binary') for k in range(len(tiers))}
            lp_problem += pulp.lpSum(z_discount[s][k] for k in range(len(tiers))) == 1, f"DiscountTierSelect_{s}"
        else:
            # No discount tiers for this supplier; assign an empty dictionary.
            z_discount[s] = {}

    # Rebate tiers
    y_rebate = {}
    for s in suppliers:
        tiers = rebate_tiers.get(s, [])
        if tiers:  # Only add tier variables and constraint if there are tiers
            y_rebate[s] = {k: pulp.LpVariable(f"y_rebate_{s}_{k}", cat='Binary') for k in range(len(tiers))}
            lp_problem += pulp.lpSum(y_rebate[s][k] for k in range(len(tiers))) == 1, f"RebateTierSelect_{s}"
        else:
            y_rebate[s] = {}

    # Define adjustment variables for discounts and rebates.
    d = {s: pulp.LpVariable(f"d_{s}", lowBound=0, cat='Continuous') for s in suppliers}
    rebate_var = {s: pulp.LpVariable(f"rebate_{s}", lowBound=0, cat='Continuous') for s in suppliers}

    # FIX: For suppliers with no discount or rebate tiers, fix the adjustment variables to 0.
    for s in suppliers:
        if not discount_tiers.get(s, []):  # If no discount tiers are provided
            lp_problem += d[s] == 0, f"Fix_d_{s}"
        if not rebate_tiers.get(s, []):    # If no rebate tiers are provided
            lp_problem += rebate_var[s] == 0, f"Fix_rebate_{s}"

    # Now add the objective.
    lp_problem += pulp.lpSum(S[s] - rebate_var[s] for s in suppliers), "Total_Effective_Cost"

    # Discount Tier constraints.
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
                lp_problem += vol_expr <= Dmax + M_discount * (1 - z_discount[s][k]), f"DiscountTierMax_{s}_{k}"
            lp_problem += d[s] >= Dperc * S0[s] - M_discount * (1 - z_discount[s][k]), f"DiscountTierLower_{s}_{k}"
            lp_problem += d[s] <= Dperc * S0[s] + M_discount * (1 - z_discount[s][k]), f"DiscountTierUpper_{s}_{k}"

    for s in suppliers:
        lp_problem += S[s] == S0[s] - d[s], f"EffectiveSpend_{s}"

    # Rebate Tier constraints.
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
                lp_problem += vol_expr <= Rmax + M_rebate * (1 - y_rebate[s][k]), f"RebateTierMax_{s}_{k}"
            lp_problem += rebate_var[s] >= Rperc * S[s] - M_rebate * (1 - y_rebate[s][k]), f"RebateTierLower_{s}_{k}"
            lp_problem += rebate_var[s] <= Rperc * S[s] + M_rebate * (1 - y_rebate[s][k]), f"RebateTierUpper_{s}_{k}"

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
                    unique_groups = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip() 
                                              for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip()})
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        subgroup = [group_val]
                    else:
                        subgroup = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == group_val]
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
            except Exception as e:
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
                            supplier_for_rule = item_attr_data[j].get("Incumbent")
                            lhs = x[(supplier_for_rule, j)]
                        elif rule["supplier_scope"] == "New Suppliers":
                            incumbent = item_attr_data[j].get("Incumbent")
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
                        items_group = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip()
                                              for j in item_attr_data if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""})
                else:
                    items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
                if rule["supplier_scope"] in ["Lowest cost supplier", "Second Lowest Cost Supplier", "Incumbent"]:
                    if rule["supplier_scope"] == "Lowest cost supplier":
                        lhs = pulp.lpSum(x[(lowest_cost_supplier[j], j)] for j in items_group)
                    elif rule["supplier_scope"] == "Second Lowest Cost Supplier":
                        lhs = pulp.lpSum(x[(second_lowest_cost_supplier[j], j)] for j in items_group)
                    elif rule["supplier_scope"] == "Incumbent":
                        lhs = pulp.lpSum(x[(item_attr_data[j].get("Incumbent"), j)] for j in items_group)
                elif rule["supplier_scope"] == "New Suppliers":
                    lhs = pulp.lpSum(pulp.lpSum(x[(s, j)] for s in suppliers if s != item_attr_data[j].get("Incumbent")) for j in items_group)
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
            if rule["grouping"] == "All" or not rule["grouping_scope"]:
                items_group = items_dynamic
            elif rule["grouping"] == "Bid ID":
                if rule["grouping_scope"].strip().lower() == "apply to all items individually":
                    items_group = items_dynamic
                else:
                    items_group = [rule["grouping_scope"]]
            else:
                items_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
            try:
                volume_target = float(rule["rule_input"])
            except Exception:
                continue
            operator = rule["operator"].lower()
            supplier_scope_value = rule["supplier_scope"]
            if supplier_scope_value is None or supplier_scope_value.lower() == "all":
                for j in items_group:
                    for s in suppliers:
                        y = pulp.LpVariable(f"award_indicator_{r_idx}_{s}_{j}", cat='Binary')
                        lp_problem += x[(s, j)] <= float(demand_data[j]) * y, f"MinVolAwarded_UB_{r_idx}_{s}_{j}"
                        if operator == "at least":
                            lp_problem += x[(s, j)] >= volume_target * y, f"MinVolAwarded_LB_{r_idx}_{s}_{j}"
                        elif operator == "exactly":
                            lp_problem += x[(s, j)] == volume_target * y, f"MinVolAwarded_Exact_{r_idx}_{s}_{j}"
                        elif operator == "at most":
                            lp_problem += x[(s, j)] <= volume_target + M * (1 - y), f"MinVolAwarded_AtMost_{r_idx}_{s}_{j}"
            else:
                for j in items_group:
                    y = pulp.LpVariable(f"award_indicator_{r_idx}_{supplier_scope_value}_{j}", cat='Binary')
                    lp_problem += x[(supplier_scope_value, j)] <= float(demand_data[j]) * y, f"MinVolAwarded_{r_idx}_{supplier_scope_value}_UB_{j}"
                    if operator == "at least":
                        lp_problem += x[(supplier_scope_value, j)] >= volume_target * y, f"MinVolAwarded_{r_idx}_{supplier_scope_value}_LB_{j}"
                    elif operator == "exactly":
                        lp_problem += x[(supplier_scope_value, j)] == volume_target * y, f"MinVolAwarded_{r_idx}_{supplier_scope_value}_Exact_{j}"
                    elif operator == "at most":
                        lp_problem += x[(supplier_scope_value, j)] <= volume_target + M * (1 - y), f"MinVolAwarded_{r_idx}_{supplier_scope_value}_AtMost_{j}"
    
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
            if rule["grouping"].strip().lower() == "all" or not rule["grouping_scope"]:
                bids_in_group = items_dynamic
            elif rule["grouping_scope"].strip().lower() == "apply to all items individually":
                if rule["grouping"].strip().lower() == "bid id":
                    unique_groups = sorted(list(item_attr_data.keys()))
                else:
                    unique_groups = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip()
                                            for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip()})
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        bids_in_group = [group_val]
                    else:
                        bids_in_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == group_val]
                    s_scope = rule["supplier_scope"].strip().lower() if rule["supplier_scope"] else None
                    if s_scope == "lowest cost supplier":
                        for j in bids_in_group:
                            lp_problem += x[(lowest_cost_supplier[j], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                    elif s_scope == "second lowest cost supplier":
                        for j in bids_in_group:
                            lp_problem += x[(second_lowest_cost_supplier[j], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                    elif s_scope == "incumbent":
                        for j in bids_in_group:
                            incumbent = item_attr_data[j].get("Incumbent")
                            if incumbent:
                                lp_problem += x[(incumbent, j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                    elif s_scope == "new suppliers":
                        for j in bids_in_group:
                            incumbent = item_attr_data[j].get("Incumbent")
                            for s in suppliers:
                                if s != incumbent:
                                    lp_problem += x[(s, j)] == 0, f"SupplierExclusion_{r_idx}_{j}_{s}"
                    else:
                        for j in bids_in_group:
                            lp_problem += x[(rule["supplier_scope"], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                continue
            else:
                if rule["grouping"].strip().lower() == "bid id":
                    bids_in_group = [rule["grouping_scope"].strip()]
                else:
                    bids_in_group = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]
                s_scope = rule["supplier_scope"].strip().lower() if rule["supplier_scope"] else None
                if s_scope == "lowest cost supplier":
                    for j in bids_in_group:
                        lp_problem += x[(lowest_cost_supplier[j], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                elif s_scope == "second lowest cost supplier":
                    for j in bids_in_group:
                        lp_problem += x[(second_lowest_cost_supplier[j], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                elif s_scope == "incumbent":
                    for j in bids_in_group:
                        incumbent = item_attr_data[j].get("Incumbent")
                        if incumbent:
                            lp_problem += x[(incumbent, j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
                elif s_scope == "new suppliers":
                    for j in bids_in_group:
                        incumbent = item_attr_data[j].get("Incumbent")
                        for s in suppliers:
                            if s != incumbent:
                                lp_problem += x[(s, j)] == 0, f"SupplierExclusion_{r_idx}_{j}_{s}"
                else:
                    for j in bids_in_group:
                        lp_problem += x[(rule["supplier_scope"], j)] == 0, f"SupplierExclusion_{r_idx}_{j}"
    
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
                    unique_groups = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip() for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""})
                items_group_list = []
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        items_group_list.append((group_val, [group_val]))
                    else:
                        group_items = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == group_val]
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
                    group_items = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == group_val]
                    items_group_list = [(group_val, group_items)] if group_items else []
            for group_val, items_group in items_group_list:
                group_demand = sum(float(demand_data[j]) for j in items_group)
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
                    unique_groups = sorted({str(item_attr_data[j].get(rule["grouping"], "")).strip() for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() != ""})
                group_list = []
                for group_val in unique_groups:
                    if rule["grouping"].strip().lower() == "bid id":
                        group_list.append([group_val])
                    else:
                        group_items = [j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == group_val]
                        if group_items:
                            group_list.append(group_items)
            else:
                if rule["grouping"].strip().lower() == "all" or not rule.get("grouping_scope"):
                    group_list = [items_dynamic]
                elif rule["grouping"].strip().lower() == "bid id":
                    group_list = [[str(rule["grouping_scope"]).strip()]]
                else:
                    group_list = [[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"], "")).strip() == str(rule["grouping_scope"]).strip()]]
            for group in group_list:
                if not group:
                    continue
                try:
                    group_demand = sum(float(demand_data[j]) for j in group)
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

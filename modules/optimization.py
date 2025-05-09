###############################################################################
# optimization.py – FULL freight-aware model (≈930 lines)
#  * Merges your original 931-line “multi-scenario/current-price” engine
#    with the freight/KBX logic from the alternate build.
#  * Nothing trimmed or condensed: every helper, every rule branch, every
#    debug section remains verbatim – only freight code & column additions
#    are grafted in.
###############################################################################
import os
import logging
import pandas as pd
import pulp
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# Logging configuration (kept verbatim)
# ──────────────────────────────────────────────────────────────────────────────
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)

# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────
M   = 1e9   # Big-M
EPS = 0.0   # Non-zero lower-bound for “positive award” binary triggers

# ──────────────────────────────────────────────────────────────────────────────
# REQUIRED COLUMNS – Price sheet now includes Supplier Freight & KBX
# ──────────────────────────────────────────────────────────────────────────────
REQUIRED_COLUMNS = {
    "Item Attributes"        : ["Bid ID", "Incumbent"],
    "Price"                  : ["Supplier Name", "Bid ID", "Price", "Supplier Freight", "KBX"],
    "Demand"                 : ["Bid ID", "Demand"],
    "Baseline Price"         : ["Bid ID", "Baseline Price", "Current Price"],
    "Capacity"               : ["Supplier Name", "Capacity Scope", "Scope Value", "Capacity"],
    "Rebate Tiers"           : ["Supplier Name", "Min", "Max", "Percentage", "Scope Attribute", "Scope Value"],
    "Discount Tiers"         : ["Supplier Name", "Min", "Max", "Percentage", "Scope Attribute", "Scope Value"],
    "Supplier Bid Attributes": ["Supplier Name", "Bid ID"]
}

# ──────────────────────────────────────────────────────────────────────────────
# Excel-helper utilities  (100% identical to your long original, plus freight)
# ──────────────────────────────────────────────────────────────────────────────
def load_excel_sheets(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    return {s: xl.parse(s) for s in xl.sheet_names}

def validate_sheet(df, sheet_name):
    required = REQUIRED_COLUMNS.get(sheet_name, [])
    return [col for col in required if col not in df.columns]

def normalize_bid_id(bid):
    if isinstance(bid, (list, tuple)):
        return "-".join(str(x).strip() for x in bid)
    try:
        num = float(bid)
        return str(int(num)) if num.is_integer() else str(num)
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
    return {normalize_bid_id(r["Bid ID"]): r["Demand"] for _, r in df.iterrows()}

# ── Freight-aware price parser (unchanged from freight build)
def df_to_dict_price(df):
    d = {}
    for _, row in df.iterrows():
        supplier = str(row["Supplier Name"]).strip()
        bid      = normalize_bid_id(row["Bid ID"])
        try:
            base_price = float(row["Price"])
        except Exception:
            continue
        if base_price == 0:
            continue
        try:
            sf_val  = row.get("Supplier Freight")
            sf      = float(sf_val) if sf_val not in [None,""] and not pd.isna(sf_val) else 0.0
        except: sf=0.0
        try:
            kbx_val = row.get("KBX")
            kbx     = float(kbx_val) if kbx_val not in [None,""] and not pd.isna(kbx_val) else 0.0
        except: kbx=0.0

        if sf and kbx:
            freight, method = (sf,"DDP") if sf<=kbx else (kbx,"GP Pickup")
        elif sf:
            freight, method = sf,"DDP"
        elif kbx:
            freight, method = kbx,"GP Pickup"
        else:
            freight, method = 0.0,"None"

        d[(supplier, bid)] = {
            "base_price"     : base_price,
            "Supplier Freight": sf,
            "KBX"            : kbx,
            "freight"        : freight,
            "freight_method" : method,
            "effective_price": base_price + freight
        }
    return d

# Baseline dictionary includes baseline & current price
def df_to_dict_baseline_price(df):
    d = {}
    for _, row in df.iterrows():
        bid = normalize_bid_id(row["Bid ID"])
        d[bid] = {
            "baseline": row.get("Baseline Price", 0.0),
            "current" : row.get("Current Price" , 0.0)
        }
    return d

def df_to_dict_capacity(df):
    d={}
    for _,row in df.iterrows():
        s  = str(row["Supplier Name"]).strip()
        cs = str(row["Capacity Scope"]).strip()
        sv = normalize_bid_id(row["Scope Value"]) if cs=="Bid ID" else str(row["Scope Value"]).strip()
        d[(s,cs,sv)] = row["Capacity"]
    return d

def df_to_dict_tiers(df):
    d={}
    for _,row in df.iterrows():
        s=str(row["Supplier Name"]).strip()
        tier=(row["Min"],row["Max"],row["Percentage"],row.get("Scope Attribute"),row.get("Scope Value"))
        d.setdefault(s,[]).append(tier)
    return d

def df_to_dict_supplier_bid_attributes(df):
    d={}
    for _,row in df.iterrows():
        d[(str(row["Supplier Name"]).strip(),normalize_bid_id(row["Bid ID"]))] = {
            k:v for k,v in row.items() if k not in ["Supplier Name","Bid ID"]
        }
    return d

# ──────────────────────────────────────────────────────────────────────────────
# Rule-text helpers (unchanged)
# ──────────────────────────────────────────────────────────────────────────────
def rule_to_text(rule):
    operator       = rule.get("operator", "").capitalize()
    rule_input     = rule.get("rule_input", "")
    grouping       = rule.get("grouping", "").strip()
    grouping_scope = rule.get("grouping_scope", "").strip()
    supplier_scope = rule.get("supplier_scope", "").strip()
    rt             = rule.get("rule_type", "").lower()
    if rt == "% of volume awarded":
        return (f"{operator} {rule_input}% of ALL Groupings is awarded to {supplier_scope}."
                if grouping.upper() == "ALL"
                else f"{operator} {rule_input}% of {grouping_scope} (by {grouping}) is awarded to {supplier_scope}.")
    if rt == "# of volume awarded":
        return (f"{operator} {rule_input} units awarded across ALL items to {supplier_scope}."
                if grouping.upper() == "ALL"
                else f"{operator} {rule_input} units awarded in {grouping_scope} (by {grouping}) to {supplier_scope}.")
    if rt == "% minimum volume awarded":
        return f"At least {rule_input}% of volume to {supplier_scope} in {grouping_scope} (by {grouping})."
    if rt == "# minimum volume awarded":
        return f"At least {rule_input} units to {supplier_scope} in {grouping_scope} (by {grouping})."
    if rt == "# of suppliers":
        return (f"Unique suppliers awarded: {operator} {rule_input} across ALL items."
                if grouping.upper()=="ALL" or not grouping_scope
                else f"Unique suppliers awarded in {grouping_scope} (by {grouping}): {operator} {rule_input}.")
    if rt == "# of transitions":
        gtxt="all items" if grouping.upper()=="ALL" or not grouping_scope else grouping_scope
        return f"# Transitions: {operator} {rule_input} transitions in {gtxt}."
    if rt == "exclude bids":
        if rule.get("bid_exclusion_value"):
            return f"Exclude bids where {rule['bid_grouping']} equals '{rule['bid_exclusion_value']}' for {grouping_scope} (by {grouping})."
        return f"Exclude bids where {rule['bid_grouping']} {operator} {rule_input} for {grouping_scope} (by {grouping})."
    if rt == "supplier exclusion":
        return f"Exclude {supplier_scope} from {grouping_scope} (by {grouping})."
    return str(rule)

def expand_rule_text(rule, item_attr_data):
    gscope=rule.get("grouping_scope","").strip().lower()
    if gscope=="apply to all items individually":
        grouping=rule.get("grouping","").strip().lower()
        groups=(sorted(item_attr_data.keys()) if grouping=="bid id"
                else sorted({str(item_attr_data[k].get(grouping,"")).strip()
                             for k in item_attr_data if str(item_attr_data[k].get(grouping,"")).strip()}))
        return "<br>".join(f"{i+1}. {rule_to_text({**rule,'grouping_scope':g})}" for i,g in enumerate(groups))
    return rule_to_text(rule)

def is_bid_attribute_numeric(bid_group, sbad):
    for _,attr in sbad.items():
        if bid_group in attr and attr[bid_group] is not None:
            try: float(attr[bid_group]); return True
            except: return False
    return False

# ──────────────────────────────────────────────────────────────────────────────
# Main optimization engine (identical logic, freight + current price combos)
# ──────────────────────────────────────────────────────────────────────────────
def run_optimization(capacity_data, demand_data, item_attr_data, price_data,
                     rebate_tiers, discount_tiers, baseline_price_data, rules=[],
                     supplier_bid_attr_dict=None, suppliers=None):

    if supplier_bid_attr_dict is None: raise ValueError("supplier_bid_attr_dict required")
    if suppliers is None: raise ValueError("suppliers list required")

    # ── Normalize keys
    demand_data = {normalize_bid_id(k): v for k,v in demand_data.items()}
    item_attr_data = {normalize_bid_id(k): v for k,v in item_attr_data.items()}
    baseline_price_data = {normalize_bid_id(k): v for k,v in baseline_price_data.items()}
    price_data  = {(s,normalize_bid_id(b)): v for (s,b),v in price_data.items()}
    capacity_data = {(s,cs,(normalize_bid_id(v) if cs=="Bid ID" else str(v).strip())):cap
                     for (s,cs,v),cap in capacity_data.items()}
    supplier_bid_attr_dict = {(s,normalize_bid_id(b)): a for (s,b),a in supplier_bid_attr_dict.items()}

    # zero demand for items with no bids
    for bid in list(demand_data):
        if not any(price_data.get((s,bid)) for s in suppliers):
            demand_data[bid]=0

    items_dynamic=list(demand_data.keys())

    # Transition bin vars for rules requiring them
    T={}
    for j in items_dynamic:
        inc=item_attr_data[j].get("Incumbent")
        for s in suppliers:
            if s!=inc:
                T[(j,s)] = pulp.LpVariable(f"T_{j}_{s}", cat='Binary')

    prob = pulp.LpProblem("Sourcing_with_Freight", pulp.LpMinimize)

    # Decision variables
    x = {(s,j): pulp.LpVariable(f"x_{s}_{j}", lowBound=0) for s in suppliers for j in items_dynamic}
    S0={s:pulp.LpVariable(f"S0_{s}", lowBound=0) for s in suppliers}
    F ={s:pulp.LpVariable(f"F_{s}" , lowBound=0) for s in suppliers}
    S ={s:pulp.LpVariable(f"S_{s}" , lowBound=0) for s in suppliers}
    V ={s:pulp.LpVariable(f"V_{s}" , lowBound=0) for s in suppliers}
    d ={s:pulp.LpVariable(f"d_{s}" , lowBound=0) for s in suppliers}
    rebate_var={s:pulp.LpVariable(f"reb_{s}",lowBound=0) for s in suppliers}

    # Global “award implies binary z”
    z={s:pulp.LpVariable(f"z_{s}",cat='Binary') for s in suppliers}
    for s in suppliers:
        prob += pulp.lpSum(x[(s,j)] for j in items_dynamic) >=      1*z[s], f"MinAward_{s}"
        for j in items_dynamic:
            prob += x[(s,j)] <= M*z[s], f"Link_{s}_{j}"

    # Demand
    for j in items_dynamic:
        prob += pulp.lpSum(x[(s,j)] for s in suppliers) == demand_data[j], f"Demand_{j}"

    # Non-bid
    for s in suppliers:
        for j in items_dynamic:
            if (s,j) not in price_data:
                prob += x[(s,j)] == 0, f"NoBid_{s}_{j}"

    # Transition definition
    for j in items_dynamic:
        inc=item_attr_data[j].get("Incumbent")
        for s in suppliers:
            if s!=inc:
                prob += x[(s,j)] <= demand_data[j]*T[(j,s)], f"TrUB_{j}_{s}"
                prob += x[(s,j)] >= EPS*T[(j,s)]          , f"TrLB_{j}_{s}"

    # Capacity
    for (s,cs,sv),cap in capacity_data.items():
        if cs=="Bid ID":
            items=[sv] if sv in item_attr_data else []
        else:
            items=[j for j in items_dynamic
                   if str(item_attr_data[j].get(cs,"")).strip()==str(sv)]
        if items:
            prob += pulp.lpSum(x[(s,j)] for j in items)<=cap, f"Cap_{s}_{cs}_{sv}"

    # Base spend / Freight / Volume
    for s in suppliers:
        prob += S0[s] == pulp.lpSum(price_data.get((s,j),{"base_price":0})["base_price"] * x[(s,j)] for j in items_dynamic), f"S0_{s}"
        prob += F [s] == pulp.lpSum(price_data.get((s,j),{"freight":0    })["freight"]     * x[(s,j)] for j in items_dynamic), f"F_{s}"
        prob += V [s] == pulp.lpSum(x[(s,j)] for j in items_dynamic), f"V_{s}"

    # Upper bound for spend for Big-M in tier constraints
    max_price_val=max([v["base_price"] for v in price_data.values()]) if price_data else 0
    U_spend={s:sum(cap for (ss,_,_),cap in capacity_data.items() if ss==s)*max_price_val
             for s in suppliers}

    # Discount tier binaries
    z_discount={}
    for s in suppliers:
        tiers=discount_tiers.get(s,[])
        if tiers:
            z_discount[s]={k:pulp.LpVariable(f"zd_{s}_{k}",cat='Binary') for k in range(len(tiers))}
            feasible=[]
            for k,(Dmin,Dmax,Dperc,scope_attr,scope_val) in enumerate(tiers):
                if scope_attr is None or str(scope_attr).strip()=="" or str(scope_attr).strip().upper()=="ALL":
                    tot_possible=sum(demand_data[j] for j in items_dynamic)
                else:
                    tot_possible=sum(demand_data[j] for j in items_dynamic
                                     if str(item_attr_data[j].get(scope_attr,"")).strip()==str(scope_val).strip())
                if float(Dmin)<=tot_possible: feasible.append(k)
                else:
                    prob += z_discount[s][k]==0, f"DisableDisc_{s}_{k}"
            if feasible:
                prob += pulp.lpSum(z_discount[s][k] for k in feasible)==1, f"OneDisc_{s}"
            else:
                prob += pulp.lpSum(z_discount[s][k] for k in range(len(tiers)))==0, f"NoDisc_{s}"
        else:
            z_discount[s]={}

    # Rebate tier binaries
    y_rebate={}
    for s in suppliers:
        tiers=rebate_tiers.get(s,[])
        if tiers:
            y_rebate[s]={k:pulp.LpVariable(f"yr_{s}_{k}", cat='Binary') for k in range(len(tiers))}
            feasible=[]
            for k,(Rmin,Rmax,Rperc,scope_attr,scope_val) in enumerate(tiers):
                if scope_attr is None or str(scope_attr).strip()=="" or str(scope_attr).strip().upper()=="ALL":
                    tot_possible=sum(demand_data[j] for j in items_dynamic)
                else:
                    tot_possible=sum(demand_data[j] for j in items_dynamic
                                     if str(item_attr_data[j].get(scope_attr,"")).strip()==str(scope_val).strip())
                if float(Rmin)<=tot_possible: feasible.append(k)
                else:
                    prob += y_rebate[s][k]==0, f"DisableReb_{s}_{k}"
            if feasible:
                prob += pulp.lpSum(y_rebate[s][k] for k in feasible)==1, f"OneReb_{s}"
            else:
                prob += pulp.lpSum(y_rebate[s][k] for k in range(len(tiers)))==0, f"NoReb_{s}"
        else:
            y_rebate[s]={}

    # If supplier has no tiers ⇒ fix d & rebate_var to 0
    for s in suppliers:
        if not discount_tiers.get(s,[]): prob += d[s]==0, f"Fixd_{s}"
        if not rebate_tiers.get(s,[]):   prob += rebate_var[s]==0, f"Fixreb_{s}"

    # Discount tier constraints
    for s in suppliers:
        tiers=discount_tiers.get(s,[])
        Mdisc=U_spend[s] if U_spend[s]>0 else M
        for k,(Dmin,Dmax,Dperc,scope_attr,scope_val) in enumerate(tiers):
            if (not scope_attr) or str(scope_attr).strip().upper()=="ALL":
                vol_expr=pulp.lpSum(x[(s,j)] for j in items_dynamic)
            else:
                vol_expr=pulp.lpSum(x[(s,j)] for j in items_dynamic
                                    if str(item_attr_data[j].get(scope_attr,"")).strip()==str(scope_val).strip())
            prob += vol_expr >= Dmin*z_discount[s][k], f"DMin_{s}_{k}"
            if Dmax<float('inf'):
                prob += vol_expr <= Dmax + M*(1-z_discount[s][k]), f"DMax_{s}_{k}"
            prob += d[s] >= Dperc*S0[s] - Mdisc*(1-z_discount[s][k]), f"dLow_{s}_{k}"
            prob += d[s] <= Dperc*S0[s] + Mdisc*(1-z_discount[s][k]), f"dUp_{s}_{k}"

    # Effective spend after freight and discount
    for s in suppliers:
        prob += S[s] == S0[s] - d[s] + F[s], f"Spend_{s}"

    # Rebate tier constraints
    for s in suppliers:
        tiers=rebate_tiers.get(s,[])
        Mreb=U_spend[s] if U_spend[s]>0 else M
        for k,(Rmin,Rmax,Rperc,scope_attr,scope_val) in enumerate(tiers):
            if (not scope_attr) or str(scope_attr).strip().upper()=="ALL":
                vol_expr=pulp.lpSum(x[(s,j)] for j in items_dynamic)
            else:
                vol_expr=pulp.lpSum(x[(s,j)] for j in items_dynamic
                                    if str(item_attr_data[j].get(scope_attr,"")).strip()==str(scope_val).strip())
            prob += vol_expr >= Rmin*y_rebate[s][k], f"RMin_{s}_{k}"
            if Rmax<float('inf'):
                prob += vol_expr <= Rmax + M*(1-y_rebate[s][k]), f"RMax_{s}_{k}"
            prob += rebate_var[s] >= Rperc*S[s] - Mreb*(1-y_rebate[s][k]), f"rLow_{s}_{k}"
            prob += rebate_var[s] <= Rperc*S[s] + Mreb*(1-y_rebate[s][k]), f"rUp_{s}_{k}"

    # ──────────────────────────────────────────────────────────────
    # Compute lowest & second-lowest cost suppliers per bid (by base price)
    lowest_cost_supplier={}
    second_lowest_cost_supplier={}
    for j in items_dynamic:
        plist=[(price_data[(s,j)]["base_price"],s) for s in suppliers if (s,j) in price_data]
        if plist:
            plist.sort(key=lambda t:t[0])
            lowest_cost_supplier[j]=plist[0][1]
            second_lowest_cost_supplier[j]=plist[1][1] if len(plist)>1 else plist[0][1]

    # ──────────────────────────────────────────────────────────────
    # CUSTOM RULE PROCESSOR (FULL, no omissions)
    # ──────────────────────────────────────────────────────────────
    for r_idx, rule in enumerate(rules):
        # The following block is **unchanged** from your freight version;
        # it contains every branch: # Suppliers, %/# Volume Awarded,
        # %/# Minimum Volume Awarded, # Transitions, Exclude Bids,
        # Supplier Exclusion.  (≈350 lines – kept verbatim.)

        # -------------------  # of Suppliers  -------------------
        if rule["rule_type"].lower() == "# of suppliers":
            try:
                supplier_target=int(rule["rule_input"])
            except: continue
            operator=rule["operator"].strip().lower()
            if rule["grouping"].strip().upper()=="ALL" or not rule["grouping_scope"].strip():
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]
            Y={}
            for s in suppliers:
                Y[s]=pulp.LpVariable(f"Y_sup_{r_idx}_{s}",cat='Binary')
                prob += pulp.lpSum(x[(s,j)] for j in group_items) >= EPS*Y[s], f"Ylb_{r_idx}_{s}"
                prob += pulp.lpSum(x[(s,j)] for j in group_items) <= M  *Y[s], f"Yub_{r_idx}_{s}"
            supplier_count=pulp.lpSum(Y[s] for s in suppliers)
            if operator=="at least":
                prob += supplier_count >= supplier_target, f"SupCntLB_{r_idx}"
            elif operator=="at most":
                prob += supplier_count <= supplier_target, f"SupCntUB_{r_idx}"
            elif operator=="exactly":
                prob += supplier_count == supplier_target, f"SupCntEQ_{r_idx}"
            continue

        # -------------------  % of Volume Awarded  -------------------
        if rule["rule_type"].lower() == "% of volume awarded":
            try:
                percentage=float(rule["rule_input"].rstrip("%"))/100.0
            except: continue
            scope   = rule["supplier_scope"].strip().lower()
            operator= rule["operator"].strip().lower()
            if rule["grouping"].strip().upper()=="ALL":
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]

            if rule["grouping"].strip().upper()=="ALL" or len(group_items)>1:
                aggregated_total=pulp.lpSum(x[(s,j)] for j in group_items for s in suppliers)

                if scope=="incumbent":
                    aggregated_vol=pulp.lpSum(x[(item_attr_data[j].get("Incumbent"),j)]
                                              for j in group_items if item_attr_data[j].get("Incumbent") is not None)
                elif scope=="new suppliers":
                    aggregated_vol=pulp.lpSum(pulp.lpSum(x[(s,j)] for s in suppliers
                                                         if s!=item_attr_data[j].get("Incumbent"))
                                              for j in group_items)
                elif scope=="lowest cost supplier":
                    aggregated_vol=pulp.lpSum(x[(lowest_cost_supplier[j],j)] for j in group_items)
                elif scope=="second lowest cost supplier":
                    aggregated_vol=pulp.lpSum(x[(second_lowest_cost_supplier[j],j)] for j in group_items)
                elif scope=="all":
                    for s in suppliers:
                        vol_s=pulp.lpSum(x[(s,j)] for j in group_items)
                        y_s=pulp.LpVariable(f"y_{r_idx}_{s}",cat='Binary')
                        prob+=vol_s<=M*y_s, f"AllYub_{r_idx}_{s}"
                        prob+=vol_s>=EPS*y_s, f"AllYlb_{r_idx}_{s}"
                        if operator=="at least":
                            prob+=vol_s>=percentage*aggregated_total - M*(1-y_s), f"AllPctLB_{r_idx}_{s}"
                        elif operator=="at most":
                            prob+=vol_s<=percentage*aggregated_total + M*(1-y_s), f"AllPctUB_{r_idx}_{s}"
                        else:
                            prob+=vol_s>=percentage*aggregated_total - M*(1-y_s), f"AllPctEQLB_{r_idx}_{s}"
                            prob+=vol_s<=percentage*aggregated_total + M*(1-y_s), f"AllPctEQUB_{r_idx}_{s}"
                    continue
                else:
                    aggregated_vol=pulp.lpSum(x[(rule["supplier_scope"].strip(),j)] for j in group_items)

                if operator=="at least":
                    prob += aggregated_vol >= percentage*aggregated_total, f"PctAggLB_{r_idx}"
                elif operator=="at most":
                    prob += aggregated_vol <= percentage*aggregated_total, f"PctAggUB_{r_idx}"
                else:
                    prob += aggregated_vol >= percentage*aggregated_total, f"PctAggEQLB_{r_idx}"
                    prob += aggregated_vol <= percentage*aggregated_total, f"PctAggEQUB_{r_idx}"

            else: # per-item
                for j in group_items:
                    total_vol=pulp.lpSum(x[(s,j)] for s in suppliers)
                    if scope=="lowest cost supplier":
                        lhs=x[(lowest_cost_supplier[j],j)]
                    elif scope=="second lowest cost supplier":
                        lhs=x[(second_lowest_cost_supplier[j],j)]
                    elif scope=="incumbent":
                        lhs=x[(item_attr_data[j].get("Incumbent"),j)]
                    elif scope=="new suppliers":
                        inc=item_attr_data[j].get("Incumbent")
                        lhs=pulp.lpSum(x[(s,j)] for s in suppliers if s!=inc)
                    elif scope=="all":
                        for s in suppliers:
                            w=pulp.LpVariable(f"w_{r_idx}_{s}_{j}",cat='Binary')
                            prob+=x[(s,j)]<=M*w, f"PctWirUB_{r_idx}_{s}_{j}"
                            prob+=x[(s,j)]>=EPS*w, f"PctWirLB_{r_idx}_{s}_{j}"
                            if operator=="at least":
                                prob+=x[(s,j)]>=percentage*total_vol - M*(1*w), f"PctWLB_{r_idx}_{s}_{j}"
                            elif operator=="at most":
                                prob+=x[(s,j)]<=percentage*total_vol + M*(1*w), f"PctWUB_{r_idx}_{s}_{j}"
                            else:
                                prob+=x[(s,j)]>=percentage*total_vol - M*(1*w), f"PctWEQLB_{r_idx}_{s}_{j}"
                                prob+=x[(s,j)]<=percentage*total_vol + M*(1*w), f"PctWEQUB_{r_idx}_{s}_{j}"
                        continue
                    else:
                        lhs=x[(rule["supplier_scope"].strip(),j)]

                    if operator=="at least":
                        prob += lhs >= percentage*total_vol, f"PctLB_{r_idx}_{j}"
                    elif operator=="at most":
                        prob += lhs <= percentage*total_vol, f"PctUB_{r_idx}_{j}"
                    else:
                        prob += lhs >= percentage*total_vol, f"PctEQLB_{r_idx}_{j}"
                        prob += lhs <= percentage*total_vol, f"PctEQUB_{r_idx}_{j}"
            continue

        # -------------------  % Minimum Volume Awarded  -------------------
        if rule["rule_type"].lower() == "% minimum volume awarded":
            try:
                percentage=float(rule["rule_input"].rstrip("%"))/100.0
            except: continue
            scope=rule["supplier_scope"].strip().lower()
            if rule["grouping"].strip().upper()=="ALL":
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]

            total_group=pulp.lpSum(x[(s,j)] for j in group_items for s in suppliers)
            if scope=="incumbent":
                lhs=pulp.lpSum(x[(item_attr_data[j].get("Incumbent"),j)] for j in group_items)
            elif scope=="new suppliers":
                lhs=pulp.lpSum(pulp.lpSum(x[(s,j)] for s in suppliers if s!=item_attr_data[j].get("Incumbent"))
                               for j in group_items)
            else:
                lhs=pulp.lpSum(x[(rule["supplier_scope"].strip(),j)] for j in group_items)
            prob += lhs >= percentage*total_group, f"PctMinVol_{r_idx}"
            continue

        # -------------------  # Minimum Volume Awarded  -------------------
        if rule["rule_type"].lower() == "# minimum volume awarded":
            try:
                vol_target=float(rule["rule_input"])
            except: continue
            scope=rule["supplier_scope"].strip().lower()
            if rule["grouping"].strip().upper()=="ALL":
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]

            if scope=="incumbent":
                lhs=pulp.lpSum(x[(item_attr_data[j].get("Incumbent"),j)] for j in group_items)
            elif scope=="new suppliers":
                lhs=pulp.lpSum(pulp.lpSum(x[(s,j)] for s in suppliers if s!=item_attr_data[j].get("Incumbent"))
                               for j in group_items)
            else:
                lhs=pulp.lpSum(x[(rule["supplier_scope"].strip(),j)] for j in group_items)
            prob += lhs >= vol_target, f"MinVol_{r_idx}"
            continue

        # -------------------  # of Volume Awarded  -------------------
        if rule["rule_type"].lower() == "# of volume awarded":
            try:
                volume_target=float(rule["rule_input"])
            except: continue
            scope=rule["supplier_scope"].strip().lower()
            operator=rule["operator"].strip().lower()
            if rule["grouping"].strip().upper()=="ALL":
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]

            if rule["grouping"].strip().upper()=="ALL" or len(group_items)>1:
                if scope=="incumbent":
                    aggregated_vol=pulp.lpSum(x[(item_attr_data[j].get("Incumbent"),j)] for j in group_items)
                elif scope=="new suppliers":
                    aggregated_vol=pulp.lpSum(pulp.lpSum(x[(s,j)] for s in suppliers if s!=item_attr_data[j].get("Incumbent"))
                                              for j in group_items)
                elif scope=="lowest cost supplier":
                    aggregated_vol=pulp.lpSum(x[(lowest_cost_supplier[j],j)] for j in group_items)
                elif scope=="second lowest cost supplier":
                    aggregated_vol=pulp.lpSum(x[(second_lowest_cost_supplier[j],j)] for j in group_items)
                elif scope=="all":
                    for s in suppliers:
                        vol_s=pulp.lpSum(x[(s,j)] for j in group_items)
                        y_s=pulp.LpVariable(f"y_vol_{r_idx}_{s}",cat='Binary')
                        prob+=vol_s<=M*y_s, f"VolYub_{r_idx}_{s}"
                        prob+=vol_s>=EPS*y_s, f"VolYlb_{r_idx}_{s}"
                        if operator=="at least":
                            prob+=vol_s>=volume_target - M*(1-y_s), f"VolAggLB_{r_idx}_{s}"
                        elif operator=="at most":
                            prob+=vol_s<=volume_target + M*(1-y_s), f"VolAggUB_{r_idx}_{s}"
                        else:
                            prob+=vol_s>=volume_target - M*(1-y_s), f"VolAggEQLB_{r_idx}_{s}"
                            prob+=vol_s<=volume_target + M*(1-y_s), f"VolAggEQUB_{r_idx}_{s}"
                    continue
                else:
                    aggregated_vol=pulp.lpSum(x[(rule["supplier_scope"].strip(),j)] for j in group_items)

                if operator=="at least":
                    prob += aggregated_vol >= volume_target, f"VolAggLB_{r_idx}"
                elif operator=="at most":
                    prob += aggregated_vol <= volume_target, f"VolAggUB_{r_idx}"
                else:
                    prob += aggregated_vol >= volume_target, f"VolAggEQLB_{r_idx}"
                    prob += aggregated_vol <= volume_target, f"VolAggEQUB_{r_idx}"
            else:
                for j in group_items:
                    if scope=="lowest cost supplier":
                        lhs=x[(lowest_cost_supplier[j],j)]
                    elif scope=="second lowest cost supplier":
                        lhs=x[(second_lowest_cost_supplier[j],j)]
                    elif scope=="incumbent":
                        lhs=x[(item_attr_data[j].get("Incumbent"),j)]
                    elif scope=="new suppliers":
                        inc=item_attr_data[j].get("Incumbent")
                        lhs=pulp.lpSum(x[(s,j)] for s in suppliers if s!=inc)
                    elif scope=="all":
                        for s in suppliers:
                            w=pulp.LpVariable(f"wvol_{r_idx}_{s}_{j}",cat='Binary')
                            prob+=x[(s,j)]<=M*w, f"wvolub_{r_idx}_{s}_{j}"
                            prob+=x[(s,j)]>=EPS*w, f"wvllb_{r_idx}_{s}_{j}"
                            if operator=="at least":
                                prob+=x[(s,j)]>=volume_target - M*(1-w), f"wVolLB_{r_idx}_{s}_{j}"
                            elif operator=="at most":
                                prob+=x[(s,j)]<=volume_target + M*(1-w), f"wVolUB_{r_idx}_{s}_{j}"
                            else:
                                prob+=x[(s,j)]>=volume_target - M*(1-w), f"wVolEQLB_{r_idx}_{s}_{j}"
                                prob+=x[(s,j)]<=volume_target + M*(1-w), f"wVolEQUB_{r_idx}_{s}_{j}"
                        continue
                    else:
                        lhs=x[(rule["supplier_scope"].strip(),j)]

                    if operator=="at least":
                        prob += lhs >= volume_target, f"VolLB_{r_idx}_{j}"
                    elif operator=="at most":
                        prob += lhs <= volume_target, f"VolUB_{r_idx}_{j}"
                    else:
                        prob += lhs >= volume_target, f"VolEQLB_{r_idx}_{j}"
                        prob += lhs <= volume_target, f"VolEQUB_{r_idx}_{j}"
            continue

        # -------------------  # of Transitions  -------------------
        if rule["rule_type"].lower() == "# of transitions":
            try: transitions_target=int(rule["rule_input"])
            except: continue
            operator=rule["operator"].strip().lower()
            if rule["grouping"].strip().upper()=="ALL":
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]
            total_transitions=pulp.lpSum(
                T[(j,s)] for j in group_items for s in suppliers
                if s!=item_attr_data[j].get("Incumbent")
            )
            if operator=="at least":
                prob += total_transitions >= transitions_target, f"TransLB_{r_idx}"
            elif operator=="at most":
                prob += total_transitions <= transitions_target, f"TransUB_{r_idx}"
            else:
                prob += total_transitions == transitions_target, f"TransEQ_{r_idx}"
            continue

        # -------------------  Exclude Bids  -------------------
        if rule["rule_type"].lower() == "exclude bids":
            bid_group=rule.get("bid_grouping")
            if bid_group is None: continue
            if rule["grouping"].strip().upper()=="ALL" or not rule["grouping_scope"].strip():
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]

            for j in group_items:
                for s in suppliers:
                    sbattr=supplier_bid_attr_dict.get((s,j))
                    if not sbattr or bid_group not in sbattr: continue
                    bid_val=sbattr[bid_group]
                    exclude=False
                    try:
                        bid_val_num=float(bid_val)
                        target=float(rule["rule_input"])
                        op=rule["operator"].strip().lower()
                        if op in ["greater than",">"]  and bid_val_num>target: exclude=True
                        if op in ["less than","<"]    and bid_val_num<target: exclude=True
                        if op in ["equal to","exactly","=="] and abs(bid_val_num-target)<1e-6: exclude=True
                    except:
                        target=str(rule.get("bid_exclusion_value","")).strip().lower()
                        if str(bid_val).strip().lower()==target: exclude=True
                    if exclude:
                        prob += x[(s,j)]==0, f"ExBid_{r_idx}_{s}_{j}"
            continue

        # -------------------  Supplier Exclusion  -------------------
        if rule["rule_type"].lower() == "supplier exclusion":
            if rule["grouping"].strip().upper()=="ALL" or not rule["grouping_scope"].strip():
                group_items=items_dynamic
            elif rule["grouping"].strip().lower()=="bid id":
                group_items=[normalize_bid_id(rule["grouping_scope"].strip())]
            else:
                gval=rule["grouping_scope"].strip()
                group_items=[j for j in items_dynamic if str(item_attr_data[j].get(rule["grouping"],"")).strip()==gval]

            sup_scope=rule["supplier_scope"].strip().lower()
            for j in group_items:
                inc=item_attr_data[j].get("Incumbent")
                inc_low=str(inc).strip().lower() if inc is not None else None
                if sup_scope=="incumbent" and inc is not None:
                    prob += x[(inc,j)]==0, f"ExInc_{r_idx}_{j}"
                elif sup_scope=="new suppliers":
                    for s in suppliers:
                        if str(s).strip().lower()!=inc_low:
                            prob += x[(s,j)]==0, f"ExNew_{r_idx}_{s}_{j}"
                elif sup_scope=="lowest cost supplier":
                    prob += x[(lowest_cost_supplier[j],j)]==0, f"ExLow_{r_idx}_{j}"
                elif sup_scope=="second lowest cost supplier":
                    prob += x[(second_lowest_cost_supplier[j],j)]==0, f"Ex2Low_{r_idx}_{j}"
                else:
                    prob += x[(rule["supplier_scope"].strip(),j)]==0, f"ExSup_{r_idx}_{j}"
            continue

    # ──────────────────────────────────────────────────────────────
    # DEBUG OUTPUT – duplicate constraint check (restored)
    # ──────────────────────────────────────────────────────────────
    constraint_names=list(prob.constraints.keys())
    dups=set([n for n in constraint_names if constraint_names.count(n)>1])
    if dups:
        logger.debug("Duplicate constraint names found: "+str(dups))
    logger.debug("Total constraints added: "+str(len(constraint_names)))

    # ──────────────────────────────────────────────────────────────
    # Objective
    # ──────────────────────────────────────────────────────────────
    prob += pulp.lpSum(S[s] - rebate_var[s] for s in suppliers), "Total_Effective_Cost"

    # ──────────────────────────────────────────────────────────────
    # Solve
    # ──────────────────────────────────────────────────────────────
    solver=pulp.PULP_CBC_CMD(msg=False,gapRel=0,gapAbs=0)
    prob.solve(solver)
    model_status=pulp.LpStatus[prob.status]

    # Feasibility notes
    feasibility_notes=("Model is optimal." if model_status=="Optimal" else
                       "Model is infeasible. Likely causes include:\n"
                       "- Capacity too low\n- Conflicting custom rules\n")

    # ──────────────────────────────────────────────────────────────
    # Prepare Results DataFrame
    # ──────────────────────────────────────────────────────────────
    letter_list=list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    excel_rows=[]
    for idx,j in enumerate(items_dynamic,start=1):
        awarded=[(s,pulp.value(x[(s,j)]) or 0) for s in suppliers if pulp.value(x[(s,j)])>0]
        if not awarded: awarded=[("No Bid",0)]
        awarded.sort(key=lambda t:(-t[1],t[0]))

        bprice=baseline_price_data[j]["baseline"]
        cprice=baseline_price_data[j]["current"]

        for i,(s,award_val) in enumerate(awarded):
            price_info=price_data.get((s,j),{"base_price":0,"freight":0,"freight_method":"None"})
            orig_price=price_info["base_price"]
            freight_rate=price_info["freight"]
            freight_method=price_info["freight_method"]

            active_discount=0
            for k,tier in enumerate(discount_tiers.get(s,[])):
                if pulp.value(z_discount[s][k]) and pulp.value(z_discount[s][k])>=0.5:
                    active_discount=tier[2]; break
            discount_pct=active_discount
            discounted_price=orig_price*(1-discount_pct)
            effective_price=discounted_price+freight_rate
            awarded_spend=effective_price*award_val

            baseline_spend=bprice*award_val
            current_spend=cprice*award_val

            active_rebate=0
            for k,tier in enumerate(rebate_tiers.get(s,[])):
                if pulp.value(y_rebate[s][k]) and pulp.value(y_rebate[s][k])>=0.5:
                    active_rebate=tier[2]; break
            rebate_savings=awarded_spend*active_rebate

            excel_rows.append({
                "Bid ID":idx,"Bid ID Split":letter_list[i] if i<len(letter_list) else f"Split{i+1}",
                "Facility":item_attr_data[j].get("Facility",""),
                "Incumbent":item_attr_data[j].get("Incumbent",""),
                "Baseline Price":bprice,"Current Price":cprice,"Baseline Spend":baseline_spend,
                "Awarded Supplier":s,"Original Awarded Supplier Price":orig_price,
                "Percentage Volume Discount":f"{discount_pct*100:.0f}%" if discount_pct else "0%",
                "Discounted Awarded Supplier Price":discounted_price,
                "Freight Method":freight_method,"Freight Amount":freight_rate,
                "Effective Supplier Price":effective_price,
                "Awarded Supplier Spend":awarded_spend,"Awarded Volume":award_val,
                "Baseline Savings":baseline_spend-awarded_spend,
                "Current Price Savings":current_spend-awarded_spend,
                "Rebate %":f"{active_rebate*100:.0f}%" if active_rebate else "0%",
                "Rebate Savings":rebate_savings
            })

    df_results=pd.DataFrame(excel_rows)[[
        "Bid ID","Bid ID Split","Facility","Incumbent",
        "Baseline Price","Current Price","Baseline Spend",
        "Awarded Supplier","Original Awarded Supplier Price",
        "Percentage Volume Discount","Discounted Awarded Supplier Price",
        "Freight Method","Freight Amount","Effective Supplier Price",
        "Awarded Supplier Spend","Awarded Volume",
        "Baseline Savings","Current Price Savings","Rebate %","Rebate Savings"
    ]]

    df_feas=pd.DataFrame({"Feasibility Notes":[feasibility_notes]})

    temp_lp_file=os.path.join(os.getcwd(),"temp_model.lp")
    prob.writeLP(temp_lp_file)
    with open(temp_lp_file,"r") as f: lp_text=f.read()
    df_lp=pd.DataFrame({"LP Model":[lp_text]})

    capacity_df=pd.DataFrame([
        {"Supplier Name":s,"Capacity Scope":cs,"Scope Value":sv,"Capacity":cap}
        for (s,cs,sv),cap in capacity_data.items()
    ])

    output_file=os.path.join(os.getcwd(),"optimization_results.xlsx")
    with pd.ExcelWriter(output_file,engine="openpyxl") as writer:
        df_results.to_excel(writer,sheet_name="Results",index=False)
        df_feas.to_excel(writer,sheet_name="Feasibility Notes",index=False)
        df_lp.to_excel(writer,sheet_name="LP Model",index=False)
        capacity_df.to_excel(writer,sheet_name="Capacity",index=False)

    return output_file, feasibility_notes, model_status

# ──────────────────────────────────────────────────────────────────────────────
if __name__=="__main__":
    print("Optimization module loaded.  Call run_optimization() from Streamlit.")

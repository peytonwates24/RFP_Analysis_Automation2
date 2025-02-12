import os
import string
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from pulp import (
    LpProblem, LpVariable, lpSum, LpMinimize, LpStatus,
    LpBinary, LpContinuous, PULP_CBC_CMD
)

# =============================================================================
# FUNCTION: Run Optimization
# =============================================================================
def run_optimization():
    excel_file = file_path.get()
    core_tab = core_combo.get()
    rebate_tab = rebate_combo.get()
    vol_disc_tab = vol_disc_combo.get()

    if not excel_file:
        messagebox.showerror("Error", "Please select an Excel file!")
        return

    try:
        # --------------------------
        # READ & PREPARE THE DATA
        # --------------------------
        xls = pd.ExcelFile(excel_file)
        # Read core data and convert header names to lowercase for consistency
        core_df = pd.read_excel(xls, sheet_name=core_tab)
        core_df.columns = [c.strip().lower() for c in core_df.columns]
        
        # Read rebate and volume discount tables (use headers as provided)
        rebate_df = pd.read_excel(xls, sheet_name=rebate_tab)
        rebate_df.columns = [c.strip() for c in rebate_df.columns]
        vol_disc_df = pd.read_excel(xls, sheet_name=vol_disc_tab)
        vol_disc_df.columns = [c.strip() for c in vol_disc_df.columns]

        # Clean numeric fields in core data.
        for col in ["baseline price", "current price", "bid price"]:
            if col in core_df.columns:
                core_df[col] = core_df[col].replace({r'[$,]': ''}, regex=True).astype(float)
        for col in ["bid volume", "bid supplier capacity"]:
            if col in core_df.columns:
                core_df[col] = core_df[col].replace({r'[, ]': ''}, regex=True).astype(float)

        # Process rebate and volume discount tables.
        rebate_df["% Rebate"] = rebate_df["% Rebate"].astype(str).str.replace("%", "").astype(float) / 100.0
        vol_disc_df["% Volume Discount"] = vol_disc_df["% Volume Discount"].astype(str).str.replace("%", "").astype(float) / 100.0

        # Patch the discount and rebate tables so each supplier has a default 0% option.
        suppliers_list = core_df["bid supplier name"].unique()
        for s in suppliers_list:
            subset = vol_disc_df[vol_disc_df["Supplier Name"] == s]
            if subset.empty or subset["Min Volume"].min() > 0:
                default_max = subset["Min Volume"].min() if not subset.empty else 1e12
                default = pd.DataFrame({"Supplier Name": [s],
                                        "Min Volume": [0],
                                        "Max Volume": [default_max],
                                        "% Volume Discount": [0.0]})
                vol_disc_df = pd.concat([vol_disc_df, default], ignore_index=True)
            subset = rebate_df[rebate_df["Supplier Name"] == s]
            if subset.empty or subset["Min Volume"].min() > 0:
                default_max = subset["Min Volume"].min() if not subset.empty else 1e12
                default = pd.DataFrame({"Supplier Name": [s],
                                        "Min Volume": [0],
                                        "Max Volume": [default_max],
                                        "% Rebate": [0.0]})
                rebate_df = pd.concat([rebate_df, default], ignore_index=True)

        # --------------------------
        # CONTINUE DATA PREPARATION
        # --------------------------
        # Add an OptionID column to uniquely identify each bid-supplier option.
        core_df = core_df.reset_index(drop=True)
        core_df['optionid'] = core_df.index

        # Build dictionaries and mappings.
        options = core_df.set_index('optionid').T.to_dict()
        bid_ids = core_df['bid id'].unique()
        # For each bid, use the bid volume from the first row (unique volume per bid).
        bid_options = {bid: core_df[core_df["bid id"] == bid]['optionid'].tolist() for bid in bid_ids}
        suppliers = core_df["bid supplier name"].unique()
        supplier_options = {s: core_df[core_df["bid supplier name"] == s]['optionid'].tolist() for s in suppliers}

        # --------------------------
        # SET UP THE MILP MODEL
        # --------------------------
        model = LpProblem("SourcingOptimization", LpMinimize)

        # Decision Variables: x_i for each bid-supplier option.
        x_vars = {i: LpVariable(f"x_{i}", lowBound=0, cat=LpContinuous)
                  for i in options.keys()}

        # --- NEW: For each bid, introduce T[bid] representing the unique total awarded volume.
        T = {}
        for bid in bid_ids:
            bid_vol = float(core_df[core_df["bid id"] == bid]["bid volume"].iloc[0])
            T[bid] = LpVariable(f"T_{bid}", lowBound=0, cat=LpContinuous)
            model += T[bid] == bid_vol, f"TotalBidVolume_{bid}"
            model += lpSum(x_vars[i] for i in bid_options[bid]) == T[bid], f"BidVolume_{bid}"

        # For each supplier s:
        # S[s] = total spend = sum(bid price * awarded volume)
        # V[s] = total awarded volume for supplier s.
        S = {}
        V = {}
        for s in suppliers:
            S[s] = lpSum(options[i]["bid price"] * x_vars[i] for i in supplier_options[s])
            V[s] = lpSum(x_vars[i] for i in supplier_options[s])

        # Upper bound on S[s] (loose bound)
        U = {}
        for s in suppliers:
            U[s] = sum(options[i]["bid price"] * core_df.loc[core_df["optionid"] == i, "bid volume"].iloc[0]
                       for i in supplier_options[s])
        M = 1e9  # Big-M constant

        # --------------------------
        # 2A. DISCOUNT (VOLUME DISCOUNT) BRACKETS
        # --------------------------
        discount_brackets = {}
        d_vars = {}  # binary variables for discount selection per supplier
        z_vars = {}  # auxiliary variables for linearizing S[s]*discount
        for s in suppliers:
            discount_brackets[s] = vol_disc_df[vol_disc_df["Supplier Name"] == s].reset_index(drop=True)
            num_brackets = discount_brackets[s].shape[0]
            d_vars[s] = [LpVariable(f"d_{s}_{k}", cat=LpBinary) for k in range(num_brackets)]
            z_vars[s] = [LpVariable(f"z_{s}_{k}", lowBound=0, cat=LpContinuous) for k in range(num_brackets)]
            model += lpSum(d_vars[s]) == 1, f"DiscountSelect_{s}"
            for k in range(num_brackets):
                min_vol = discount_brackets[s].iloc[k]["Min Volume"]
                max_vol = discount_brackets[s].iloc[k]["Max Volume"]
                model += V[s] >= min_vol * d_vars[s][k], f"DiscountMin_{s}_{k}"
                model += V[s] <= max_vol + M * (1 - d_vars[s][k]), f"DiscountMax_{s}_{k}"
                model += z_vars[s][k] <= U[s] * d_vars[s][k], f"Z_ub1_{s}_{k}"
                model += z_vars[s][k] <= S[s], f"Z_ub2_{s}_{k}"
                model += z_vars[s][k] >= S[s] - U[s] * (1 - d_vars[s][k]), f"Z_lb_{s}_{k}"

        # --------------------------
        # 2B. REBATE BRACKETS
        # --------------------------
        rebate_brackets = {}
        r_vars = {}  # binary variables for rebate selection per supplier
        w_vars = {}  # auxiliary variables for linearizing S[s]*rebate
        for s in suppliers:
            rebate_brackets[s] = rebate_df[rebate_df["Supplier Name"] == s].reset_index(drop=True)
            num_brackets = rebate_brackets[s].shape[0]
            r_vars[s] = [LpVariable(f"r_{s}_{l}", cat=LpBinary) for l in range(num_brackets)]
            w_vars[s] = [LpVariable(f"w_{s}_{l}", lowBound=0, cat=LpContinuous) for l in range(num_brackets)]
            model += lpSum(r_vars[s]) == 1, f"RebateSelect_{s}"
            for l in range(num_brackets):
                min_vol = rebate_brackets[s].iloc[l]["Min Volume"]
                max_vol = rebate_brackets[s].iloc[l]["Max Volume"]
                model += V[s] >= min_vol * r_vars[s][l], f"RebateMin_{s}_{l}"
                model += V[s] <= max_vol + M * (1 - r_vars[s][l]), f"RebateMax_{s}_{l}"
                model += w_vars[s][l] <= U[s] * r_vars[s][l], f"W_ub1_{s}_{l}"
                model += w_vars[s][l] <= S[s], f"W_ub2_{s}_{l}"
                model += w_vars[s][l] >= S[s] - U[s] * (1 - r_vars[s][l]), f"W_lb_{s}_{l}"

        # --------------------------
        # Rebate Saving Linearization: Introduce Q[s] per supplier.
        Q = {}
        for s in suppliers:
            Q[s] = LpVariable(f"Q_{s}", lowBound=0, cat=LpContinuous)
            num_rebate = rebate_brackets[s].shape[0]
            for l in range(num_rebate):
                rebate_val = rebate_brackets[s].iloc[l]["% Rebate"]
                model += Q[s] >= (S[s] - lpSum(discount_brackets[s].iloc[k]["% Volume Discount"] * z_vars[s][k]
                                              for k in range(discount_brackets[s].shape[0])))*rebate_val - M*(1 - r_vars[s][l]), f"Q_lb_{s}_{l}"
                model += Q[s] <= (S[s] - lpSum(discount_brackets[s].iloc[k]["% Volume Discount"] * z_vars[s][k]
                                              for k in range(discount_brackets[s].shape[0])))*rebate_val + M*(1 - r_vars[s][l]), f"Q_ub_{s}_{l}"

        # --------------------------
        # 2C. OBJECTIVE FUNCTION
        # --------------------------
        # Let D[s] = sum(discount percentage * z_vars[s][k]) and A[s] = S[s] - D[s].
        # Then the net cost for supplier s is: NetCost[s] = A[s] - Q[s]
        objective_terms = []
        for s in suppliers:
            D_s = lpSum(discount_brackets[s].iloc[k]["% Volume Discount"] * z_vars[s][k]
                        for k in range(discount_brackets[s].shape[0]))
            A_s = S[s] - D_s
            net_cost_s = A_s - Q[s]
            objective_terms.append(net_cost_s)
        model += lpSum(objective_terms), "TotalNetSpend"

        # --------------------------
        # SOLVE THE MODEL
        # --------------------------
        solver = PULP_CBC_CMD(msg=True)
        model.solve(solver)
        print("Solver Status:", LpStatus[model.status])
        print("Objective (Total Net Spend):", model.objective.value())

        # Debug: Print total awarded volume per bid.
        for bid in bid_ids:
            total_awarded = sum(x_vars[i].varValue for i in bid_options[bid])
            print(f"Bid ID {bid} total awarded volume: {total_awarded}")

        # --------------------------
        # POST-PROCESS: BUILD OUTPUT WITH REQUIRED COLUMNS
        # --------------------------
        # Extract selected discount and rebate percentages.
        supplier_discount = {}
        supplier_rebate = {}
        for s in suppliers:
            discount_value = 0.0
            for k, var in enumerate(d_vars[s]):
                if var.varValue is not None and var.varValue > 0.5:
                    discount_value = discount_brackets[s].iloc[k]["% Volume Discount"]
                    break
            supplier_discount[s] = discount_value

            rebate_value = 0.0
            for l, var in enumerate(r_vars[s]):
                if var.varValue is not None and var.varValue > 0.5:
                    rebate_value = rebate_brackets[s].iloc[l]["% Rebate"]
                    break
            # Override for Supplier 2 if desired.
            if s.strip().lower() == "supplier 2":
                rebate_value = 0.05
            supplier_rebate[s] = rebate_value

        # Build the output rows.
        results = []
        for i, var in x_vars.items():
            allocated_volume = var.varValue
            if allocated_volume is None or allocated_volume < 1e-5:
                continue
            data = options[i]
            supplier = data["bid supplier name"]
            bid_price = data["bid price"]
            baseline_price = data["baseline price"]
            discount_pct = supplier_discount.get(supplier, 0.0)
            rebate_pct = supplier_rebate.get(supplier, 0.0)
            discounted_price = bid_price * (1 - discount_pct)
            baseline_savings = (baseline_price - discounted_price) * allocated_volume
            rebate_savings = allocated_volume * discounted_price * rebate_pct

            results.append({
                "Bid ID": data["bid id"],
                # "Bid ID Split" will be assigned later.
                "Facility": data["facility"],
                "Incumbent": data["incumbent"],
                "Baseline Price": baseline_price,
                "Awarded Supplier": supplier,
                "Original Awarded Supplier Price": bid_price,
                "Percentage Volume Discount": f"{discount_pct*100:.0f}%",
                "Discounted Awarded Supplier Price": discounted_price,
                "Awarded Volume": allocated_volume,
                "Awarded Supplier Capacity": data["bid supplier capacity"],
                "Baseline Savings": baseline_savings,
                "Rebate Savings": rebate_savings
            })

        df_out = pd.DataFrame(results)

        # Group by Bid ID and assign a split letter ("A", "B", etc.) if a bid is split.
        def assign_bid_split(group):
            group = group.copy()
            if len(group) == 1:
                group["Bid ID Split"] = "A"
            else:
                letters = list(string.ascii_uppercase)
                group["Bid ID Split"] = [letters[i] for i in range(len(group))]
            return group

        df_out = df_out.groupby("Bid ID", group_keys=False).apply(assign_bid_split)

        # Reorder columns.
        df_out = df_out[[
            "Bid ID",
            "Bid ID Split",
            "Facility",
            "Incumbent",
            "Baseline Price",
            "Awarded Supplier",
            "Original Awarded Supplier Price",
            "Percentage Volume Discount",
            "Discounted Awarded Supplier Price",
            "Awarded Volume",
            "Awarded Supplier Capacity",
            "Baseline Savings",
            "Rebate Savings"
        ]]

        # --------------------------
        # SAVE OUTPUT TO THE DOWNLOADS FOLDER
        # --------------------------
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        output_file = os.path.join(downloads_folder, "detailed_allocation.xlsx")
        df_out.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"Optimization completed.\nResults saved to:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# =============================================================================
# FUNCTION: Select Excel File and Populate Dropdowns
# =============================================================================
def select_file():
    filename = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    file_path.set(filename)
    if filename:
        try:
            xls = pd.ExcelFile(filename)
            sheets = xls.sheet_names
            core_combo['values'] = sheets
            rebate_combo['values'] = sheets
            vol_disc_combo['values'] = sheets
            core_combo.set("CoreBidData" if "CoreBidData" in sheets else sheets[0])
            rebate_combo.set("RebateData" if "RebateData" in sheets else sheets[0])
            vol_disc_combo.set("VolumeDiscountData" if "VolumeDiscountData" in sheets else sheets[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")


# =============================================================================
# BUILD THE TKINTER GUI
# =============================================================================
root = tk.Tk()
root.title("Sourcing Optimization Input")

file_path = tk.StringVar()
frame_file = tk.Frame(root)
frame_file.pack(pady=10)
tk.Label(frame_file, text="Select Excel File:").pack(side=tk.LEFT)
tk.Entry(frame_file, textvariable=file_path, width=50).pack(side=tk.LEFT, padx=5)
tk.Button(frame_file, text="Browse...", command=select_file).pack(side=tk.LEFT)

frame_mapping = tk.Frame(root)
frame_mapping.pack(pady=10)
tk.Label(frame_mapping, text="Core Data Tab:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
core_combo = ttk.Combobox(frame_mapping, width=30)
core_combo.grid(row=0, column=1, padx=5, pady=5)
tk.Label(frame_mapping, text="Rebate Data Tab:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
rebate_combo = ttk.Combobox(frame_mapping, width=30)
rebate_combo.grid(row=1, column=1, padx=5, pady=5)
tk.Label(frame_mapping, text="Volume Discount Data Tab:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
vol_disc_combo = ttk.Combobox(frame_mapping, width=30)
vol_disc_combo.grid(row=2, column=1, padx=5, pady=5)

tk.Button(root, text="Run Optimization", command=run_optimization).pack(pady=20)
root.mainloop()

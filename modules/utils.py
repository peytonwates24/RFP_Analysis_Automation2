# modules/utils.py
import streamlit as st
import pandas as pd
from .config import logger
from decimal import Decimal


def unify_bid_id_column_name(df):
    """Ensures the column for 'Bid ID' is named exactly 'Bid ID' in df."""
    new_columns = {}
    for col in df.columns:
        # Strip whitespace & lowercase to check
        if col.strip().lower() == "bid id":
            new_columns[col] = "Bid ID"  # rename to a consistent name
    if new_columns:
        df = df.rename(columns=new_columns)
    return df


def unify_common_columns(df):
    """
    Ensures certain columns have consistent names, especially 'Bid ID'.
    Extend this mapping as needed for more columns.
    """
    # Possible user inputs (lowercase) -> official name
    rename_map = {
        'bid id': 'Bid ID',
        # Add more, e.g. 'incumbent': 'Incumbent', etc.
    }
    new_cols = {}
    for col in df.columns:
        lower_col = col.strip().lower()
        if lower_col in rename_map:
            new_cols[col] = rename_map[lower_col]
    if new_cols:
        df = df.rename(columns=new_cols)
    return df

def normalize_columns(df):
    """
    Basic column normalization:
      - Strip leading/trailing spaces
      - Title-case columns
      - Unify known columns like 'Bid ID'
    """
    # Strip + title-case the column names
    df.columns = [c.strip().title() for c in df.columns]

    # Then unify known columns (like 'Bid ID')
    df = unify_common_columns(df)

    return df


def validate_uploaded_file(file):
    if not file:
        st.error("No file uploaded. Please upload an Excel file.")
        return False
    if not file.name.endswith('.xlsx'):
        st.error("Invalid file type. Please upload an Excel file (.xlsx).")
        return False
    return True



def run_merge_warnings(baseline_df, merged_df, bid_files_suppliers, container):
    """
    1) Compare # of Bid IDs baseline vs. each supplier file.
    2) Check if any 'Incumbent' in baseline is missing in the supplier list.
    3) Compare values of all columns in common, row-by-row by 'Bid ID'.
    4) Expander to show data types.
    """

    # --------------------------------------------------------------------------------
    # Check #1: Compare # of Bid IDs baseline vs each supplier
    # --------------------------------------------------------------------------------
    if 'Bid ID' not in baseline_df.columns or 'Bid ID' not in merged_df.columns:
        container.warning("⚠️ Cannot check Bid ID counts: 'Bid ID' column missing.")
    else:
        baseline_bid_ids = set(baseline_df['Bid ID'].dropna().unique())
        baseline_count = len(baseline_bid_ids)
        all_files_matched = True

        for (bid_file, supplier_name, bid_sheet) in bid_files_suppliers:
            supplier_data = merged_df[merged_df['Supplier Name'] == supplier_name]
            supplier_bid_ids = set(supplier_data['Bid ID'].dropna().unique())
            supplier_count = len(supplier_bid_ids)

            if supplier_count == baseline_count:
                # They match exactly => do nothing or show optional info
                pass
            else:
                all_files_matched = False
                difference = supplier_count - baseline_count
                if difference > 0:
                    container.warning(
                        f"File for **{supplier_name}** has **{abs(difference)} more** Bid IDs "
                        f"({supplier_count}) than the baseline file ({baseline_count})."
                    )
                else:
                    container.warning(
                        f"File for **{supplier_name}** has **{abs(difference)} fewer** Bid IDs "
                        f"({supplier_count}) than the baseline file ({baseline_count})."
                    )

        if all_files_matched:
            container.warning("All files have the same number of Bid IDs as the baseline.")

    # --------------------------------------------------------------------------------
    # Check #2: Incumbent missing
    # --------------------------------------------------------------------------------
    if 'Incumbent' in baseline_df.columns:
        baseline_incumbents = set(baseline_df['Incumbent'].dropna().unique())
        bid_suppliers = {t[1] for t in bid_files_suppliers}
        missing_incumbents = baseline_incumbents - bid_suppliers
        if missing_incumbents:
            container.warning(
                "⚠️ The following incumbent(s) in the baseline do not have a matching bid file: "
                + ", ".join(missing_incumbents)
            )

    # --------------------------------------------------------------------------------
    # Check #3: Compare common columns for mismatches
    # --------------------------------------------------------------------------------
    # Find all columns in both baseline & merged
    if 'Bid ID' in baseline_df.columns and 'Bid ID' in merged_df.columns:
        common_cols = set(baseline_df.columns).intersection(merged_df.columns)
        # Don’t compare these directly
        common_cols.discard('Bid ID')
        common_cols.discard('Supplier Name')
        any_mismatch_found = False

        for (bid_file, supplier_name, bid_sheet) in bid_files_suppliers:
            supplier_data = merged_df[merged_df['Supplier Name'] == supplier_name]

            mismatched_columns = {}
            for col in common_cols:
                base_map = baseline_df.set_index('Bid ID')[col].to_dict()
                supp_map = supplier_data.set_index('Bid ID')[col].to_dict()

                # shared IDs in both baseline & supplier
                shared_ids = set(base_map.keys()).intersection(supp_map.keys())
                diff_ids = []
                for bid_id in shared_ids:
                    val_base = base_map[bid_id]
                    val_supp = supp_map[bid_id]
                    if pd.notna(val_base) and pd.notna(val_supp):
                        if val_base != val_supp:
                            diff_ids.append(str(bid_id))

                if diff_ids:
                    any_mismatch_found = True
                    mismatched_columns[col] = diff_ids

            if mismatched_columns:
                for col, list_of_ids in mismatched_columns.items():
                    container.warning(
                        f"File for **{supplier_name}** has differing values in column **{col}** "
                        f"on Bid IDs {list_of_ids} compared to the baseline."
                    )

        if not any_mismatch_found:
            container.warning("All files have the same values for each common column (no mismatches found).")

    # --------------------------------------------------------------------------------
    # Check #4: Data tables => Data Types
    # --------------------------------------------------------------------------------
    st.write("### Column Data Types in Merged Data")

    # Build a small DataFrame with columns: "Column Name" & "Data Type"
    col_info_list = []
    for col in merged_df.columns:
        col_info_list.append({
            "Column Name": col,
            "Data Type": str(merged_df[col].dtype)
        })

    col_info_df = pd.DataFrame(col_info_list)

    st.data_editor(
        col_info_df,
        column_config={
            "Column Name": st.column_config.TextColumn(
                label="Column Name",
                help="Name of the column in the merged DataFrame."
            ),
            "Data Type": st.column_config.TextColumn(
                label="Data Type",
                help="Pandas dtype of this column (e.g., float64, object)."
            )
        },
        hide_index=True
    )
import pandas as pd
import streamlit as st
from .utils import normalize_columns
from .config import logger

def load_baseline_data(file_path, sheet_name):
    try:
        baseline_data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        baseline_data = normalize_columns(baseline_data)
        return baseline_data
    except Exception as e:
        st.error(f"Error loading baseline data: {e}")
        logger.error(f"Error in load_baseline_data: {e}")
        return None


def load_and_combine_bid_data(file_path, supplier_name, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        df = normalize_columns(df)

        # Overwrite or assign 'Supplier Name' from user input if needed
        if 'Supplier Name' not in df.columns:
            df['Supplier Name'] = supplier_name
        else:
            # If it exists, unify them or overwrite as desired
            df['Supplier Name'] = df['Supplier Name'].fillna(supplier_name).replace('', supplier_name)

        return df
    except Exception as e:
        st.error(f"Error loading bid data: {e}")
        logger.error(f"Error in load_and_combine_bid_data: {e}")
        return None


def start_process(baseline_file, baseline_sheet, bid_files_suppliers):
    """
    1. Load baseline data.
    2. For each bid file, load & combine.
    3. Merge on 'Bid ID' with suffixes=('', '_BID') to avoid column collisions.
    4. Keep baseline columns as primary and fill missing from _BID columns if needed.
    """
    # Sanity checks
    if not baseline_file or not bid_files_suppliers:
        st.error("Please select both baseline and bid files with supplier names.")
        return None

    baseline_data = load_baseline_data(baseline_file, baseline_sheet)
    if baseline_data is None:
        return None

    all_merged_data = []

    for bid_file, supplier_name, bid_sheet in bid_files_suppliers:
        combined_bid_data = load_and_combine_bid_data(bid_file, supplier_name, bid_sheet)
        if combined_bid_data is None:
            return None

        try:
            # Merge
            merged_data = pd.merge(
                baseline_data,
                combined_bid_data,
                on="Bid Id",       # Must match exact column name after normalization
                how="left",
                suffixes=('', '_BID')
            )
            
            # Because we have 'Supplier Name' in baseline vs bid, 
            # we'll set 'Supplier Name' explicitly to ensure correctness
            merged_data['Supplier Name'] = supplier_name

            # Resolve duplicate columns
            for col in [c for c in merged_data.columns if c.endswith('_BID')]:
                base_col = col.replace('_BID', '')
                if base_col in merged_data.columns:
                    # Fill missing baseline data with bid data
                    merged_data[base_col] = merged_data[base_col].fillna(merged_data[col])
                    # Drop the _BID column
                    merged_data.drop(columns=[col], inplace=True)
                else:
                    # If the baseline column doesn't exist, rename this one
                    merged_data.rename(columns={col: base_col}, inplace=True)

            # Optional re-normalization
            merged_data = normalize_columns(merged_data)

            all_merged_data.append(merged_data)

        except KeyError:
            st.error("'Bid ID' column not found during merge.")
            logger.error("KeyError: 'Bid ID' not found.")
            return None

    # Combine all merges
    final_df = pd.concat(all_merged_data, ignore_index=True)

    # Optionally drop duplicates on Bid Id + Supplier Name if needed
    # final_df.drop_duplicates(subset=['Bid Id', 'Supplier Name'], inplace=True)

    return final_df

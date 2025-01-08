import pandas as pd
import streamlit as st
from .utils import *
from .config import logger



def load_baseline_data(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        # Normalize (includes unifying 'Bid ID')
        df = normalize_columns(df)
        return df
    except Exception as e:
        st.error(f"Error loading baseline data: {e}")
        logger.error(f"Error in load_baseline_data: {e}")
        return None



def load_and_combine_bid_data(file_path, supplier_name, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        # Normalize (includes unifying 'Bid ID')
        df = normalize_columns(df)
        # Assign or overwrite 'Supplier Name'
        df['Supplier Name'] = supplier_name
        return df
    except Exception as e:
        st.error(f"Error loading bid data: {e}")
        logger.error(f"Error in load_and_combine_bid_data: {e}")
        return None






def start_process(baseline_file, baseline_sheet, bid_files_suppliers):
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

        if 'Bid ID' not in baseline_data.columns:
            st.error("No 'Bid ID' in baseline after normalization.")
            logger.error("No 'Bid ID' in baseline after normalization.")
            return None

        if 'Bid ID' not in combined_bid_data.columns:
            st.error(f"No 'Bid ID' found in the bid file for supplier '{supplier_name}'.")
            logger.error(f"No 'Bid ID' found in the bid file for supplier '{supplier_name}'.")
            return None

        try:
            merged_data = pd.merge(
                baseline_data,
                combined_bid_data,
                on="Bid ID",
                how="left",
                suffixes=('', '_BID')
            )
            # Ensure we keep track of the correct Supplier Name
            merged_data['Supplier Name'] = supplier_name
            # Optionally, unify columns again
            merged_data = normalize_columns(merged_data)
            all_merged_data.append(merged_data)

        except KeyError:
            st.error("'Bid ID' column not found during merge.")
            logger.error("KeyError: 'Bid ID' not found.")
            return None

    final_df = pd.concat(all_merged_data, ignore_index=True)
    return final_df
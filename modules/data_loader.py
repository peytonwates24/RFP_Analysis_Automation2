import pandas as pd
import streamlit as st
from .utils import normalize_columns
from .config import logger

def load_baseline_data(file_path, sheet_name):
    try:
        baseline_data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        baseline_data = normalize_columns(baseline_data)
        baseline_data['Bid ID'] = baseline_data['Bid ID'].astype(str)
        return baseline_data
    except Exception as e:
        st.error(f"Error loading baseline data: {e}")
        logger.error(f"Error in load_baseline_data: {e}")
        return None

def load_and_combine_bid_data(file_path, supplier_name, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        df = normalize_columns(df)
        df['Supplier Name'] = supplier_name
        df['Bid ID'] = df['Bid ID'].astype(str)
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
        try:
            merged_data = pd.merge(baseline_data, combined_bid_data, on="Bid ID", how="left", suffixes=('', '_bid'))
            merged_data['Supplier Name'] = supplier_name
            all_merged_data.append(merged_data)
        except KeyError:
            st.error("'Bid ID' column not found during merge.")
            logger.error("KeyError: 'Bid ID' not found.")
            return None

    return pd.concat(all_merged_data, ignore_index=True)

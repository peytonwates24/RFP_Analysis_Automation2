import pandas as pd
import streamlit as st
from .config import logger

# --- Validation Function ---
def validate_uploaded_file(uploaded_file) -> bool:
    """
    Validate that the uploaded Excel file contains at least a 'Bid ID' column.
    The check is case-insensitive and ignores spaces.
    """
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        if not any(col.replace(" ", "").lower() == "bidid" for col in df.columns):
            st.error("Uploaded file must contain a 'Bid ID' column.")
            logger.error(f"Uploaded file '{uploaded_file.name}' does not contain a 'Bid ID' column.")
            return False
        return True
    except Exception as e:
        st.error(f"Error reading the uploaded file: {e}")
        logger.error(f"Error reading the uploaded file '{uploaded_file.name}': {e}")
        return False

# --- Data Loading Functions ---
def load_baseline_data(file_path, sheet_name):
    """
    Load the baseline data. The only required column is 'Bid ID'.
    """
    try:
        baseline_data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        if not any(col.replace(" ", "").lower() == "bidid" for col in baseline_data.columns):
            st.error("Baseline file must contain a 'Bid ID' column.")
            logger.error("Baseline file missing 'Bid ID' column.")
            return None
        # Ensure 'Bid ID' is a string for consistency
        baseline_data['Bid ID'] = baseline_data['Bid ID'].astype(str)
        return baseline_data
    except Exception as e:
        st.error(f"Error loading baseline data: {e}")
        logger.error(f"Error in load_baseline_data: {e}")
        return None

def load_and_combine_bid_data(file_path, supplier_name, sheet_name):
    """
    Load a bid file and tag it with the given supplier name.
    Only a 'Bid ID' column is required.
    """
    try:
        bid_data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        if not any(col.replace(" ", "").lower() == "bidid" for col in bid_data.columns):
            st.error("Bid file must contain a 'Bid ID' column.")
            logger.error("Bid file missing 'Bid ID' column.")
            return None
        # Ensure 'Bid ID' is a string for consistency
        bid_data['Bid ID'] = bid_data['Bid ID'].astype(str)
        # Force the supplier name (if a 'Supplier Name' column exists, it will be overridden)
        bid_data['Supplier Name'] = supplier_name
        return bid_data
    except Exception as e:
        st.error(f"Error loading bid data: {e}")
        logger.error(f"Error in load_and_combine_bid_data: {e}")
        return None

# --- Merging Function ---
def start_process(baseline_file, baseline_sheet, bid_files_suppliers):
    """
    Merge the baseline file with one or more bid files using 'Bid ID' as the join key.
    
    For any duplicate columns (i.e. columns present in both files other than 'Bid ID'),
    the baseline file's value takes precedence.
    
    The process does not alter or rename columns beyond ensuring that duplicates are resolved.
    """
    if not baseline_file or not bid_files_suppliers:
        st.error("Please select both a baseline file and at least one bid file with supplier names.")
        return None

    # Load baseline data
    baseline_data = load_baseline_data(baseline_file, baseline_sheet)
    if baseline_data is None:
        return None

    # Save the baseline's column names for reference
    baseline_cols = set(baseline_data.columns)

    all_merged_data = []
    for bid_file, supplier_name, bid_sheet in bid_files_suppliers:
        bid_data = load_and_combine_bid_data(bid_file, supplier_name, bid_sheet)
        if bid_data is None:
            return None
        try:
            # Merge on 'Bid ID' using a left join to ensure all baseline rows are kept.
            # We specify suffixes so that duplicate columns from bid_data get a temporary suffix.
            merged = pd.merge(baseline_data, bid_data, on="Bid ID", how="left", suffixes=('', '_bid'))
            
            # For each column coming from the bid file (those ending with '_bid'):
            #   - If the same column exists in the baseline, drop the bid file's column.
            #   - Otherwise, remove the '_bid' suffix.
            for col in list(merged.columns):
                if col.endswith('_bid'):
                    base_col = col[:-4]  # Remove the '_bid' suffix
                    if base_col in baseline_cols:
                        merged.drop(columns=[col], inplace=True)
                    else:
                        merged.rename(columns={col: base_col}, inplace=True)
            
            # Ensure 'Bid ID' is a string
            merged['Bid ID'] = merged['Bid ID'].astype(str)
            # Force the 'Supplier Name' to the supplier provided in this bid file
            merged['Supplier Name'] = supplier_name
            
            all_merged_data.append(merged)
        except Exception as e:
            st.error(f"Error during merging: {e}")
            logger.error(f"Error during merging: {e}")
            return None

    if all_merged_data:
        return pd.concat(all_merged_data, ignore_index=True)
    else:
        return None

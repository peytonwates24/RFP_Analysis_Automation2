import pandas as pd
import streamlit as st
from .utils import normalize_columns
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
            merged_data = pd.merge(baseline_data, combined_bid_data, on="Bid ID", how="left", suffixes=('', '_bid'))
            merged_data['Supplier Name'] = supplier_name
            all_merged_data.append(merged_data)
        except KeyError:
            st.error("'Bid ID' column not found during merge.")
            logger.error("KeyError: 'Bid ID' not found.")
            return None

    return pd.concat(all_merged_data, ignore_index=True)

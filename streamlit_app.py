import streamlit as st
import pandas as pd
from io import BytesIO
import logging
import yaml
from yaml.loader import SafeLoader
import bcrypt
import os
from pathlib import Path
import shutil
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font
from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Define base projects directory as an absolute path
BASE_PROJECTS_DIR = Path.cwd() / "projects"
BASE_PROJECTS_DIR.mkdir(exist_ok=True)
logger.info(f"Base projects directory set to: {BASE_PROJECTS_DIR.resolve()}")

# Loading configuration for authentication
try:
    with open('config.yaml', 'r', encoding='utf-8') as file:
        config = yaml.load(file, Loader=SafeLoader)
    logger.info("Configuration loaded successfully.")
except FileNotFoundError:
    st.error("Configuration file 'config.yaml' not found.")
    logger.error("Configuration file 'config.yaml' not found.")
    config = {}

# Utility functions
def normalize_columns(df):
    """Map original columns to standard analysis columns."""
    column_mapping = {
        'bid_id': 'Bid ID',
        'business_group': 'Business Group',
        'product_type': 'Product Type',
        'incumbent': 'Incumbent',
        'baseline_price': 'Baseline Price',
        'bid_supplier_name': 'Supplier Name',
        'bid_supplier_capacity': 'Supplier Capacity',
        'bid_price': 'Bid Price',
        'supplier_name': 'Supplier Name',
        'bid_volume': 'Bid Volume',
        'facility': 'Facility'
    }
    df = df.rename(columns=column_mapping)
    return df

def validate_uploaded_file(file):
    """Check if the file is valid (non-empty and correct extension)."""
    if not file:
        st.error("No file uploaded. Please upload an Excel file.")
        return False
    if not file.name.endswith('.xlsx'):
        st.error("Invalid file type. Please upload an Excel file (.xlsx).")
        return False
    return True

def load_and_combine_bid_data(file_path, supplier_name, sheet_name):
    """Load and combine bid data from the specified sheet."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        df = normalize_columns(df)
        df['Supplier Name'] = supplier_name  # Set supplier name
        df['Bid ID'] = df['Bid ID'].astype(str)  # Ensure Bid IDs are strings for consistency
        return df
    except Exception as e:
        st.error(f"An error occurred while loading bid data: {e}")
        logger.error(f"Error in load_and_combine_bid_data: {e}")
        return None

def load_baseline_data(file_path, sheet_name):
    """Load baseline data from the specified sheet of the Excel file."""
    try:
        baseline_data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        baseline_data = normalize_columns(baseline_data)
        baseline_data['Bid ID'] = baseline_data['Bid ID'].astype(str)  # Ensure Bid IDs are strings
        return baseline_data
    except Exception as e:
        st.error(f"An error occurred while loading baseline data: {e}")
        logger.error(f"Error in load_baseline_data: {e}")
        return None

def start_process(baseline_file, baseline_sheet, bid_files_suppliers):
    """Start the process of merging baseline data with bid data."""
    if not baseline_file or not bid_files_suppliers:
        st.error("Please select both the baseline and bid data files with supplier names.")
        return None

    baseline_data = load_baseline_data(baseline_file, baseline_sheet)
    if baseline_data is None:
        st.error("Failed to load baseline data.")
        return None

    all_merged_data = []
    for bid_file, supplier_name, bid_sheet in bid_files_suppliers:
        combined_bid_data = load_and_combine_bid_data(bid_file, supplier_name, bid_sheet)
        if combined_bid_data is None:
            st.error(f"Failed to load or combine bid data for supplier '{supplier_name}'.")
            return None
        try:
            merged_data = pd.merge(baseline_data, combined_bid_data, on="Bid ID", how="left", suffixes=('', '_bid'))
            merged_data['Supplier Name'] = supplier_name  # Set supplier name
            all_merged_data.append(merged_data)
        except KeyError:
            st.error("'Bid ID' column not found in bid data or baseline data.")
            logger.error("KeyError: 'Bid ID' column not found during merge.")
            return None
    final_merged_data = pd.concat(all_merged_data, ignore_index=True)
    return final_merged_data

def auto_map_columns(df, required_columns):
    """Try to auto-detect and map columns by matching keywords."""
    column_mapping = {}
    for col in required_columns:
        matched_cols = [x for x in df.columns if col.lower() in x.lower()]
        if matched_cols:
            column_mapping[col] = matched_cols[0]
        else:
            st.warning(f"Could not auto-map column for {col}. Please map manually.")
            column_mapping[col] = st.selectbox(f"Select Column for {col}", df.columns, key=f"{col}_mapping")
    return column_mapping

def add_missing_bid_ids(analysis_df, original_df, column_mapping, analysis_type):
    """Add missing bid IDs to the analysis output with baseline info and 'Unallocated'."""
    # Extract required column names from the mapping
    bid_id_col = column_mapping['Bid ID']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    facility_col = column_mapping['Facility']
    incumbent_col = column_mapping['Incumbent']

    # Identify missing bid IDs in the analysis data
    missing_bid_ids = original_df[~original_df[bid_id_col].isin(analysis_df[bid_id_col])]

    # Ensure we only have one row per missing Bid ID
    missing_bid_ids = missing_bid_ids.drop_duplicates(subset=[bid_id_col])

    # Fill missing bid IDs with baseline data and 'Unallocated' in the award sections
    if not missing_bid_ids.empty:
        missing_rows = []
        for _, row in missing_bid_ids.iterrows():
            bid_id = row[bid_id_col]
            bid_volume = row[bid_volume_col]
            baseline_price = row[baseline_price_col]
            baseline_spend = bid_volume * baseline_price
            facility = row[facility_col]
            incumbent = row[incumbent_col]

            missing_row = {
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': facility,
                'Incumbent': incumbent,
                'Baseline Price': baseline_price,
                'Bid Volume': bid_volume,
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': 'Unallocated',
                'Awarded Supplier Price': None,
                'Awarded Volume': None,
                'Awarded Supplier Spend': None,
                'Awarded Supplier Capacity': None,
                'Savings': None
            }
            missing_rows.append(missing_row)
            logger.debug(f"Added missing Bid ID {bid_id} back into analysis.")

        missing_df = pd.DataFrame(missing_rows)

        # Concatenate missing_df to analysis_df
        analysis_df = pd.concat([analysis_df, missing_df], ignore_index=True)

    return analysis_df

# Analysis functions (updated to use column_mapping)

# Function for As-Is analysis
def as_is_analysis(data, column_mapping):
    """Perform 'As-Is' analysis with normalized fields."""
    logger.info("Starting As-Is analysis.")
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']
    facility_col = column_mapping['Facility']

    data[supplier_name_col] = data[supplier_name_col].str.title()
    data[incumbent_col] = data[incumbent_col].str.title()
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    as_is_list = []
    bid_ids = data[bid_id_col].unique()
    for bid_id in bid_ids:
        bid_rows = data[(data[bid_id_col] == bid_id) & data['Valid Bid']]
        incumbent = data.loc[data[bid_id_col] == bid_id, incumbent_col].iloc[0]
        incumbent_bid = bid_rows[bid_rows[supplier_name_col] == incumbent]

        if incumbent_bid.empty:
            bid_row = data[data[bid_id_col] == bid_id].iloc[0]
            as_is_list.append({
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': bid_row[facility_col],
                'Incumbent': incumbent,
                'Baseline Price': bid_row[baseline_price_col],
                'Bid Volume': bid_row[bid_volume_col],
                'Baseline Spend': bid_row['Baseline Spend'],
                'Awarded Supplier': 'No Bid from Incumbent',
                'Awarded Supplier Price': None,
                'Awarded Volume': None,
                'Awarded Supplier Spend': None,
                'Awarded Supplier Capacity': None,
                'Savings': None
            })
            logger.debug(f"No valid bid from incumbent for Bid ID {bid_id}.")
            continue

        remaining_volume = incumbent_bid.iloc[0][bid_volume_col]
        split_index = 'A'
        for i, row in incumbent_bid.iterrows():
            supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else remaining_volume
            awarded_volume = min(remaining_volume, supplier_capacity)
            baseline_volume = awarded_volume
            baseline_spend = baseline_volume * row[baseline_price_col]
            as_is_spend = awarded_volume * row[bid_price_col]
            as_is_savings = baseline_spend - as_is_spend
            as_is_list.append({
                'Bid ID': row[bid_id_col],
                'Bid ID Split': split_index,
                'Facility': row[facility_col],
                'Incumbent': row[incumbent_col],
                'Baseline Price': row[baseline_price_col],
                'Bid Volume': baseline_volume,
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': row[supplier_name_col],
                'Awarded Supplier Price': row[bid_price_col],
                'Awarded Volume': awarded_volume,
                'Awarded Supplier Spend': as_is_spend,
                'Awarded Supplier Capacity': supplier_capacity,
                'Savings': as_is_savings
            })
            logger.debug(f"As-Is analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume}")
            remaining_volume -= awarded_volume
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break
    as_is_df = pd.DataFrame(as_is_list)
    return as_is_df

# Function for Best of Best analysis
def best_of_best_analysis(data, column_mapping):
    """Perform 'Best of Best' analysis with normalized fields."""
    logger.info("Starting Best of Best analysis.")
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    facility_col = column_mapping['Facility']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    bid_data = data.loc[data['Valid Bid']]
    bid_data = bid_data.sort_values([bid_id_col, bid_price_col])
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]
    best_of_best_list = []
    bid_ids = data[bid_id_col].unique()
    for bid_id in bid_ids:
        bid_rows = bid_data[bid_data[bid_id_col] == bid_id]
        if bid_rows.empty:
            bid_row = data[data[bid_id_col] == bid_id].iloc[0]
            best_of_best_list.append({
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': bid_row[facility_col],
                'Incumbent': bid_row[incumbent_col],
                'Baseline Price': bid_row[baseline_price_col],
                'Bid Volume': bid_row[bid_volume_col],
                'Baseline Spend': bid_row['Baseline Spend'],
                'Awarded Supplier': 'No Bids',
                'Awarded Supplier Price': None,
                'Awarded Volume': None,
                'Awarded Supplier Spend': None,
                'Awarded Supplier Capacity': None,
                'Savings': None
            })
            logger.debug(f"No valid bids for Bid ID {bid_id}.")
            continue
        remaining_volume = bid_rows.iloc[0][bid_volume_col]
        split_index = 'A'
        for i, row in bid_rows.iterrows():
            supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else remaining_volume
            awarded_volume = min(remaining_volume, supplier_capacity)
            baseline_volume = awarded_volume
            baseline_spend = baseline_volume * row[baseline_price_col]
            best_of_best_spend = awarded_volume * row[bid_price_col]
            best_of_best_savings = baseline_spend - best_of_best_spend
            best_of_best_list.append({
                'Bid ID': row[bid_id_col],
                'Bid ID Split': split_index,
                'Facility': row[facility_col],
                'Incumbent': row[incumbent_col],
                'Baseline Price': row[baseline_price_col],
                'Bid Volume': baseline_volume,
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': row[supplier_name_col],
                'Awarded Supplier Price': row[bid_price_col],
                'Awarded Volume': awarded_volume,
                'Awarded Supplier Spend': best_of_best_spend,
                'Awarded Supplier Capacity': supplier_capacity,
                'Savings': best_of_best_savings
            })
            logger.debug(f"Best of Best analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume}")
            remaining_volume -= awarded_volume
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break
    best_of_best_df = pd.DataFrame(best_of_best_list)
    return best_of_best_df

# Function for Best of Best Excluding Suppliers analysis
def best_of_best_excluding_suppliers(data, column_mapping, excluded_conditions):
    """Perform 'Best of Best Excluding Suppliers' analysis."""
    logger.info("Starting Best of Best Excluding Suppliers analysis.")
    # Column mappings
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']
    facility_col = column_mapping['Facility']

    data[supplier_name_col] = data[supplier_name_col].str.title()
    data[incumbent_col] = data[incumbent_col].str.title()

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    # Apply exclusion rules
    for condition in excluded_conditions:
        supplier, field, logic, value, exclude_all = condition
        if exclude_all:
            data = data[data[supplier_name_col] != supplier]
            logger.debug(f"Excluding all bids from supplier {supplier}.")
        else:
            if logic == "Equal to":
                data = data[~((data[supplier_name_col] == supplier) & (data[field] == value))]
                logger.debug(f"Excluding bids from supplier {supplier} where {field} == {value}.")
            elif logic == "Not equal to":
                data = data[~((data[supplier_name_col] == supplier) & (data[field] != value))]
                logger.debug(f"Excluding bids from supplier {supplier} where {field} != {value}.")

    bid_data = data.loc[data['Valid Bid']]
    bid_data = bid_data.sort_values([bid_id_col, bid_price_col])
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]
    best_of_best_excl_list = []
    bid_ids = data[bid_id_col].unique()
    for bid_id in bid_ids:
        bid_rows = bid_data[bid_data[bid_id_col] == bid_id]
        if bid_rows.empty:
            bid_row = data[data[bid_id_col] == bid_id].iloc[0]
            baseline_spend = bid_row[bid_volume_col] * bid_row[baseline_price_col]
            best_of_best_excl_list.append({
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': bid_row[facility_col],
                'Incumbent': bid_row[incumbent_col],
                'Baseline Price': bid_row[baseline_price_col],
                'Bid Volume': bid_row[bid_volume_col],
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': 'Unallocated',
                'Awarded Supplier Price': None,
                'Awarded Volume': bid_row[bid_volume_col],
                'Awarded Supplier Spend': None,
                'Awarded Supplier Capacity': None,
                'Savings': None
            })
            logger.debug(f"All suppliers excluded or no valid bids for Bid ID {bid_id}. Marked as Unallocated.")
            continue
        remaining_volume = bid_rows.iloc[0][bid_volume_col]
        split_index = 'A'
        for i, row in bid_rows.iterrows():
            supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else remaining_volume
            awarded_volume = min(remaining_volume, supplier_capacity)
            baseline_volume = awarded_volume
            baseline_spend = baseline_volume * row[baseline_price_col]
            awarded_spend = awarded_volume * row[bid_price_col]
            savings = baseline_spend - awarded_spend
            best_of_best_excl_list.append({
                'Bid ID': row[bid_id_col],
                'Bid ID Split': split_index,
                'Facility': row[facility_col],
                'Incumbent': row[incumbent_col],
                'Baseline Price': row[baseline_price_col],
                'Bid Volume': baseline_volume,
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': row[supplier_name_col],
                'Awarded Supplier Price': row[bid_price_col],
                'Awarded Volume': awarded_volume,
                'Awarded Supplier Spend': awarded_spend,
                'Awarded Supplier Capacity': supplier_capacity,
                'Savings': savings
            })
            logger.debug(f"Best of Best Excl analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume}")
            remaining_volume -= awarded_volume
            if remaining_volume <= 0:
                break
            split_index = chr(ord(split_index) + 1)
    best_of_best_excl_df = pd.DataFrame(best_of_best_excl_list)
    return best_of_best_excl_df

# Function for As-Is Excluding Suppliers analysis
def as_is_excluding_suppliers_analysis(data, column_mapping, excluded_conditions):
    """Perform 'As-Is Excluding Suppliers' analysis with exclusion rules."""
    logger.info("Starting As-Is Excluding Suppliers analysis.")
    # Column mappings
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']
    facility_col = column_mapping['Facility']
    baseline_price_col = column_mapping['Baseline Price']

    data[supplier_name_col] = data[supplier_name_col].str.title()
    data[incumbent_col] = data[incumbent_col].str.title()

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    # Apply exclusion rules specific to this analysis
    for condition in excluded_conditions:
        supplier, field, logic, value, exclude_all = condition
        if exclude_all:
            data = data[data[supplier_name_col] != supplier]
            logger.debug(f"Excluding all bids from supplier {supplier} in As-Is Excluding Suppliers analysis.")
        else:
            if logic == "Equal to":
                data = data[~((data[supplier_name_col] == supplier) & (data[field] == value))]
                logger.debug(f"Excluding bids from supplier {supplier} where {field} == {value}.")
            elif logic == "Not equal to":
                data = data[~((data[supplier_name_col] == supplier) & (data[field] != value))]
                logger.debug(f"Excluding bids from supplier {supplier} where {field} != {value}.")

    bid_data = data.loc[data['Valid Bid']]
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]
    as_is_excl_list = []
    bid_ids = data[bid_id_col].unique()

    for bid_id in bid_ids:
        bid_rows = bid_data[bid_data[bid_id_col] == bid_id]
        all_rows = data[data[bid_id_col] == bid_id]
        incumbent = all_rows[incumbent_col].iloc[0]
        facility = all_rows[facility_col].iloc[0]
        baseline_price = all_rows[baseline_price_col].iloc[0]
        bid_volume = all_rows[bid_volume_col].iloc[0]
        baseline_spend = bid_volume * baseline_price

        # Check if incumbent is excluded
        incumbent_excluded = False
        for condition in excluded_conditions:
            supplier, field, logic, value, exclude_all = condition
            if supplier == incumbent and (exclude_all or (logic == "Equal to" and all_rows[field].iloc[0] == value) or (logic == "Not equal to" and all_rows[field].iloc[0] != value)):
                incumbent_excluded = True
                break

        if not incumbent_excluded:
            # Incumbent is not excluded
            incumbent_bid = bid_rows[bid_rows[supplier_name_col] == incumbent]
            if not incumbent_bid.empty:
                # Incumbent did bid
                row = incumbent_bid.iloc[0]
                supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else bid_volume
                awarded_volume = min(bid_volume, supplier_capacity)
                awarded_spend = awarded_volume * row[bid_price_col]
                savings = baseline_spend - awarded_spend

                as_is_excl_list.append({
                    'Bid ID': bid_id,
                    'Bid ID Split': 'A',
                    'Facility': facility,
                    'Incumbent': incumbent,
                    'Baseline Price': baseline_price,
                    'Bid Volume': awarded_volume,
                    'Baseline Spend': awarded_volume * baseline_price,
                    'Awarded Supplier': incumbent,
                    'Awarded Supplier Price': row[bid_price_col],
                    'Awarded Volume': awarded_volume,
                    'Awarded Supplier Spend': awarded_spend,
                    'Awarded Supplier Capacity': supplier_capacity,
                    'Savings': savings
                })
                logger.debug(f"As-Is Excl analysis for Bid ID {bid_id}: Awarded to incumbent.")

                remaining_volume = bid_volume - awarded_volume
                if remaining_volume > 0:
                    # Remaining volume is unallocated
                    as_is_excl_list.append({
                        'Bid ID': bid_id,
                        'Bid ID Split': 'B',
                        'Facility': facility,
                        'Incumbent': incumbent,
                        'Baseline Price': baseline_price,
                        'Bid Volume': remaining_volume,
                        'Baseline Spend': remaining_volume * baseline_price,
                        'Awarded Supplier': 'Unallocated',
                        'Awarded Supplier Price': None,
                        'Awarded Volume': remaining_volume,
                        'Awarded Supplier Spend': None,
                        'Awarded Supplier Capacity': None,
                        'Savings': None
                    })
                    logger.debug(f"Remaining volume for Bid ID {bid_id} is unallocated after awarding to incumbent.")
            else:
                # Incumbent did not bid or bid is invalid
                as_is_excl_list.append({
                    'Bid ID': bid_id,
                    'Bid ID Split': 'A',
                    'Facility': facility,
                    'Incumbent': incumbent,
                    'Baseline Price': baseline_price,
                    'Bid Volume': bid_volume,
                    'Baseline Spend': baseline_spend,
                    'Awarded Supplier': 'Unallocated',
                    'Awarded Supplier Price': None,
                    'Awarded Volume': bid_volume,
                    'Awarded Supplier Spend': None,
                    'Awarded Supplier Capacity': None,
                    'Savings': None
                })
                logger.debug(f"Incumbent did not bid or invalid bid for Bid ID {bid_id}. Entire volume is unallocated.")
        else:
            # Incumbent is excluded
            # Allocate to the lowest priced suppliers
            valid_bids = bid_rows[bid_rows[supplier_name_col] != incumbent]
            valid_bids = valid_bids.sort_values(by=bid_price_col)
            remaining_volume = bid_volume
            split_index = 'A'

            if valid_bids.empty:
                # No valid bids, mark as Unallocated
                as_is_excl_list.append({
                    'Bid ID': bid_id,
                    'Bid ID Split': split_index,
                    'Facility': facility,
                    'Incumbent': incumbent,
                    'Baseline Price': baseline_price,
                    'Bid Volume': bid_volume,
                    'Baseline Spend': baseline_spend,
                    'Awarded Supplier': 'Unallocated',
                    'Awarded Supplier Price': None,
                    'Awarded Volume': bid_volume,
                    'Awarded Supplier Spend': None,
                    'Awarded Supplier Capacity': None,
                    'Savings': None
                })
                logger.debug(f"No valid bids for Bid ID {bid_id} after exclusions. Entire volume is unallocated.")
                continue

            for _, row in valid_bids.iterrows():
                supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else remaining_volume
                awarded_volume = min(remaining_volume, supplier_capacity)
                awarded_spend = awarded_volume * row[bid_price_col]
                baseline_spend_allocated = awarded_volume * baseline_price
                savings = baseline_spend_allocated - awarded_spend

                as_is_excl_list.append({
                    'Bid ID': bid_id,
                    'Bid ID Split': split_index,
                    'Facility': facility,
                    'Incumbent': incumbent,
                    'Baseline Price': baseline_price,
                    'Bid Volume': awarded_volume,
                    'Baseline Spend': baseline_spend_allocated,
                    'Awarded Supplier': row[supplier_name_col],
                    'Awarded Supplier Price': row[bid_price_col],
                    'Awarded Volume': awarded_volume,
                    'Awarded Supplier Spend': awarded_spend,
                    'Awarded Supplier Capacity': supplier_capacity,
                    'Savings': savings
                })
                logger.debug(f"As-Is Excl analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume} to {row[supplier_name_col]}")

                remaining_volume -= awarded_volume
                if remaining_volume <= 0:
                    break
                split_index = chr(ord(split_index) + 1)

            if remaining_volume > 0:
                # Remaining volume is unallocated
                as_is_excl_list.append({
                    'Bid ID': bid_id,
                    'Bid ID Split': split_index,
                    'Facility': facility,
                    'Incumbent': incumbent,
                    'Baseline Price': baseline_price,
                    'Bid Volume': remaining_volume,
                    'Baseline Spend': remaining_volume * baseline_price,
                    'Awarded Supplier': 'Unallocated',
                    'Awarded Supplier Price': None,
                    'Awarded Volume': remaining_volume,
                    'Awarded Supplier Spend': None,
                    'Awarded Supplier Capacity': None,
                    'Savings': None
                })
                logger.debug(f"Remaining volume for Bid ID {bid_id} is unallocated after allocating to suppliers.")

    as_is_excl_df = pd.DataFrame(as_is_excl_list)
    return as_is_excl_df

# Bid Coverage Report Functions (Updated to use 'Awarded Supplier' directly)

# Function for Competitiveness Report
def competitiveness_report(data, column_mapping, group_by_field):
    """Generate Competitiveness Report with corrected calculations."""
    logger.info(f"Generating Competitiveness Report grouped by {group_by_field}.")

    # Extract column names from column_mapping
    bid_price_col = column_mapping['Bid Price']
    bid_id_col = column_mapping['Bid ID']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = 'Awarded Supplier'  # Use 'Awarded Supplier' directly

    # Prepare data
    suppliers = data[supplier_name_col].unique()
    total_suppliers = len(suppliers)

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    grouped = data.groupby(group_by_field)
    report_rows = []

    for group, group_data in grouped:
        unique_bid_ids = group_data[bid_id_col].unique()
        total_bid_ids = len(unique_bid_ids)
        possible_bids = total_suppliers * total_bid_ids

        bids_received = group_data[group_data['Valid Bid']].shape[0]

        bid_ids_with_no_bids = total_bid_ids - group_data[group_data['Valid Bid']][bid_id_col].nunique()

        bid_ids_multiple_bids = group_data[group_data['Valid Bid']].groupby(bid_id_col)[supplier_name_col].nunique()
        percent_multiple_bids = (bid_ids_multiple_bids > 1).sum() / total_bid_ids * 100 if total_bid_ids > 0 else 0

        # Incumbent not bidding
        bid_ids_incumbent_no_bid = []
        for bid_id in unique_bid_ids:
            bid_rows = group_data[group_data[bid_id_col] == bid_id]
            incumbent = bid_rows[incumbent_col].iloc[0]
            incumbent_bid = bid_rows[(bid_rows[supplier_name_col] == incumbent) & (bid_rows['Valid Bid'])]
            if incumbent_bid.empty:
                bid_ids_incumbent_no_bid.append(bid_id)
        num_incumbent_no_bid = len(bid_ids_incumbent_no_bid)
        bid_ids_incumbent_no_bid_list = ', '.join(map(str, bid_ids_incumbent_no_bid))

        report_rows.append({
            'Group': group,
            '# of Possible Bids': possible_bids,
            '# of Bids Received': bids_received,
            'Bid IDs with No Bids': bid_ids_with_no_bids,
            '% of Bid IDs with Multiple Bids': f"{percent_multiple_bids:.0f}%",
            '# of Bid IDs Where Incumbent Did Not Bid': num_incumbent_no_bid,
            'List of Bid IDs Where Incumbent Did Not Bid': bid_ids_incumbent_no_bid_list
        })

    report_df = pd.DataFrame(report_rows)
    return report_df

# Function for Supplier Coverage Report
def supplier_coverage_report(data, column_mapping, group_by_field):
    """Generate Supplier Coverage Report with All Bids and grouped tables."""
    logger.info(f"Generating Supplier Coverage Report grouped by {group_by_field}.")

    # Extract column names from column_mapping
    bid_price_col = column_mapping['Bid Price']
    bid_id_col = column_mapping['Bid ID']
    supplier_name_col = 'Awarded Supplier'  # Use 'Awarded Supplier' directly

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    total_bid_ids = data[bid_id_col].nunique()
    suppliers = data[supplier_name_col].unique()
    all_bids_rows = []
    for supplier in suppliers:
        bids_provided = data[(data[supplier_name_col] == supplier) & (data['Valid Bid'])][bid_id_col].nunique()
        coverage = (bids_provided / total_bid_ids) * 100 if total_bid_ids > 0 else 0
        all_bids_rows.append({
            'Supplier': supplier,
            '# of Bid IDs': total_bid_ids,
            '# of Bids Provided': bids_provided,
            '% Coverage': f"{coverage:.0f}%"
        })
    all_bids_df = pd.DataFrame(all_bids_rows)

    # Grouped Tables
    grouped_tables = {}
    groups = data[group_by_field].unique()
    for group in groups:
        group_data = data[data[group_by_field] == group]
        group_total_bid_ids = group_data[bid_id_col].nunique()
        group_rows = []
        for supplier in suppliers:
            bids_provided = group_data[(group_data[supplier_name_col] == supplier) & (group_data['Valid Bid'])][bid_id_col].nunique()
            coverage = (bids_provided / group_total_bid_ids) * 100 if group_total_bid_ids > 0 else 0
            group_rows.append({
                'Supplier': supplier,
                '# of Bid IDs': group_total_bid_ids,
                '# of Bids Provided': bids_provided,
                '% Coverage': f"{coverage:.0f}%"
            })
        group_df = pd.DataFrame(group_rows)
        grouped_tables[f"Supplier Coverage - {group}"] = group_df

    return {'Supplier Coverage - All Bids': all_bids_df, **grouped_tables}

# Function for Facility Coverage Report
def facility_coverage_report(data, column_mapping, group_by_field):
    """Generate Facility Coverage Report grouped by the specified field."""
    logger.info(f"Generating Facility Coverage Report grouped by {group_by_field}.")

    facility_col = column_mapping['Facility']
    supplier_name_col = 'Awarded Supplier'  # Use 'Awarded Supplier' directly
    bid_price_col = column_mapping['Bid Price']
    bid_id_col = column_mapping['Bid ID']

    facilities = data[facility_col].unique()
    suppliers = data[supplier_name_col].unique()
    report = pd.DataFrame({'Supplier': suppliers})
    report.set_index('Supplier', inplace=True)

    # Treat bids with Bid Price NaN or 0 as 'No Bid'
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    for facility in facilities:
        facility_bids = data.loc[(data[facility_col] == facility) & (data['Valid Bid'])]
        total_bid_ids = data[data[facility_col] == facility][bid_id_col].nunique()
        coverage = facility_bids.groupby(supplier_name_col)[bid_id_col].nunique() / total_bid_ids
        coverage = coverage.reindex(suppliers).fillna(0) * 100  # Ensure alignment with suppliers
        report[facility] = coverage
    report.reset_index(inplace=True)
    return report

# Function to handle Bid Coverage Report
def bid_coverage_report(data, column_mapping, variations, group_by_field):
    """Generate Bid Coverage Reports based on selected variations and grouping."""
    logger.info(f"Running Bid Coverage Report with variations: {variations} and grouping by {group_by_field}.")
    reports = {}
    if "Competitiveness Report" in variations:
        competitiveness = competitiveness_report(data, column_mapping, group_by_field)
        reports['Competitiveness Report'] = competitiveness
        logger.info("Competitiveness Report generated.")
    if "Supplier Coverage" in variations:
        supplier_coverage = supplier_coverage_report(data, column_mapping, group_by_field)
        reports.update(supplier_coverage)  # Include all tables
        logger.info("Supplier Coverage Report generated.")
    if "Facility Coverage" in variations:
        facility_coverage = facility_coverage_report(data, column_mapping, group_by_field)
        reports['Facility Coverage'] = facility_coverage
        logger.info("Facility Coverage Report generated.")
    return reports

def customizable_analysis(data, column_mapping):
    """Perform 'Customizable Analysis' and prepare data for Excel output."""
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    facility_col = column_mapping['Facility']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']

    # Ensure necessary columns are numeric
    data[bid_volume_col] = pd.to_numeric(data[bid_volume_col], errors='coerce')
    data[supplier_capacity_col] = pd.to_numeric(data[supplier_capacity_col], errors='coerce')
    data[bid_price_col] = pd.to_numeric(data[bid_price_col], errors='coerce')
    data[baseline_price_col] = pd.to_numeric(data[baseline_price_col], errors='coerce')

    # Calculate Savings
    data['Savings'] = (data[baseline_price_col] - data[bid_price_col]) * data[bid_volume_col]

    # Create Supplier Name with Bid Price
    data['Supplier Name with Bid Price'] = data[supplier_name_col] + " ($" + data[bid_price_col].round(2).astype(str) + ")"

    # Calculate Baseline Spend
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]

    # Get unique Bid IDs
    bid_ids = data[bid_id_col].unique()

 # Prepare the customizable analysis DataFrame
    customizable_list = []
    for bid_id in bid_ids:
        bid_row = data[data[bid_id_col] == bid_id].iloc[0]
        customizable_list.append({
            'Bid ID': bid_id,
            'Facility': bid_row[facility_col],
            'Incumbent': bid_row[incumbent_col],
            'Baseline Price': bid_row[baseline_price_col],
            'Bid Volume': bid_row[bid_volume_col],
            'Baseline Spend': bid_row['Baseline Spend'],
            'Awarded Supplier': '',  # To be selected via data validation in Excel
            'Supplier Name': '',     # New column added here
            'Awarded Supplier Price': None,  # Formula-based
            'Awarded Volume': None,  # Formula-based
            'Awarded Supplier Spend': None,  # Formula-based
            'Awarded Supplier Capacity': None,  # Formula-based
            'Savings': None  # Formula-based
        })
    customizable_df = pd.DataFrame(customizable_list)
    return customizable_df



# Project Management Functions
def get_user_projects(username):
    """Retrieve the list of projects for a given user."""
    user_dir = BASE_PROJECTS_DIR / username
    user_dir.mkdir(parents=True, exist_ok=True)
    projects = [p.name for p in user_dir.iterdir() if p.is_dir()]
    logger.info(f"Retrieved projects for user '{username}': {projects}")
    return projects

def create_project(username, project_name):
    """Create a new project with predefined subfolders."""
    user_dir = BASE_PROJECTS_DIR / username
    project_dir = user_dir / project_name
    if project_dir.exists():
        st.error(f"Project '{project_name}' already exists.")
        logger.warning(f"Attempted to create duplicate project '{project_name}' for user '{username}'.")
        return False
    try:
        project_dir.mkdir(parents=True)
        subfolders = ["Baseline", "Round 1 Analysis", "Round 2 Analysis", "Supplier Feedback", "Negotiations"]
        for subfolder in subfolders:
            (project_dir / subfolder).mkdir()
            logger.info(f"Created subfolder '{subfolder}' in project '{project_name}'.")
        st.success(f"Project '{project_name}' created successfully.")
        logger.info(f"User '{username}' created project '{project_name}'.")
        return True
    except Exception as e:
        st.error(f"Error creating project '{project_name}': {e}")
        logger.error(f"Error creating project '{project_name}': {e}")
        return False

def delete_project(username, project_name):
    """Delete an existing project."""
    user_dir = BASE_PROJECTS_DIR / username
    project_dir = user_dir / project_name
    if not project_dir.exists():
        st.error(f"Project '{project_name}' does not exist.")
        logger.warning(f"Attempted to delete non-existent project '{project_name}' for user '{username}'.")
        return False
    try:
        shutil.rmtree(project_dir)
        st.success(f"Project '{project_name}' deleted successfully.")
        logger.info(f"User '{username}' deleted project '{project_name}'.")
        return True
    except Exception as e:
        st.error(f"Error deleting project '{project_name}': {e}")
        logger.error(f"Error deleting project '{project_name}': {e}")
        return False

def apply_custom_css():
    """Apply custom CSS for styling the app."""
    st.markdown(
        f"""
        <style>
        /* Set a subtle background color */
        body {{
            background-color: #f0f2f6;
        }}

        /* Remove the default main menu and footer for a cleaner look */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}

        /* Style for the header */
        .header {{
            display: flex;
            align-items: center;
            padding: 10px 20px;
            background-color: #ffffff;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }}
        .header img {{
            height: 50px;
            margin-right: 20px;
        }}
        .header .page-title {{
            font-size: 24px;
            font-weight: bold;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Streamlit App
def main():
    # Apply custom CSS
    apply_custom_css()

    # Initialize session state variables
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'username' not in st.session_state:
        st.session_state.username = ''
    if 'merged_data' not in st.session_state:
        st.session_state.merged_data = None
    if 'original_merged_data' not in st.session_state:
        st.session_state.original_merged_data = None
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = {}
    if 'columns' not in st.session_state:
        st.session_state.columns = []
    if 'baseline_data' not in st.session_state:
        st.session_state.baseline_data = None
    if 'exclusions_bob' not in st.session_state:
        st.session_state.exclusions_bob = []
    if 'exclusions_ais' not in st.session_state:
        st.session_state.exclusions_ais = []
    if 'current_section' not in st.session_state:
        st.session_state.current_section = 'home'
    if 'selected_project' not in st.session_state:
        st.session_state.selected_project = None
    if 'selected_subfolder' not in st.session_state:
        st.session_state.selected_subfolder = None
    if 'customizable_grouping_column' not in st.session_state:
        st.session_state.customizable_grouping_column = None

    # Header with logo and page title
    st.markdown(
        f"""
        <div class="header">
            <img src="https://scfuturemakers.com/wp-content/uploads/2017/11/Georgia-Pacific_overview_video-854x480-c-default.jpg" alt="Logo">
            <div class="page-title">{st.session_state.current_section.capitalize()}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Right side: User Info and Logout button
    col1, col2 = st.columns([3, 1])

    with col1:
        pass  # Logo and page title are handled above

    with col2:
        if st.session_state.authenticated:
            st.markdown(f"<div style='text-align: right;'>**Welcome, {st.session_state.username}!**</div>", unsafe_allow_html=True)
            logout = st.button("Logout", key='logout_button')
            if logout:
                st.session_state.authenticated = False
                st.session_state.username = ''
                st.success("You have been logged out.")
                logger.info(f"User logged out successfully.")
        else:
            # Display a button to navigate to 'My Projects'
            login = st.button("Login to My Projects", key='login_button')
            if login:
                st.session_state.current_section = 'analysis'

    # Sidebar navigation using buttons
    st.sidebar.title('Navigation')

    # Define a function to handle navigation button clicks
    def navigate_to(section):
        st.session_state.current_section = section
        st.session_state.selected_project = None
        st.session_state.selected_subfolder = None

    # Create buttons in the sidebar for navigation
    if st.sidebar.button('Home'):
        navigate_to('home')
    if st.sidebar.button('Start a New Analysis'):
        navigate_to('upload')
    if st.sidebar.button('My Projects'):
        navigate_to('analysis')
    if st.sidebar.button('Settings'):
        navigate_to('settings')
    if st.sidebar.button('About'):
        navigate_to('about')

    # If in 'My Projects', display folder tree in sidebar
    if st.session_state.current_section == 'analysis' and st.session_state.authenticated:
        st.sidebar.markdown("---")
        st.sidebar.subheader("Your Projects")
        user_projects = get_user_projects(st.session_state.username)
        subfolders = ["Baseline", "Round 1 Analysis", "Round 2 Analysis", "Supplier Feedback", "Negotiations"]
        for project in user_projects:
            with st.sidebar.expander(project, expanded=False):
                for subfolder in subfolders:
                    button_label = f"{project}/{subfolder}"
                    # Use unique key
                    if st.button(button_label, key=f"{project}_{subfolder}_sidebar"):
                        st.session_state.current_section = 'project_folder'
                        st.session_state.selected_project = project
                        st.session_state.selected_subfolder = subfolder

    # Display content based on the current section
    section = st.session_state.current_section

    if section == 'home':
        st.write("Welcome to the Scourcing COE Analysis Tool.")

        if not st.session_state.authenticated:
            st.write("You are not logged in. Please navigate to 'My Projects' to log in and access your projects.")
        else:
            st.write(f"Welcome, {st.session_state.username}!")

    elif section == 'analysis':
        st.title('My Projects')
        if not st.session_state.authenticated:
            st.header("Log In to Access My Projects")
            # Create a login form
            with st.form(key='login_form'):
                username = st.text_input("Username")
                password = st.text_input("Password", type='password')
                submit_button = st.form_submit_button(label='Login')

            if submit_button:
                if 'credentials' in config and 'usernames' in config['credentials']:
                    if username in config['credentials']['usernames']:
                        stored_hashed_password = config['credentials']['usernames'][username]['password']
                        # Verify the entered password against the stored hashed password
                        if bcrypt.checkpw(password.encode('utf-8'), stored_hashed_password.encode('utf-8')):
                            st.session_state.authenticated = True
                            st.session_state.username = username
                            st.success(f"Logged in as {username}")
                            logger.info(f"User {username} logged in successfully.")
                        else:
                            st.error("Incorrect password. Please try again.")
                            logger.warning(f"Incorrect password attempt for user {username}.")
                    else:
                        st.error("Username not found. Please check and try again.")
                        logger.warning(f"Login attempt with unknown username: {username}")
                else:
                    st.error("Authentication configuration is missing.")
                    logger.error("Authentication configuration is missing in 'config.yaml'.")

        else:
            # Project Management UI
            st.subheader("Manage Your Projects")
            col_start, col_delete = st.columns([1, 1])

            with col_start:
                # Start a New Project Form
                with st.form(key='create_project_form'):
                    st.markdown("### Create New Project")
                    new_project_name = st.text_input("Enter Project Name")
                    create_project_button = st.form_submit_button(label='Create Project')
                    if create_project_button:
                        if new_project_name.strip() == "":
                            st.error("Project name cannot be empty.")
                            logger.warning("Attempted to create a project with an empty name.")
                        else:
                            success = create_project(st.session_state.username, new_project_name.strip())
                            if success:
                                st.success(f"Project '{new_project_name.strip()}' created successfully.")

            with col_delete:
                # Delete a Project Form
                with st.form(key='delete_project_form'):
                    st.markdown("### Delete Existing Project")
                    projects = get_user_projects(st.session_state.username)
                    if projects:
                        project_to_delete = st.selectbox("Select Project to Delete", projects)
                        confirm_delete = st.form_submit_button(label='Confirm Delete')
                        if confirm_delete:
                            confirm = st.checkbox("I confirm that I want to delete this project.", key='confirm_delete_checkbox')
                            if confirm:
                                success = delete_project(st.session_state.username, project_to_delete)
                                if success:
                                    st.success(f"Project '{project_to_delete}' deleted successfully.")
                            else:
                                st.warning("Deletion not confirmed.")
                    else:
                        st.info("No projects available to delete.")
                        logger.info("No projects available to delete for the user.")

            st.markdown("---")
            st.subheader("Your Projects")

            projects = get_user_projects(st.session_state.username)
            if projects:
                for project in projects:
                    with st.container():
                        st.markdown(f"### {project}")
                        subfolders = ["Baseline", "Round 1 Analysis", "Round 2 Analysis", "Supplier Feedback", "Negotiations"]
                        cols = st.columns(len(subfolders))
                        for idx, subfolder in enumerate(subfolders):
                            button_label = f"{project}/{subfolder}"
                            # Use unique key
                            if cols[idx].button(button_label, key=f"{project}_{subfolder}_main"):
                                st.session_state.current_section = 'project_folder'
                                st.session_state.selected_project = project
                                st.session_state.selected_subfolder = subfolder
            else:
                st.info("No projects found. Start by creating a new project.")

    elif section == 'upload':
        st.title('Start a New Analysis')
        st.write("Upload your data here.")

        # Select data input method
        data_input_method = st.radio(
            "Select Data Input Method",
            ('Separate Bid & Baseline files', 'Merged Data'),
            index=0,
            key='data_input_method'
        )

        # Use the selected method in your code
        if data_input_method == 'Separate Bid & Baseline files':
            # Existing steps for separate files
            st.header("Upload Baseline and Bid Files")

            # Upload baseline file
            baseline_file = st.file_uploader("Upload Baseline Sheet", type=["xlsx"])

            # Sheet selection for Baseline File
            baseline_sheet = None
            if baseline_file:
                try:
                    excel_baseline = pd.ExcelFile(baseline_file, engine='openpyxl')
                    baseline_sheet = st.selectbox(
                        "Select Baseline Sheet",
                        excel_baseline.sheet_names,
                        key='baseline_sheet_selection'
                    )
                except Exception as e:
                    st.error(f"Error reading baseline file: {e}")
                    logger.error(f"Error reading baseline file: {e}")

            num_files = st.number_input("Number of Bid Sheets to Upload", min_value=1, step=1)

            bid_files_suppliers = []
            for i in range(int(num_files)):
                bid_file = st.file_uploader(f"Upload Bid Sheet {i + 1}", type=["xlsx"], key=f'bid_file_{i}')
                supplier_name = st.text_input(f"Supplier Name for Bid Sheet {i + 1}", key=f'supplier_name_{i}')
                # Sheet selection for each Bid File
                bid_sheet = None
                if bid_file and supplier_name:
                    try:
                        excel_bid = pd.ExcelFile(bid_file, engine='openpyxl')
                        bid_sheet = st.selectbox(
                            f"Select Sheet for Bid Sheet {i + 1}",
                            excel_bid.sheet_names,
                            key=f'bid_sheet_selection_{i}'
                        )
                    except Exception as e:
                        st.error(f"Error reading Bid Sheet {i + 1}: {e}")
                        logger.error(f"Error reading Bid Sheet {i + 1}: {e}")
                if bid_file and supplier_name and bid_sheet:
                    bid_files_suppliers.append((bid_file, supplier_name, bid_sheet))
                    logger.info(f"Uploaded Bid Sheet {i + 1} for supplier '{supplier_name}' with sheet '{bid_sheet}'.")

            # Merge Data
            if st.button("Merge Data"):
                if validate_uploaded_file(baseline_file) and bid_files_suppliers:
                    if not baseline_sheet:
                        st.error("Please select a sheet for the baseline file.")
                    else:
                        merged_data = start_process(baseline_file, baseline_sheet, bid_files_suppliers)
                        if merged_data is not None:
                            st.session_state.merged_data = merged_data
                            st.session_state.original_merged_data = merged_data.copy()
                            st.session_state.columns = list(merged_data.columns)
                            st.session_state.baseline_data = load_baseline_data(baseline_file, baseline_sheet)
                            st.success("Data merged successfully. Please map the columns for analysis.")
                            logger.info("Data merged successfully.")

        else:
            # For Merged Data input method
            st.header("Upload Merged Data File")
            merged_file = st.file_uploader("Upload Merged Data File", type=["xlsx"], key='merged_data_file')
            merged_sheet = None
            if merged_file:
                try:
                    excel_merged = pd.ExcelFile(merged_file, engine='openpyxl')
                    merged_sheet = st.selectbox(
                        "Select Sheet",
                        excel_merged.sheet_names,
                        key='merged_sheet_selection'
                    )
                except Exception as e:
                    st.error(f"Error reading merged data file: {e}")
                    logger.error(f"Error reading merged data file: {e}")

            if merged_file and merged_sheet:
                try:
                    merged_data = pd.read_excel(merged_file, sheet_name=merged_sheet, engine='openpyxl')
                    # Normalize columns
                    merged_data = normalize_columns(merged_data)
                    st.session_state.merged_data = merged_data
                    st.session_state.original_merged_data = merged_data.copy()
                    st.session_state.columns = list(merged_data.columns)

                    st.success("Merged data loaded successfully. Please map the columns for analysis.")
                    logger.info("Merged data loaded successfully.")
                except Exception as e:
                    st.error(f"Error loading merged data: {e}")
                    logger.error(f"Error loading merged data: {e}")

        # Proceed to Column Mapping if merged_data is available
        if st.session_state.merged_data is not None:
            required_columns = ['Bid ID', 'Incumbent', 'Facility', 'Baseline Price',
                                'Bid Volume', 'Bid Price', 'Supplier Capacity', 'Supplier Name']

            # Initialize column_mapping if not already initialized
            if 'column_mapping' not in st.session_state:
                st.session_state.column_mapping = {}

            # Map columns
            st.write("Map the following columns:")
            for col in required_columns:
                st.session_state.column_mapping[col] = st.selectbox(
                    f"Select Column for {col}",
                    st.session_state.merged_data.columns,
                    key=f"{col}_mapping"
                )

            # After mapping, check if all required columns are mapped
            missing_columns = [col for col in required_columns if col not in st.session_state.column_mapping or st.session_state.column_mapping[col] not in st.session_state.merged_data.columns]
            if missing_columns:
                st.error(f"The following required columns are not mapped or do not exist in the data: {', '.join(missing_columns)}")
            else:
                # After mapping, set 'Awarded Supplier' automatically
                st.session_state.merged_data['Awarded Supplier'] = st.session_state.merged_data[st.session_state.column_mapping['Supplier Name']]


            analyses_to_run = st.multiselect("Select Scenario Analyses to Run", [
                "As-Is",
                "Best of Best",
                "Best of Best Excluding Suppliers",
                "As-Is Excluding Suppliers",
                "Bid Coverage Report",
                "Customizable Analysis"
            ])


            # Exclusion rules for Best of Best Excluding Suppliers
            if "Best of Best Excluding Suppliers" in analyses_to_run:
                with st.expander("Configure Exclusion Rules for Best of Best Excluding Suppliers"):
                    st.header("Exclusion Rules for Best of Best Excluding Suppliers")

                    supplier_name = st.selectbox("Select Supplier to Exclude", st.session_state.merged_data['Awarded Supplier'].unique(), key="supplier_name_excl_bob")
                    field = st.selectbox("Select Field for Rule", st.session_state.merged_data.columns, key="field_excl_bob")
                    logic = st.selectbox("Select Logic (Equal to or Not equal to)", ["Equal to", "Not equal to"], key="logic_excl_bob")
                    value = st.selectbox("Select Value", st.session_state.merged_data[field].unique(), key="value_excl_bob")
                    exclude_all = st.checkbox("Exclude from all Bid IDs", key="exclude_all_excl_bob")

                    if st.button("Add Exclusion Rule", key="add_excl_bob"):
                        if 'exclusions_bob' not in st.session_state:
                            st.session_state.exclusions_bob = []
                        st.session_state.exclusions_bob.append((supplier_name, field, logic, value, exclude_all))
                        logger.debug(f"Added exclusion rule for BOB Excl Suppliers: {supplier_name}, {field}, {logic}, {value}, Exclude All: {exclude_all}")

                    if st.button("Clear Exclusion Rules", key="clear_excl_bob"):
                        st.session_state.exclusions_bob = []
                        logger.debug("Cleared all exclusion rules for BOB Excl Suppliers.")

                    if 'exclusions_bob' in st.session_state and st.session_state.exclusions_bob:
                        st.write("Current Exclusion Rules for Best of Best Excluding Suppliers:")
                        for i, excl in enumerate(st.session_state.exclusions_bob):
                            st.write(f"{i + 1}. Supplier: {excl[0]}, Field: {excl[1]}, Logic: {excl[2]}, Value: {excl[3]}, Exclude All: {excl[4]}")

            # Exclusion rules for As-Is Excluding Suppliers
            if "As-Is Excluding Suppliers" in analyses_to_run:
                with st.expander("Configure Exclusion Rules for As-Is Excluding Suppliers"):
                    st.header("Exclusion Rules for As-Is Excluding Suppliers")

                    supplier_name_ais = st.selectbox("Select Supplier to Exclude", st.session_state.merged_data['Awarded Supplier'].unique(), key="supplier_name_excl_ais")
                    field_ais = st.selectbox("Select Field for Rule", st.session_state.merged_data.columns, key="field_excl_ais")
                    logic_ais = st.selectbox("Select Logic (Equal to or Not equal to)", ["Equal to", "Not equal to"], key="logic_excl_ais")
                    value_ais = st.selectbox("Select Value", st.session_state.merged_data[field_ais].unique(), key="value_excl_ais")
                    exclude_all_ais = st.checkbox("Exclude from all Bid IDs", key="exclude_all_excl_ais")

                    if st.button("Add Exclusion Rule", key="add_excl_ais"):
                        if 'exclusions_ais' not in st.session_state:
                            st.session_state.exclusions_ais = []
                        st.session_state.exclusions_ais.append((supplier_name_ais, field_ais, logic_ais, value_ais, exclude_all_ais))
                        logger.debug(f"Added exclusion rule for As-Is Excl Suppliers: {supplier_name_ais}, {field_ais}, {logic_ais}, {value_ais}, Exclude All: {exclude_all_ais}")

                    if st.button("Clear Exclusion Rules", key="clear_excl_ais"):
                        st.session_state.exclusions_ais = []
                        logger.debug("Cleared all exclusion rules for As-Is Excl Suppliers.")

                    if 'exclusions_ais' in st.session_state and st.session_state.exclusions_ais:
                        st.write("Current Exclusion Rules for As-Is Excluding Suppliers:")
                        for i, excl in enumerate(st.session_state.exclusions_ais):
                            st.write(f"{i + 1}. Supplier: {excl[0]}, Field: {excl[1]}, Logic: {excl[2]}, Value: {excl[3]}, Exclude All: {excl[4]}")

            # Bid Coverage Report Configuration
            if "Bid Coverage Report" in analyses_to_run:
                with st.expander("Configure Bid Coverage Report"):
                    st.header("Bid Coverage Report Configuration")

                    # Select variations
                    bid_coverage_variations = st.multiselect("Select Bid Coverage Report Variations", [
                        "Competitiveness Report",
                        "Supplier Coverage",
                        "Facility Coverage"
                    ], key="bid_coverage_variations")

                    # Select group by field
                    group_by_field = st.selectbox("Group by", st.session_state.merged_data.columns, key="bid_coverage_group_by")


            if "Customizable Analysis" in analyses_to_run:
                with st.expander("Configure Customizable Analysis"):
                    st.header("Customizable Analysis Configuration")
                    # Select Grouping Column
                    grouping_column_raw = st.selectbox(
                        "Select Grouping Column",
                        st.session_state.merged_data.columns,
                        key="customizable_grouping_column"
                    )
                    
                    # Check if the selected grouping column is already mapped
                    grouping_column_already_mapped = False
                    grouping_column_mapped = grouping_column_raw  # Default to raw name
                    for standard_col, mapped_col in st.session_state.column_mapping.items():
                        if mapped_col == grouping_column_raw:
                            grouping_column_mapped = standard_col
                            grouping_column_already_mapped = True
                            break
                    
                    # Store both the raw and mapped grouping column names and the flag
                    st.session_state.grouping_column_raw = grouping_column_raw
                    st.session_state.grouping_column_mapped = grouping_column_mapped
                    st.session_state.grouping_column_already_mapped = grouping_column_already_mapped

            if st.button("Run Analysis"):
                with st.spinner("Running analysis..."):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        baseline_data = st.session_state.baseline_data
                        original_merged_data = st.session_state.original_merged_data

                        # As-Is Analysis
                        if "As-Is" in analyses_to_run:
                            as_is_df = as_is_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                            as_is_df = add_missing_bid_ids(as_is_df, original_merged_data, st.session_state.column_mapping, 'As-Is')
                            as_is_df.to_excel(writer, sheet_name='As-Is', index=False)
                            logger.info("As-Is analysis completed.")

                        # Best of Best Analysis
                        if "Best of Best" in analyses_to_run:
                            best_of_best_df = best_of_best_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                            best_of_best_df = add_missing_bid_ids(best_of_best_df, original_merged_data, st.session_state.column_mapping, 'Best of Best')
                            best_of_best_df.to_excel(writer, sheet_name='Best of Best', index=False)
                            logger.info("Best of Best analysis completed.")

                        # Best of Best Excluding Suppliers Analysis
                        if "Best of Best Excluding Suppliers" in analyses_to_run:
                            exclusions_list_bob = st.session_state.exclusions_bob if 'exclusions_bob' in st.session_state else []
                            best_of_best_excl_df = best_of_best_excluding_suppliers(st.session_state.merged_data, st.session_state.column_mapping, exclusions_list_bob)
                            best_of_best_excl_df = add_missing_bid_ids(best_of_best_excl_df, original_merged_data, st.session_state.column_mapping, 'BOB Excl Suppliers')
                            best_of_best_excl_df.to_excel(writer, sheet_name='BOB Excl Suppliers', index=False)
                            logger.info("Best of Best Excluding Suppliers analysis completed.")

                        # As-Is Excluding Suppliers Analysis
                        if "As-Is Excluding Suppliers" in analyses_to_run:
                            exclusions_list_ais = st.session_state.exclusions_ais if 'exclusions_ais' in st.session_state else []
                            as_is_excl_df = as_is_excluding_suppliers_analysis(st.session_state.merged_data, st.session_state.column_mapping, exclusions_list_ais)
                            as_is_excl_df = add_missing_bid_ids(as_is_excl_df, original_merged_data, st.session_state.column_mapping, 'As-Is Excl Suppliers')
                            as_is_excl_df.to_excel(writer, sheet_name='As-Is Excl Suppliers', index=False)
                            logger.info("As-Is Excluding Suppliers analysis completed.")

                        # Bid Coverage Report Processing
                        if "Bid Coverage Report" in analyses_to_run:
                            variations = st.session_state.bid_coverage_variations if 'bid_coverage_variations' in st.session_state else []
                            group_by_field = st.session_state.bid_coverage_group_by if 'bid_coverage_group_by' in st.session_state else st.session_state.merged_data.columns[0]
                            if variations:
                                bid_coverage_reports = bid_coverage_report(st.session_state.merged_data, st.session_state.column_mapping, variations, group_by_field)

                                # Initialize startrow for Supplier Coverage sheet
                                supplier_coverage_startrow = 0

                                for report_name, report_df in bid_coverage_reports.items():
                                    if "Supplier Coverage" in report_name:
                                        sheet_name = "Supplier Coverage"
                                        if sheet_name not in writer.sheets:
                                            # Create the worksheet
                                            report_df.to_excel(writer, sheet_name=sheet_name, startrow=supplier_coverage_startrow, index=False)
                                            supplier_coverage_startrow += len(report_df) + 2  # +2 for one blank row and one for the header
                                        else:
                                            # Write the report name as a header
                                            worksheet = writer.sheets[sheet_name]
                                            worksheet.cell(row=supplier_coverage_startrow + 1, column=1, value=report_name)
                                            supplier_coverage_startrow += 1

                                            report_df.to_excel(writer, sheet_name=sheet_name, startrow=supplier_coverage_startrow, index=False)
                                            supplier_coverage_startrow += len(report_df) + 2  # +2 for one blank row and one for the header

                                        logger.info(f"{report_name} added to sheet '{sheet_name}'.")
                                    else:
                                        # Clean sheet name by replacing spaces
                                        sheet_name_clean = report_name.replace(" ", "_")
                                        # Ensure sheet name is within Excel's limit of 31 characters
                                        if len(sheet_name_clean) > 31:
                                            sheet_name_clean = sheet_name_clean[:31]
                                        report_df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                                        logger.info(f"{report_name} generated and added to Excel.")
                            else:
                                st.warning("No Bid Coverage Report variations selected.")


                        if "Customizable Analysis" in analyses_to_run:
                            # Generate the customizable analysis DataFrame
                            customizable_df = customizable_analysis(st.session_state.merged_data, st.session_state.column_mapping)

                            # Define raw and mapped grouping column names
                            grouping_column_raw = st.session_state.grouping_column_raw  # e.g., 'Facility'
                            grouping_column_mapped = st.session_state.grouping_column_mapped  # e.g., 'Plant'

                            # If grouping column was not already mapped, add it to customizable_df
                            if not st.session_state.grouping_column_already_mapped:
                                if grouping_column_mapped not in customizable_df.columns:
                                    # Add the grouping column next to 'Bid ID'
                                    bid_id_idx = customizable_df.columns.get_loc('Bid ID')
                                    customizable_df.insert(bid_id_idx + 1, grouping_column_mapped, '')

                            # Write 'Customizable Template' sheet
                            customizable_df.to_excel(writer, sheet_name='Customizable Template', index=False)

                            # Prepare 'Customizable Reference' DataFrame
                            customizable_reference_df = st.session_state.merged_data.copy()

                            # If grouping column was not already mapped, add it to customizable_reference_df
                            if not st.session_state.grouping_column_already_mapped:
                                if grouping_column_mapped not in customizable_reference_df.columns:
                                    # Use 'grouping_column_mapped' instead of 'grouping_column_raw' to access the actual column name
                                    bid_id_to_grouping = st.session_state.merged_data.set_index(
                                        st.session_state.column_mapping['Bid ID']
                                    )[grouping_column_mapped].to_dict()
                                    customizable_reference_df[grouping_column_mapped] = customizable_reference_df[
                                        st.session_state.column_mapping['Bid ID']
                                    ].map(bid_id_to_grouping)

                            # Write 'Customizable Reference' sheet
                            customizable_reference_df.to_excel(writer, sheet_name='Customizable Reference', index=False)
                            logger.info("Customizable Analysis data prepared.")

                            # ======= Ensure 'Scenario Converter' Sheet Exists =======
                            if 'Scenario Converter' not in writer.book.sheetnames:
                                # Create 'Scenario Converter' sheet with 'Bid ID' and empty scenario columns
                                bid_ids = st.session_state.merged_data[st.session_state.column_mapping['Bid ID']].unique()
                                scenario_converter_df = pd.DataFrame({'Bid ID': bid_ids})
                                for i in range(1, 8):
                                    scenario_converter_df[f'Scenario {i}'] = ''  # Initialize empty columns
                                scenario_converter_df.to_excel(writer, sheet_name='Scenario Converter', index=False)
                                logger.info("'Scenario Converter' sheet created.")
                            else:
                                logger.info("'Scenario Converter' sheet already exists.")

                            # ======= Ensure Grouping Column Exists in 'merged_data' =======
                            if not st.session_state.grouping_column_already_mapped:
                                if grouping_column_mapped not in st.session_state.merged_data.columns:
                                    # Use 'grouping_column_mapped' instead of 'grouping_column_raw' to access the actual column name
                                    st.session_state.merged_data[grouping_column_mapped] = st.session_state.merged_data[
                                        st.session_state.column_mapping['Bid ID']
                                    ].map(bid_id_to_grouping)
                                    logger.info(f"Added grouping column '{grouping_column_mapped}' to 'merged_data'.")

                            # ======= Access the Workbook and Sheets =======
                            workbook = writer.book
                            customizable_template_sheet = workbook['Customizable Template']
                            customizable_reference_sheet = workbook['Customizable Reference']
                            scenario_converter_sheet = workbook['Scenario Converter']

                            # Get the max row numbers
                            max_row_template = customizable_template_sheet.max_row
                            max_row_reference = customizable_reference_sheet.max_row
                            max_row_converter = scenario_converter_sheet.max_row

                            # Map column names to letters in 'Customizable Reference' sheet
                            reference_col_letter = {cell.value: get_column_letter(cell.column) for cell in customizable_reference_sheet[1]}
                            # Map column names to letters in 'Customizable Template' sheet
                            template_col_letter = {cell.value: get_column_letter(cell.column) for cell in customizable_template_sheet[1]}

                            # Get the column letters for the grouping column in 'Customizable Reference' sheet
                            ref_grouping_col = reference_col_letter.get(grouping_column_mapped)
                            # Get the column letters for the grouping column in 'Customizable Template' sheet
                            temp_grouping_col = template_col_letter.get(grouping_column_mapped)

                            # Update 'Supplier Name' formulas in 'Customizable Template' sheet
                            awarded_supplier_col_letter = template_col_letter['Awarded Supplier']
                            supplier_name_col_letter = template_col_letter['Supplier Name']

                            for row in range(2, max_row_template + 1):
                                awarded_supplier_cell = f"{awarded_supplier_col_letter}{row}"
                                supplier_name_cell = f"{supplier_name_col_letter}{row}"
                                # Corrected formula syntax with closing parenthesis
                                formula_supplier_name = f'=IF({awarded_supplier_cell}<>"", LEFT({awarded_supplier_cell}, FIND("(", {awarded_supplier_cell}) - 2), "")'
                                customizable_template_sheet[supplier_name_cell].value = formula_supplier_name

                            # Create supplier lists per Bid ID in a hidden sheet
                            if 'SupplierLists' not in workbook.sheetnames:
                                supplier_list_sheet = workbook.create_sheet("SupplierLists")
                            else:
                                supplier_list_sheet = workbook['SupplierLists']

                            bid_id_col_reference = st.session_state.column_mapping['Bid ID']
                            supplier_name_with_bid_price_col_reference = st.session_state.column_mapping.get('Supplier Name with Bid Price', 'Supplier Name with Bid Price')
                            bid_id_supplier_list_ranges = {}
                            current_row = 1

                            bid_ids = st.session_state.merged_data[bid_id_col_reference].unique()
                            data = st.session_state.merged_data

                            for bid_id in bid_ids:
                                bid_data = data[data[bid_id_col_reference] == bid_id]
                                bid_data_filtered = bid_data[
                                    (bid_data[st.session_state.column_mapping['Bid Price']].notna()) &
                                    (bid_data[st.session_state.column_mapping['Bid Price']] != 0)
                                ]
                                if not bid_data_filtered.empty:
                                    bid_data_sorted = bid_data_filtered.sort_values(by=st.session_state.column_mapping['Bid Price'])
                                    suppliers = bid_data_sorted[supplier_name_with_bid_price_col_reference].dropna().tolist()
                                    start_row = current_row
                                    for supplier in suppliers:
                                        supplier_list_sheet.cell(row=current_row, column=1, value=supplier)
                                        current_row += 1
                                    end_row = current_row - 1
                                    bid_id_supplier_list_ranges[bid_id] = (start_row, end_row)
                                    current_row += 1  # Empty row for separation
                                else:
                                    bid_id_supplier_list_ranges[bid_id] = None

                            # Hide the 'SupplierLists' sheet
                            supplier_list_sheet.sheet_state = 'hidden'

                            # Set data validation and formulas in 'Customizable Template' sheet
                            ref_sheet_name = 'Customizable Reference'
                            temp_sheet_name = 'Customizable Template'

                            ref_bid_id_col = reference_col_letter[st.session_state.column_mapping['Bid ID']]
                            ref_supplier_name_col = reference_col_letter[st.session_state.column_mapping['Supplier Name']]
                            ref_bid_price_col = reference_col_letter[st.session_state.column_mapping['Bid Price']]
                            ref_supplier_capacity_col = reference_col_letter[st.session_state.column_mapping['Supplier Capacity']]
                            ref_supplier_name_with_bid_price_col = reference_col_letter.get(supplier_name_with_bid_price_col_reference, 'A')  # Default to 'A' if not found

                            temp_bid_id_col = template_col_letter['Bid ID']
                            temp_supplier_name_col = template_col_letter['Supplier Name']
                            temp_bid_volume_col = template_col_letter['Bid Volume']
                            temp_baseline_price_col = template_col_letter['Baseline Price']

                            for row in range(2, max_row_template + 1):
                                bid_id_cell = f"{temp_bid_id_col}{row}"
                                awarded_supplier_cell = f"{awarded_supplier_col_letter}{row}"
                                supplier_name_cell = f"{temp_supplier_name_col}{row}"
                                bid_volume_cell = f"{temp_bid_volume_col}{row}"
                                baseline_price_cell = f"{temp_baseline_price_col}{row}"
                                grouping_cell = f"{temp_grouping_col}{row}" if temp_grouping_col else None

                                bid_id_value = customizable_template_sheet[bid_id_cell].value

                                if bid_id_value in bid_id_supplier_list_ranges and bid_id_supplier_list_ranges[bid_id_value]:
                                    start_row, end_row = bid_id_supplier_list_ranges[bid_id_value]
                                    supplier_list_range = f"'SupplierLists'!$A${start_row}:$A${end_row}"

                                    # Set data validation for 'Awarded Supplier'
                                    dv = DataValidation(type="list", formula1=f"{supplier_list_range}", allow_blank=True)
                                    customizable_template_sheet.add_data_validation(dv)
                                    dv.add(customizable_template_sheet[awarded_supplier_cell])

                                    # Construct the formulas using SUMIFS
                                    # Awarded Supplier Price
                                    formula_price = (
                                        f"=IFERROR(SUMIFS('{ref_sheet_name}'!{ref_bid_price_col}:{ref_bid_price_col}, "
                                        f"'{ref_sheet_name}'!{ref_bid_id_col}:{ref_bid_id_col}, {bid_id_cell}, "
                                        f"'{ref_sheet_name}'!{ref_supplier_name_col}:{ref_supplier_name_col}, {supplier_name_cell}), \"\")"
                                    )
                                    awarded_supplier_price_cell = f"{template_col_letter['Awarded Supplier Price']}{row}"
                                    customizable_template_sheet[awarded_supplier_price_cell].value = formula_price

                                    # Awarded Supplier Capacity
                                    formula_supplier_capacity = (
                                        f"=IFERROR(SUMIFS('{ref_sheet_name}'!{ref_supplier_capacity_col}:{ref_supplier_capacity_col}, "
                                        f"'{ref_sheet_name}'!{ref_bid_id_col}:{ref_bid_id_col}, {bid_id_cell}, "
                                        f"'{ref_sheet_name}'!{ref_supplier_name_col}:{ref_supplier_name_col}, {supplier_name_cell}), \"\")"
                                    )
                                    awarded_supplier_capacity_cell = f"{template_col_letter['Awarded Supplier Capacity']}{row}"
                                    customizable_template_sheet[awarded_supplier_capacity_cell].value = formula_supplier_capacity

                                    # Awarded Volume
                                    awarded_volume_cell = f"{template_col_letter['Awarded Volume']}{row}"
                                    formula_awarded_volume = f"=IFERROR(MIN({bid_volume_cell}, {awarded_supplier_capacity_cell}), \"\")"
                                    customizable_template_sheet[awarded_volume_cell].value = formula_awarded_volume

                                    # Awarded Supplier Spend
                                    awarded_supplier_spend_cell = f"{template_col_letter['Awarded Supplier Spend']}{row}"
                                    formula_spend = f"=IF({awarded_supplier_price_cell}<>\"\", {awarded_supplier_price_cell}*{awarded_volume_cell}, \"\")"
                                    customizable_template_sheet[awarded_supplier_spend_cell].value = formula_spend

                                    # Savings
                                    savings_cell = f"{template_col_letter['Savings']}{row}"
                                    formula_savings = f"=IF({awarded_supplier_price_cell}<>\"\", ({baseline_price_cell}-{awarded_supplier_price_cell})*{awarded_volume_cell}, \"\")"
                                    customizable_template_sheet[savings_cell].value = formula_savings

                                    # Baseline Spend
                                    baseline_spend_cell = f"{template_col_letter['Baseline Spend']}{row}"
                                    formula_baseline_spend = f"={baseline_price_cell}*{bid_volume_cell}"
                                    customizable_template_sheet[baseline_spend_cell].value = formula_baseline_spend

                                    # If grouping column exists, populate it
                                    if grouping_cell and ref_grouping_col:
                                        col_index = column_index_from_string(ref_grouping_col) - column_index_from_string(ref_bid_id_col) + 1
                                        formula_grouping = (
                                            f"=IFERROR(VLOOKUP({bid_id_cell}, '{ref_sheet_name}'!{ref_bid_id_col}:{ref_grouping_col}, "
                                            f"{col_index}, FALSE), \"\")"
                                        )
                                        customizable_template_sheet[grouping_cell].value = formula_grouping
                                else:
                                    pass  # No valid suppliers for this Bid ID

                            # Apply formatting to 'Customizable Reference' sheet
                            currency_columns_reference = ['Baseline Spend', 'Savings', st.session_state.column_mapping['Bid Price'], st.session_state.column_mapping['Baseline Price']]
                            number_columns_reference = [st.session_state.column_mapping['Bid Volume'], st.session_state.column_mapping['Supplier Capacity']]

                            for col_name in currency_columns_reference:
                                col_letter = reference_col_letter.get(col_name)
                                if col_letter:
                                    for row_num in range(2, max_row_reference + 1):
                                        cell = customizable_reference_sheet[f"{col_letter}{row_num}"]
                                        cell.number_format = '$#,##0.00'

                            for col_name in number_columns_reference:
                                col_letter = reference_col_letter.get(col_name)
                                if col_letter:
                                    for row_num in range(2, max_row_reference + 1):
                                        cell = customizable_reference_sheet[f"{col_letter}{row_num}"]
                                        cell.number_format = '#,##0'

                            # Apply formatting to 'Customizable Template' sheet
                            currency_columns_template = ['Baseline Spend', 'Baseline Price', 'Awarded Supplier Price', 'Awarded Supplier Spend', 'Savings']
                            number_columns_template = ['Bid Volume', 'Awarded Volume', 'Awarded Supplier Capacity']

                            for col_name in currency_columns_template:
                                col_letter = template_col_letter.get(col_name)
                                if col_letter:
                                    for row_num in range(2, max_row_template + 1):
                                        cell = customizable_template_sheet[f"{col_letter}{row_num}"]
                                        cell.number_format = '$#,##0.00'

                            for col_name in number_columns_template:
                                col_letter = template_col_letter.get(col_name)
                                if col_letter:
                                    for row_num in range(2, max_row_template + 1):
                                        cell = customizable_template_sheet[f"{col_letter}{row_num}"]
                                        cell.number_format = '#,##0'


                            # ======= Create the Scenario Selector sheet =======
                            # Continue with existing code for 'Scenario Selector' sheet
                            scenario_selector_df = customizable_df.copy()
                            grouping_column = st.session_state.grouping_column_raw 
                            # Remove 'Supplier Name' column from 'Scenario Selector' sheet
                            if 'Supplier Name' in scenario_selector_df.columns:
                                scenario_selector_df.drop(columns=['Supplier Name'], inplace=True)

                            # Ensure the grouping column is present
                            if not st.session_state.grouping_column_already_mapped:
                                if grouping_column not in scenario_selector_df.columns:
                                    # Add the grouping column next to 'Bid ID'
                                    bid_id_idx = scenario_selector_df.columns.get_loc('Bid ID')
                                    scenario_selector_df.insert(bid_id_idx + 1, grouping_column, '')

                            scenario_selector_df.to_excel(writer, sheet_name='Scenario Selector', index=False)
                            scenario_selector_sheet = workbook['Scenario Selector']

                            # Map column names to letters in 'Scenario Selector' sheet
                            scenario_selector_col_letter = {cell.value: get_column_letter(cell.column) for cell in scenario_selector_sheet[1]}
                            max_row_selector = scenario_selector_sheet.max_row
                            scenario_converter_sheet = workbook['Scenario Converter']
                            max_row_converter = scenario_converter_sheet.max_row

                            # Prepare formula components for 'Scenario Selector' sheet
                            scenario_converter_data_range = f"'Scenario Converter'!$B$2:$H${max_row_converter}"
                            scenario_converter_header_range = "'Scenario Converter'!$B$1:$H$1"
                            scenario_bid_ids_range = f"'Scenario Converter'!$A$2:$A${max_row_converter}"

                            bid_id_col_selector = scenario_selector_col_letter['Bid ID']
                            awarded_supplier_col = scenario_selector_col_letter['Awarded Supplier']
                            # Get the column letter for the grouping column in 'Scenario Selector' sheet
                            selector_grouping_col = scenario_selector_col_letter.get(grouping_column)

                            for row in range(2, max_row_selector + 1):
                                bid_id_cell = f"{bid_id_col_selector}{row}"
                                awarded_supplier_cell = f"{awarded_supplier_col}{row}"
                                bid_volume_cell = f"{scenario_selector_col_letter['Bid Volume']}{row}"
                                baseline_price_cell = f"{scenario_selector_col_letter['Baseline Price']}{row}"
                                grouping_cell = f"{selector_grouping_col}{row}" if selector_grouping_col else None

                                # Formula to pull 'Awarded Supplier' based on selected scenario
                                formula_awarded_supplier = (
                                    f"=IFERROR(INDEX({scenario_converter_data_range}, MATCH({bid_id_cell}, {scenario_bid_ids_range}, 0), "
                                    f"MATCH('Scenario Reports'!$A$1, {scenario_converter_header_range}, 0)), \"\")"
                                )
                                scenario_selector_sheet[awarded_supplier_cell].value = formula_awarded_supplier

                                # Use 'Awarded Supplier' directly in SUMIFS
                                awarded_supplier_price_cell = f"{scenario_selector_col_letter['Awarded Supplier Price']}{row}"
                                awarded_supplier_capacity_cell = f"{scenario_selector_col_letter['Awarded Supplier Capacity']}{row}"

                                # Awarded Supplier Price
                                formula_price = (
                                    f"=IFERROR(SUMIFS('{ref_sheet_name}'!{ref_bid_price_col}:{ref_bid_price_col}, "
                                    f"'{ref_sheet_name}'!{ref_bid_id_col}:{ref_bid_id_col}, {bid_id_cell}, "
                                    f"'{ref_sheet_name}'!{ref_supplier_name_col}:{ref_supplier_name_col}, {awarded_supplier_cell}), \"\")"
                                )
                                scenario_selector_sheet[awarded_supplier_price_cell].value = formula_price

                                # Awarded Supplier Capacity
                                formula_supplier_capacity = (
                                    f"=IFERROR(SUMIFS('{ref_sheet_name}'!{ref_supplier_capacity_col}:{ref_supplier_capacity_col}, "
                                    f"'{ref_sheet_name}'!{ref_bid_id_col}:{ref_bid_id_col}, {bid_id_cell}, "
                                    f"'{ref_sheet_name}'!{ref_supplier_name_col}:{ref_supplier_name_col}, {awarded_supplier_cell}), \"\")"
                                )
                                scenario_selector_sheet[awarded_supplier_capacity_cell].value = formula_supplier_capacity

                                # Awarded Volume
                                awarded_volume_cell = f"{scenario_selector_col_letter['Awarded Volume']}{row}"
                                formula_awarded_volume = f"=IF({bid_volume_cell}=\"\", \"\", MIN({bid_volume_cell}, {awarded_supplier_capacity_cell}))"
                                scenario_selector_sheet[awarded_volume_cell].value = formula_awarded_volume

                                # Awarded Supplier Spend
                                awarded_supplier_spend_cell = f"{scenario_selector_col_letter['Awarded Supplier Spend']}{row}"
                                formula_spend = f"=IF({awarded_supplier_price_cell}<>\"\", {awarded_supplier_price_cell}*{awarded_volume_cell}, \"\")"
                                scenario_selector_sheet[awarded_supplier_spend_cell].value = formula_spend

                                # Savings
                                savings_cell = f"{scenario_selector_col_letter['Savings']}{row}"
                                formula_savings = f"=IF({awarded_supplier_price_cell}<>\"\", ({baseline_price_cell}-{awarded_supplier_price_cell})*{awarded_volume_cell}, \"\")"
                                scenario_selector_sheet[savings_cell].value = formula_savings

                                # Baseline Spend
                                baseline_spend_cell = f"{scenario_selector_col_letter['Baseline Spend']}{row}"
                                formula_baseline_spend = f"={baseline_price_cell}*{bid_volume_cell}"
                                scenario_selector_sheet[baseline_spend_cell].value = formula_baseline_spend

                                # If grouping column exists, populate it
                                if grouping_cell and ref_grouping_col:
                                    col_index = column_index_from_string(ref_grouping_col) - column_index_from_string(ref_bid_id_col) + 1
                                    formula_grouping = (
                                        f"=IFERROR(VLOOKUP({bid_id_cell}, '{ref_sheet_name}'!{ref_bid_id_col}:{ref_grouping_col}, "
                                        f"{col_index}, FALSE), \"\")"
                                    )
                                    scenario_selector_sheet[grouping_cell].value = formula_grouping

                            # Apply formatting to 'Scenario Selector' sheet
                            currency_columns_selector = ['Baseline Spend', 'Baseline Price', 'Awarded Supplier Price', 'Awarded Supplier Spend', 'Savings']
                            number_columns_selector = ['Bid Volume', 'Awarded Volume', 'Awarded Supplier Capacity']

                            for col_name in currency_columns_selector:
                                col_letter = scenario_selector_col_letter.get(col_name)
                                if col_letter:
                                    for row_num in range(2, max_row_selector + 1):
                                        cell = scenario_selector_sheet[f"{col_letter}{row_num}"]
                                        cell.number_format = '$#,##0.00'

                            for col_name in number_columns_selector:
                                col_letter = scenario_selector_col_letter.get(col_name)
                                if col_letter:
                                    for row_num in range(2, max_row_selector + 1):
                                        cell = scenario_selector_sheet[f"{col_letter}{row_num}"]
                                        cell.number_format = '#,##0'


                                        # ======= Create the Scenario Reports sheet =======
                                        # Check if 'Scenario Reports' sheet exists; if not, create it
                                        if 'Scenario Reports' not in workbook.sheetnames:
                                            scenario_reports_sheet = workbook.create_sheet('Scenario Reports')
                                            logger.info("'Scenario Reports' sheet created.")
                                        else:
                                            scenario_reports_sheet = workbook['Scenario Reports']
                                            logger.info("'Scenario Reports' sheet already exists.")

                                        # Define the starting row for group summaries
                                        starting_row = 4  # As per user instruction, first summary starts at A4

                                        # Determine the number of unique suppliers (can be dynamic)
                                        unique_suppliers = st.session_state.merged_data[st.session_state.column_mapping['Supplier Name']].unique()
                                        num_suppliers = len(unique_suppliers)

                                        # Get the grouping column name
                                        grouping_column = st.session_state.grouping_column_raw  # e.g., 'Plant'

                                        # ======= Create a Separate DataFrame for Scenario Reports =======
                                        # **Change:** Use 'customizable_reference_df' as the source instead of 'merged_data'
                                        if grouping_column not in customizable_reference_df.columns:
                                            st.error(f"Grouping column '{grouping_column}' not found in merged data.")
                                            logger.error(f"Grouping column '{grouping_column}' not found in merged data.")
                                        else:
                                            # Create scenario_reports_df with 'Bid ID', grouping column, and 'Supplier Name' from 'Customizable Reference'
                                            scenario_reports_df = customizable_reference_df[[st.session_state.column_mapping['Bid ID'], grouping_column, st.session_state.column_mapping['Supplier Name']]].copy()

                                            # Get unique grouping values
                                            unique_groups = scenario_reports_df[grouping_column].dropna().unique()

                                            # ======= Ensure 'Scenario Selector' Sheet Exists =======
                                            if 'Scenario Selector' not in workbook.sheetnames:
                                                # Create 'Scenario Selector' sheet with headers if it doesn't exist
                                                scenario_selector_sheet = workbook.create_sheet('Scenario Selector')
                                                headers = ['Bid ID', grouping_column, st.session_state.column_mapping['Supplier Name'],
                                                        'Awarded Supplier Price', 'Awarded Supplier Capacity',
                                                        'Bid Volume', 'Baseline Price', 'Savings']  # Add other necessary headers
                                                for col_num, header in enumerate(headers, start=1):
                                                    cell = scenario_selector_sheet.cell(row=1, column=col_num, value=header)
                                                    cell.font = Font(bold=True)
                                                logger.info("'Scenario Selector' sheet created with headers.")
                                            else:
                                                scenario_selector_sheet = workbook['Scenario Selector']
                                                scenario_selector_headers = [cell.value for cell in scenario_selector_sheet[1]]
                                                logger.info("'Scenario Selector' sheet already exists.")

                                            # ======= Detect Presence of Grouping Column in 'Scenario Selector' =======
                                            # Assuming 'grouping_column_mapped' is the name of the grouping column in 'Scenario Selector'
                                            if st.session_state.grouping_column_mapped in scenario_selector_headers:
                                                column_offset = 1  # Shift references by 1 to the right
                                                logger.info(f"Grouping column '{st.session_state.grouping_column_mapped}' detected in 'Scenario Selector'. Column references will be shifted by {column_offset}.")
                                            else:
                                                column_offset = 0  # No shift needed
                                                logger.info("No grouping column detected in 'Scenario Selector'. Column references remain unchanged.")


                                            # Helper function to shift column letters based on header row count
                                            def shift_column(col_letter, scenario_selector_sheet, header_row=1):
                                                """
                                                Shifts a column letter by 1 if there are 13 column headers in the specified header row
                                                of 'Scenario Selector' sheet, does not shift if there are 12 column headers.
                                                
                                                Parameters:
                                                - col_letter (str): The original column letter (e.g., 'B').
                                                - scenario_selector_sheet (Worksheet): The 'Scenario Selector' worksheet object.
                                                - header_row (int): The row number where headers are located. Default is 1.
                                                
                                                Returns:
                                                - str: The shifted column letter.
                                                """
                                                # Extract the header row
                                                header_cells = scenario_selector_sheet[header_row]
                                                
                                                # Count non-empty header cells
                                                header_col_count = sum(1 for cell in header_cells if cell.value is not None)
                                                
                                                if header_col_count == 13:
                                                    offset = 1
                                                    logger.info(f"Detected {header_col_count} column headers in 'Scenario Selector' sheet. Shifting columns by 1.")
                                                elif header_col_count == 12:
                                                    offset = 0
                                                    logger.info(f"Detected {header_col_count} column headers in 'Scenario Selector' sheet. No shift applied.")
                                                else:
                                                    offset = 0  # Default shift
                                                    logger.warning(f"Unexpected number of column headers ({header_col_count}) in 'Scenario Selector' sheet. No shift applied.")
                                                
                                                # Convert column letter to index
                                                try:
                                                    col_index = column_index_from_string(col_letter)
                                                except ValueError:
                                                    logger.error(f"Invalid column letter provided: {col_letter}")
                                                    raise ValueError(f"Invalid column letter: {col_letter}")
                                                
                                                # Calculate new column index
                                                new_col_index = col_index + offset
                                                
                                                # Ensure the new column index is within Excel's limits (1 to 16384)
                                                if not 1 <= new_col_index <= 16384:
                                                    logger.error(f"Shifted column index {new_col_index} out of Excel's column range.")
                                                    raise ValueError(f"Shifted column index {new_col_index} out of Excel's column range.")
                                                
                                                # Convert back to column letter
                                                new_col_letter = get_column_letter(new_col_index)
                                                
                                                logger.debug(f"Column '{col_letter}' shifted by {offset} to '{new_col_letter}'.")
                                                
                                                return new_col_letter



                                            for group in unique_groups:
                                                # Insert Group Label
                                                scenario_reports_sheet[f"A{starting_row}"] = group
                                                scenario_reports_sheet[f"A{starting_row}"].font = Font(bold=True)

                                                # Insert Header Row
                                                headers = [
                                                    'Supplier Name',
                                                    'Awarded Volume',
                                                    '% of Business',
                                                    'Baseline Avg',
                                                    'Avg Bid Price',
                                                    '%Δ b/w Baseline and Avg Bid',
                                                    'RFP Savings'
                                                ]
                                                for col_num, header in enumerate(headers, start=1):
                                                    cell = scenario_reports_sheet.cell(row=starting_row + 1, column=col_num, value=header)
                                                    cell.font = Font(bold=True)

                                                # ======= Insert Formula for Supplier Name in the First Supplier Row =======
                                                supplier_name_cell = f"A{starting_row + 2}"  # Assigning to column A
                                                # Dynamic reference to the group label cell within the 'Scenario Reports' sheet
                                                group_label_cell = f"'Scenario Reports'!$A${starting_row}"

                                                try:
                                                    # Original column references in 'Scenario Selector' sheet
                                                    original_supplier_name_col = 'G'
                                                    original_group_col = 'B'

                                                    # Adjust column references based on column_offset
                                                    adjusted_supplier_name_col = shift_column(original_supplier_name_col, scenario_selector_sheet)
                                                    adjusted_group_col = shift_column(original_group_col, scenario_selector_sheet)

                                                    # Construct the formula with IFERROR to handle potential errors, without the leading '='
                                                    formula_supplier_name = (
                                                        f"IFERROR(UNIQUE(FILTER('Scenario Selector'!{adjusted_supplier_name_col}:{adjusted_supplier_name_col}, "
                                                        f"('Scenario Selector'!{adjusted_supplier_name_col}:{adjusted_supplier_name_col}<>\"\") * ('Scenario Selector'!{original_group_col}:{original_group_col}={group_label_cell}))), \"\")"
                                                    )

                                                    # Optional: Log the formula for debugging purposes
                                                    logger.info(f"Assigning formula to {supplier_name_cell}: {formula_supplier_name}")

                                                    # Assign the formula as text to the specified cell in the 'Scenario Reports' sheet
                                                    scenario_reports_sheet[supplier_name_cell].value = formula_supplier_name
                                                    logger.info(f"Successfully assigned formula as text to {supplier_name_cell}")
                                                except Exception as e:
                                                    logger.error(f"Failed to assign formula to {supplier_name_cell}: {e}")
                                                    st.error(f"An error occurred while assigning formulas to {supplier_name_cell}: {e}")

                                                # ======= Insert Formulas for Other Columns in the First Supplier Row =======
                                                awarded_volume_cell = f"B{starting_row + 2}"
                                                percent_business_cell = f"C{starting_row + 2}"
                                                avg_baseline_price_cell = f"D{starting_row + 2}"
                                                avg_bid_price_cell = f"E{starting_row + 2}"
                                                percent_delta_cell = f"F{starting_row + 2}"
                                                rfp_savings_cell = f"G{starting_row + 2}"

                                                try:
                                                    # Original column references for other formulas
                                                    original_awarded_volume_col = 'I'
                                                    original_bid_volume_col = 'I'  # Assuming 'Bid Volume' is in column 'I'
                                                    original_bid_price_col = 'D'
                                                    original_avg_bid_price_col = 'H'
                                                    original_savings_col = 'L'

                                                    # Adjust column references based on column_offset
                                                    adjusted_awarded_volume_col = shift_column(original_awarded_volume_col, scenario_selector_sheet)
                                                    adjusted_bid_volume_col = shift_column(original_bid_volume_col, scenario_selector_sheet)
                                                    adjusted_bid_price_col = shift_column(original_bid_price_col, scenario_selector_sheet)
                                                    adjusted_avg_bid_price_col = shift_column(original_avg_bid_price_col, scenario_selector_sheet)
                                                    adjusted_savings_col = shift_column(original_savings_col, scenario_selector_sheet)

                                                    # Awarded Volume Formula
                                                    formula_awarded_volume = (
                                                        f"=IF({supplier_name_cell}=\"\", \"\", SUMIFS('Scenario Selector'!${adjusted_awarded_volume_col}:${adjusted_awarded_volume_col}, "
                                                        f"'Scenario Selector'!${original_group_col}:${original_group_col}, {group_label_cell}, 'Scenario Selector'!${adjusted_supplier_name_col}:${adjusted_supplier_name_col}, {supplier_name_cell}))"
                                                    )
                                                    scenario_reports_sheet[awarded_volume_cell].value = formula_awarded_volume

                                                    # % of Business Formula
                                                    formula_percent_business = (
                                                        f"=IF({awarded_volume_cell}=0, \"\", {awarded_volume_cell}/SUMIFS('Scenario Selector'!${adjusted_awarded_volume_col}:${adjusted_awarded_volume_col}, "
                                                        f"'Scenario Selector'!${original_group_col}:${original_group_col}, {group_label_cell}))"
                                                    )
                                                    scenario_reports_sheet[percent_business_cell].value = formula_percent_business

                                                    # Avg Baseline Price Formula
                                                    formula_avg_baseline_price = (
                                                        f"=IF({supplier_name_cell}=\"\", \"\", AVERAGEIFS('Scenario Selector'!${adjusted_bid_price_col}:{adjusted_bid_price_col}, "
                                                        f"'Scenario Selector'!${original_group_col}:${original_group_col}, {group_label_cell}, 'Scenario Selector'!${adjusted_supplier_name_col}:${adjusted_supplier_name_col}, {supplier_name_cell}))"
                                                    )
                                                    scenario_reports_sheet[avg_baseline_price_cell].value = formula_avg_baseline_price

                                                    # Avg Bid Price Formula
                                                    formula_avg_bid_price = (
                                                        f"=IF({supplier_name_cell}=\"\", \"\", AVERAGEIFS('Scenario Selector'!${adjusted_avg_bid_price_col}:{adjusted_avg_bid_price_col}, "
                                                        f"'Scenario Selector'!${original_group_col}:${original_group_col}, {group_label_cell}, 'Scenario Selector'!${adjusted_supplier_name_col}:${adjusted_supplier_name_col}, {supplier_name_cell}))"
                                                    )
                                                    scenario_reports_sheet[avg_bid_price_cell].value = formula_avg_bid_price

                                                    # % Delta between Baseline and Avg Bid Formula
                                                    formula_percent_delta = (
                                                        f"=IF(AND({avg_baseline_price_cell}>0, {avg_bid_price_cell}>0), "
                                                        f"({avg_baseline_price_cell}-{avg_bid_price_cell})/{avg_baseline_price_cell}, \"\")"
                                                    )
                                                    scenario_reports_sheet[percent_delta_cell].value = formula_percent_delta

                                                    # RFP Savings Formula
                                                    formula_rfp_savings = (
                                                        f"=IFERROR(IF({supplier_name_cell}=\"\", \"\", AVERAGEIFS('Scenario Selector'!${adjusted_savings_col}:${adjusted_savings_col}, "
                                                        f"'Scenario Selector'!${original_group_col}:${original_group_col}, {group_label_cell}, 'Scenario Selector'!${adjusted_supplier_name_col}:${adjusted_supplier_name_col}, {supplier_name_cell})),\"\")"
                                                    )
                                                    scenario_reports_sheet[rfp_savings_cell].value = formula_rfp_savings

                                                    logger.info(f"Successfully assigned formulas to columns B-G in row {starting_row + 2}")
                                                except Exception as e:
                                                    logger.error(f"Failed to assign formulas to columns B-G in row {starting_row + 2}: {e}")
                                                    st.error(f"An error occurred while assigning formulas to columns B-G in row {starting_row + 2}: {e}")

                                                # ======= Advance the starting row =======
                                                # Since UNIQUE spills the results, we'll need to estimate the number of suppliers.
                                                # For simplicity, we'll add a fixed number of rows per group.
                                                # Adjust 'max_suppliers_per_group' as needed based on your data.
                                                max_suppliers_per_group = 10  # Example: Allow up to 10 suppliers per group
                                                starting_row += 2 + max_suppliers_per_group + 3  # Group label + header + supplier rows + spacing

                                                                            # ======= Add Drop-Down to 'Scenario Reports' Sheet =======
                                                # Restore the drop-down menu in cell A1
                                                try:
                                                    dv_scenario = DataValidation(type="list", formula1="'Scenario Converter'!$B$1:$H$1", allow_blank=True)
                                                    scenario_reports_sheet.add_data_validation(dv_scenario)
                                                    dv_scenario.add(scenario_reports_sheet['A1'])  # Placing drop-down in A1
                                                    logger.info("Scenario Reports sheet created with drop-down in cell A1.")
                                                except Exception as e:
                                                    st.error(f"Error adding drop-down to Scenario Reports sheet: {e}")
                                                    logger.error(f"Error adding drop-down to Scenario Reports sheet: {e}")

                                                logger.info(f"Advanced starting row to {starting_row}")

                            logger.info("Scenario Selector sheet created.")
                            logger.info("Customizable Analysis completed.")





                    # Convert the buffer to bytes
                    processed_data = output.getvalue()

                    # Provide a download button for the user
                    st.download_button(
                        label="Download Analysis Results",
                        data=processed_data,
                        file_name="scenario_analysis_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    logger.info("Analysis results prepared for download.")




    elif section == 'project_folder':
        if st.session_state.selected_project and st.session_state.selected_subfolder:
            st.title(f"{st.session_state.selected_project} - {st.session_state.selected_subfolder}")
            project_dir = BASE_PROJECTS_DIR / st.session_state.username / st.session_state.selected_project / st.session_state.selected_subfolder

            if not project_dir.exists():
                st.error("Selected folder does not exist.")
                logger.error(f"Folder {project_dir} does not exist.")
            else:
                # Path Navigation Buttons
                st.markdown("---")
                st.subheader("Navigation")
                path_components = [st.session_state.selected_project, st.session_state.selected_subfolder]
                path_keys = ['path_button_project', 'path_button_subfolder']
                path_buttons = st.columns(len(path_components))
                for i, (component, key) in enumerate(zip(path_components, path_keys)):
                    if i < len(path_components) - 1:
                        # Clickable button for higher-level folders
                        if path_buttons[i].button(component, key=key):
                            st.session_state.selected_subfolder = None
                    else:
                        # Disabled button for current folder
                        path_buttons[i].button(component, disabled=True, key=key+'_current')

                # List existing files with download buttons
                st.write(f"Contents of {st.session_state.selected_subfolder}:")
                files = [f.name for f in project_dir.iterdir() if f.is_file()]
                if files:
                    for file in files:
                        file_path = project_dir / file
                        # Provide a download link for each file
                        with open(file_path, "rb") as f:
                            data = f.read()
                        st.download_button(
                            label=f"Download {file}",
                            data=data,
                            file_name=file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.write("No files found in this folder.")

                st.markdown("---")
                st.subheader("Upload Files to This Folder")

                # File uploader for the selected subfolder
                uploaded_files = st.file_uploader(
                    "Upload Excel Files",
                    type=["xlsx"],
                    accept_multiple_files=True,
                    key='file_uploader_project_folder'
                )

                if uploaded_files:
                    for uploaded_file in uploaded_files:
                        if validate_uploaded_file(uploaded_file):
                            file_path = project_dir / uploaded_file.name
                            if file_path.exists():
                                overwrite = st.checkbox(f"Overwrite existing file '{uploaded_file.name}'?", key=f"overwrite_{uploaded_file.name}")
                                if not overwrite:
                                    st.warning(f"File '{uploaded_file.name}' not uploaded. It already exists.")
                                    logger.warning(f"User chose not to overwrite existing file '{uploaded_file.name}'.")
                                    continue
                            try:
                                with open(file_path, "wb") as f:
                                    f.write(uploaded_file.getbuffer())
                                st.success(f"File '{uploaded_file.name}' uploaded successfully.")
                                logger.info(f"File '{uploaded_file.name}' uploaded to '{project_dir}'.")
                            except Exception as e:
                                st.error(f"Failed to upload file '{uploaded_file.name}': {e}")
                                logger.error(f"Failed to upload file '{uploaded_file.name}': {e}")

        elif st.session_state.selected_project:
            # Display project button as disabled
            st.markdown("---")
            st.subheader("Navigation")
            project_button = st.button(st.session_state.selected_project, disabled=True, key='path_button_project_main')

            # List subfolders
            project_dir = BASE_PROJECTS_DIR / st.session_state.username / st.session_state.selected_project
            subfolders = ["Baseline", "Round 1 Analysis", "Round 2 Analysis", "Supplier Feedback", "Negotiations"]
            st.write("Subfolders:")
            for subfolder in subfolders:
                if st.button(subfolder, key=f"subfolder_{subfolder}"):
                    st.session_state.selected_subfolder = subfolder

        else:
            st.error("No project selected.")
            logger.error("No project selected.")

    elif section == 'settings':
        st.title('Settings')
        st.write("This section is under construction.")

    elif section == 'about':
        st.title('About')
        st.write("This section is under construction.")

    else:
        st.title('Home')
        st.write("This section is under construction.")

if __name__ == '__main__':
    main()

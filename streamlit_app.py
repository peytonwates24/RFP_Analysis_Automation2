import streamlit as st
import pandas as pd
from io import BytesIO
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Utility functions
def normalize_columns(df):
    """Map original columns to standard analysis columns."""
    column_mapping = {
        'bid_id': 'Bid ID',
        'business_group': 'Business Group',
        'product_type': 'Product Type',
        'incumbent': 'Incumbent',
        'baseline_price': 'Baseline Price',
        'bid_supplier_name': 'Awarded Supplier',
        'bid_supplier_capacity': 'Awarded Supplier Capacity',
        'bid_price': 'Awarded Supplier Price',
        'supplier_name': 'Awarded Supplier',
        'bid_volume': 'Bid Volume',
        'facility': 'Facility'
    }
    df = df.rename(columns=column_mapping)
    return df

def consolidate_columns(df):
    """Consolidate columns with duplicated names containing '_x' and '_y' suffixes."""
    cols_to_check = [col.rsplit('_', 1)[0] for col in df.columns if '_' in col and col.endswith(('_x', '_y'))]
    for col in set(cols_to_check):
        x_col = f"{col}_x"
        y_col = f"{col}_y"
        if x_col in df.columns and y_col in df.columns:
            if df[x_col].equals(df[y_col]):
                df[col] = df[x_col]
            else:
                df[col] = df[x_col].combine_first(df[y_col])
            df.drop([x_col, y_col], axis=1, inplace=True)
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

def load_and_combine_bid_data(file_path, supplier_name):
    """Load and combine bid data from relevant sheets."""
    try:
        sheet_names = pd.ExcelFile(file_path, engine='openpyxl').sheet_names
        relevant_sheets = [sheet for sheet in sheet_names if "bidsheet" in sheet.lower() or "bid sheet" in sheet.lower()]
        if len(sheet_names) == 1:
            relevant_sheets = sheet_names
        elif not relevant_sheets:
            st.error(f"No sheets named 'Bid Sheet' found. Available sheets: {', '.join(sheet_names)}.")
            return None
        data_frames = []
        for sheet in relevant_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
            df = normalize_columns(df)
            df['Awarded Supplier'] = supplier_name  # Set supplier name
            data_frames.append(df)
        combined_data = pd.concat(data_frames, ignore_index=True)
        return combined_data
    except Exception as e:
        st.error(f"An error occurred: {e}")
        logger.error(f"Error in load_and_combine_bid_data: {e}")
        return None

def load_baseline_data(file_path):
    """Load baseline data from the first sheet of the Excel file."""
    try:
        baseline_data = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')
        baseline_data = normalize_columns(baseline_data)
        return baseline_data
    except Exception as e:
        st.error(f"An error occurred: {e}")
        logger.error(f"Error in load_baseline_data: {e}")
        return None

def start_process(baseline_data, bid_files_suppliers):
    """Start the process of merging baseline data with bid data."""
    if baseline_data.empty or not bid_files_suppliers:
        st.error("Please select both the baseline and bid data files with supplier names.")
        return None
    all_merged_data = []
    for bid_file, supplier_name in bid_files_suppliers:
        combined_bid_data = load_and_combine_bid_data(bid_file, supplier_name)
        if combined_bid_data is None:
            st.error("Failed to load or combine bid data.")
            return None
        try:
            merged_data = pd.merge(baseline_data, combined_bid_data, on="Bid ID", how="left")
            merged_data = consolidate_columns(merged_data)
        except KeyError:
            st.error("'Bid ID' column not found in bid data or baseline data.")
            logger.error("KeyError: 'Bid ID' column not found during merge.")
            return None
        all_merged_data.append(merged_data)
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
    as_is_list = []
    bid_ids = data[bid_id_col].unique()
    for bid_id in bid_ids:
        bid_rows = data[(data[bid_id_col] == bid_id) & data[bid_price_col].notna()]
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
            logger.debug(f"No bid from incumbent for Bid ID {bid_id}.")
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

    bid_data = data.loc[data[bid_price_col].notna()]
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
            logger.debug(f"No bids for Bid ID {bid_id}.")
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

    bid_data = data.loc[data[bid_price_col].notna()]
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
            logger.debug(f"All suppliers excluded for Bid ID {bid_id}. Marked as Unallocated.")
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
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break
    best_of_best_excl_df = pd.DataFrame(best_of_best_excl_list)
    return best_of_best_excl_df

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
    
    bid_data = data.loc[data[bid_price_col].notna()]
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
                # Incumbent did not bid
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
                logger.debug(f"Incumbent did not bid for Bid ID {bid_id}. Entire volume is unallocated.")
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

# Streamlit App
def main():
    st.title("Scenario Analysis Tool")

    if 'merged_data' not in st.session_state:
        st.session_state.merged_data = None
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = {}
    if 'columns' not in st.session_state:
        st.session_state.columns = []

    # Upload baseline file
    baseline_file = st.file_uploader("Upload Baseline Sheet", type=["xlsx"])
    num_files = st.number_input("Number of Bid Sheets to Upload", min_value=1, step=1)

    bid_files_suppliers = []
    for i in range(num_files):
        bid_file = st.file_uploader(f"Upload Bid Sheet {i + 1}", type=["xlsx"], key=f'bid_file_{i}')
        supplier_name = st.text_input(f"Supplier Name for Bid Sheet {i + 1}", key=f'supplier_name_{i}')
        if bid_file and supplier_name:
            bid_files_suppliers.append((bid_file, supplier_name))

    # Merge Data
    if st.button("Merge Data"):
        if validate_uploaded_file(baseline_file) and bid_files_suppliers:
            baseline_data = load_baseline_data(baseline_file)
            if baseline_data is not None:
                merged_data = start_process(baseline_data, bid_files_suppliers)
                if merged_data is not None:
                    st.session_state.merged_data = merged_data
                    st.session_state.original_merged_data = merged_data.copy()
                    st.session_state.columns = list(merged_data.columns)
                    st.session_state.baseline_data = baseline_data
                    st.success("Data Merged Successfully. Please map the columns for analysis.")
                    logger.info("Data merged successfully.")

    if st.session_state.merged_data is not None:
        required_columns = ['Bid ID', 'Incumbent', 'Facility', 'Baseline Price', 'Bid Volume', 'Bid Price', 'Supplier Capacity', 'Supplier Name']

        # Ensure column_mapping persists
        if not st.session_state.column_mapping or set(st.session_state.column_mapping.keys()) != set(required_columns):
            st.session_state.column_mapping = auto_map_columns(st.session_state.merged_data, required_columns)

        st.write("Map the following columns:")
        for col in required_columns:
            st.session_state.column_mapping[col] = st.selectbox(f"Select Column for {col}", st.session_state.merged_data.columns, key=f"{col}_mapping")

        analyses_to_run = st.multiselect("Select Scenario Analyses to Run", [
            "As-Is",
            "Best of Best",
            "Best of Best Excluding Suppliers",
            "As-Is Excluding Suppliers"  # Added new analysis option
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

                if 'exclusions_bob' in st.session_state:
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

                if 'exclusions_ais' in st.session_state:
                    st.write("Current Exclusion Rules for As-Is Excluding Suppliers:")
                    for i, excl in enumerate(st.session_state.exclusions_ais):
                        st.write(f"{i + 1}. Supplier: {excl[0]}, Field: {excl[1]}, Logic: {excl[2]}, Value: {excl[3]}, Exclude All: {excl[4]}")

        if st.button("Run Analysis"):
            with st.spinner("Running analysis..."):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    baseline_data = st.session_state.baseline_data
                    original_merged_data = st.session_state.original_merged_data

                    if "As-Is" in analyses_to_run:
                        as_is_df = as_is_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                        as_is_df = add_missing_bid_ids(as_is_df, original_merged_data, st.session_state.column_mapping, 'As-Is')
                        as_is_df.to_excel(writer, sheet_name='As-Is', index=False)
                        logger.info("As-Is analysis completed.")

                    if "Best of Best" in analyses_to_run:
                        best_of_best_df = best_of_best_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                        best_of_best_df = add_missing_bid_ids(best_of_best_df, original_merged_data, st.session_state.column_mapping, 'Best of Best')
                        best_of_best_df.to_excel(writer, sheet_name='Best of Best', index=False)
                        logger.info("Best of Best analysis completed.")

                    if "Best of Best Excluding Suppliers" in analyses_to_run:
                        exclusions_list_bob = st.session_state.exclusions_bob if 'exclusions_bob' in st.session_state else []
                        best_of_best_excl_df = best_of_best_excluding_suppliers(st.session_state.merged_data, st.session_state.column_mapping, exclusions_list_bob)
                        best_of_best_excl_df = add_missing_bid_ids(best_of_best_excl_df, original_merged_data, st.session_state.column_mapping, 'BOB Excl Suppliers')
                        best_of_best_excl_df.to_excel(writer, sheet_name='BOB Excl Suppliers', index=False)
                        logger.info("Best of Best Excluding Suppliers analysis completed.")

                    if "As-Is Excluding Suppliers" in analyses_to_run:
                        exclusions_list_ais = st.session_state.exclusions_ais if 'exclusions_ais' in st.session_state else []
                        as_is_excl_df = as_is_excluding_suppliers_analysis(st.session_state.merged_data, st.session_state.column_mapping, exclusions_list_ais)
                        as_is_excl_df = add_missing_bid_ids(as_is_excl_df, original_merged_data, st.session_state.column_mapping, 'As-Is Excl Suppliers')
                        as_is_excl_df.to_excel(writer, sheet_name='As-Is Excl Suppliers', index=False)
                        logger.info("As-Is Excluding Suppliers analysis completed.")

                processed_data = output.getvalue()

            st.download_button(
                label="Download Analysis Results",
                data=processed_data,
                file_name="scenario_analysis_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            logger.info("Analysis results prepared for download.")

if __name__ == '__main__':
    main()


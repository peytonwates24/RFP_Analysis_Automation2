import pandas as pd
from .config import logger

# Scenario Analysis

def add_missing_bid_ids(best_of_best_excl_df, original_merged_data, column_mapping, sheet_name):
    """
    Add missing Bid IDs to the analysis results.

    Parameters:
        best_of_best_excl_df (pd.DataFrame): The current analysis DataFrame.
        original_merged_data (pd.DataFrame): The original merged DataFrame before exclusions.
        column_mapping (dict): A mapping of standard column names to actual DataFrame column names.
        sheet_name (str): The name of the sheet for logging or further processing.

    Returns:
        pd.DataFrame: Updated analysis DataFrame with missing Bid IDs handled.
    """
    # Extract Bid ID column name from column_mapping
    bid_id_col = column_mapping['Bid ID']
    
    # Identify missing Bid IDs
    missing_bid_ids = original_merged_data[bid_id_col].unique()
    processed_bid_ids = best_of_best_excl_df[bid_id_col].unique()
    bids_to_add = set(missing_bid_ids) - set(processed_bid_ids)
    
    for bid_id in bids_to_add:
        original_bid = original_merged_data[original_merged_data[bid_id_col] == bid_id].iloc[0]
        row_dict = {
            'Bid ID': bid_id,
            'Bid ID Split': 'A',
            'Facility': original_bid[column_mapping['Facility']],
            'Incumbent': original_bid[column_mapping['Incumbent']],
            'Baseline Price': original_bid[column_mapping['Baseline Price']],
            'Current Price': original_bid[column_mapping['Current Price']] if 'Current Price' in column_mapping else np.nan,
            'Bid Volume': original_bid[column_mapping['Bid Volume']],
            'Baseline Spend': original_bid['Baseline Spend'],  # Ensure 'Baseline Spend' is calculated beforehand
            'Awarded Supplier': 'Unallocated',
            'Awarded Supplier Price': np.nan,
            'Awarded Volume': np.nan,
            'Awarded Supplier Spend': np.nan,
            'Awarded Supplier Capacity': np.nan,
            'Baseline Savings': np.nan,
            'Current Price Savings': np.nan
        }
        best_of_best_excl_df = best_of_best_excl_df.append(row_dict, ignore_index=True)
        logger.debug(f"Added unallocated Bid ID {bid_id} to the analysis.")
    
    return best_of_best_excl_df

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

def as_is_analysis(data, column_mapping):
    """Perform 'As-Is' analysis with normalized fields, including Current Price Savings."""
    logger.info("Starting As-Is analysis.")
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']
    facility_col = column_mapping['Facility']

    # Check if 'Current Price' is mapped and not 'None'
    has_current_price = 'Current Price' in column_mapping and column_mapping['Current Price'] != 'None'
    if has_current_price:
        current_price_col = column_mapping['Current Price']
        data['Current Price'] = data[current_price_col]

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
            row_dict = {
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': bid_row[facility_col],
                'Incumbent': incumbent,
                'Baseline Price': bid_row[baseline_price_col],
                'Current Price': None if has_current_price else bid_row[baseline_price_col],  # Optional
                'Bid Volume': bid_row[bid_volume_col],
                'Baseline Spend': bid_row['Baseline Spend'],
                'Awarded Supplier': 'No Bid from Incumbent',
                'Awarded Supplier Price': None,
                'Awarded Volume': None,
                'Awarded Supplier Spend': None,
                'Awarded Supplier Capacity': None,
                'Baseline Savings': None,
                'Current Price Savings': None if has_current_price else None
            }
            as_is_list.append(row_dict)
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
            baseleline_as_is_savings = baseline_spend - as_is_spend


            row_dict = {
                'Bid ID': row[bid_id_col],
                'Bid ID Split': split_index,
                'Facility': row[facility_col],
                'Incumbent': incumbent,
                'Baseline Price': row[baseline_price_col],
                'Bid Volume': baseline_volume,
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': row[supplier_name_col],
                'Awarded Supplier Price': row[bid_price_col],
                'Awarded Volume': awarded_volume,
                'Awarded Supplier Spend': as_is_spend,
                'Awarded Supplier Capacity': supplier_capacity,
                'Baseline Savings': baseleline_as_is_savings
            }

            if has_current_price:
                current_price = data.loc[data[bid_id_col] == bid_id, 'Current Price'].iloc[0]
                row_dict['Current Price'] = current_price
                if row_dict['Awarded Supplier Price'] is not None and row_dict['Bid Volume'] is not None:
                    row_dict['Current Price Savings'] = (current_price - row_dict['Awarded Supplier Price']) * row_dict['Bid Volume']
                else:
                    row_dict['Current Price Savings'] = None

            as_is_list.append(row_dict)
            logger.debug(f"As-Is analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume}")
            remaining_volume -= awarded_volume
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break

    as_is_df = pd.DataFrame(as_is_list)

    # Define desired column order
    desired_columns = [
        'Bid ID', 'Bid ID Split', 'Facility', 'Incumbent',
        'Baseline Price'
    ]

    if has_current_price:
        desired_columns.append('Current Price')

    desired_columns.extend([
        'Bid Volume', 'Baseline Spend',
        'Awarded Supplier', 'Awarded Supplier Price',
        'Awarded Volume', 'Awarded Supplier Spend',
        'Awarded Supplier Capacity', 'Baseline Savings'
    ])

    if has_current_price:
        desired_columns.append('Current Price Savings')

    # Reorder columns
    as_is_df = as_is_df.reindex(columns=desired_columns)

    return as_is_df

def best_of_best_analysis(data, column_mapping):
    """Perform 'Best of Best' analysis with normalized fields, including Current Price Savings."""
    logger.info("Starting Best of Best analysis.")
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    facility_col = column_mapping['Facility']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']

    # Check if 'Current Price' is mapped and not 'None'
    has_current_price = 'Current Price' in column_mapping and column_mapping['Current Price'] != 'None'
    if has_current_price:
        current_price_col = column_mapping['Current Price']
        data['Current Price'] = data[current_price_col]

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
            row_dict = {
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': bid_row[facility_col],
                'Incumbent': bid_row[incumbent_col],
                'Baseline Price': bid_row[baseline_price_col],
                'Current Price': None if not has_current_price else bid_row[current_price_col],
                'Bid Volume': bid_row[bid_volume_col],
                'Baseline Spend': bid_row['Baseline Spend'],
                'Awarded Supplier': 'No Bids',
                'Awarded Supplier Price': None,
                'Awarded Volume': None,
                'Awarded Supplier Spend': None,
                'Awarded Supplier Capacity': None,
                'Baseline Savings': None,
                'Current Price Savings': None if not has_current_price else None
            }
            best_of_best_list.append(row_dict)
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
            baseline_savings = baseline_spend - best_of_best_spend

            row_dict = {
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
                'Baseline Savings': baseline_savings
            }

            if has_current_price:
                current_price = data.loc[data[bid_id_col] == bid_id, 'Current Price'].iloc[0]
                row_dict['Current Price'] = current_price
                if row_dict['Awarded Supplier Price'] is not None and row_dict['Bid Volume'] is not None:
                    row_dict['Current Price Savings'] = (current_price - row_dict['Awarded Supplier Price']) * row_dict['Bid Volume']
                else:
                    row_dict['Current Price Savings'] = None

            best_of_best_list.append(row_dict)
            logger.debug(f"Best of Best analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume}")
            remaining_volume -= awarded_volume
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break

    best_of_best_df = pd.DataFrame(best_of_best_list)

    # Define desired column order
    desired_columns = [
        'Bid ID', 'Bid ID Split', 'Facility', 'Incumbent',
        'Baseline Price'
    ]

    if has_current_price:
        desired_columns.append('Current Price')

    desired_columns.extend([
        'Bid Volume', 'Baseline Spend',
        'Awarded Supplier', 'Awarded Supplier Price',
        'Awarded Volume', 'Awarded Supplier Spend',
        'Awarded Supplier Capacity', 'Baseline Savings'
    ])

    if has_current_price:
        desired_columns.append('Current Price Savings')

    # Reorder columns
    best_of_best_df = best_of_best_df.reindex(columns=desired_columns)

    return best_of_best_df

def best_of_best_excluding_suppliers(data, column_mapping, excluded_conditions):
    """
    Perform 'Best of Best Excluding Suppliers' analysis, including Current Price and Savings.

    Parameters:
        data (pd.DataFrame): The input data containing bid information.
        column_mapping (dict): A mapping of standard column names to actual DataFrame column names.
        excluded_conditions (list of tuples): Conditions to exclude certain suppliers or bids.

    Returns:
        pd.DataFrame: A DataFrame containing the analysis results.

    Raises:
        ValueError: If required columns are missing in column_mapping or data.
    """
    logger.info("Starting Best of Best Excluding Suppliers analysis.")

    # Define required columns
    required_columns = [
        'Bid Price', 'Bid Volume', 'Baseline Price', 'Supplier Capacity',
        'Bid ID', 'Incumbent', 'Supplier Name', 'Facility'
    ]
    has_current_price = 'Current Price' in column_mapping and column_mapping['Current Price'] != 'None'
    if has_current_price:
        required_columns.append('Current Price')

    # Validate column_mapping
    missing_columns = [col for col in required_columns if col not in column_mapping]
    if missing_columns:
        logger.error(f"Missing columns in column_mapping: {missing_columns}")
        raise ValueError(f"Missing columns in column_mapping: {missing_columns}")

    # Validate data columns
    data_columns = [column_mapping[col] for col in required_columns]
    missing_data_columns = [col for col in data_columns if col not in data.columns]
    if missing_data_columns:
        logger.error(f"Missing columns in data: {missing_data_columns}")
        raise ValueError(f"Missing columns in data: {missing_data_columns}")

    # Extract column names from column_mapping
    bid_price_col = column_mapping['Bid Price']
    bid_volume_col = column_mapping['Bid Volume']
    baseline_price_col = column_mapping['Baseline Price']
    supplier_capacity_col = column_mapping['Supplier Capacity']
    bid_id_col = column_mapping['Bid ID']
    incumbent_col = column_mapping['Incumbent']
    supplier_name_col = column_mapping['Supplier Name']
    facility_col = column_mapping['Facility']
    current_price_col = column_mapping.get('Current Price', None)

    # Convert 'Bid Price' to numeric by removing '$' and handling non-numeric entries
    data[bid_price_col] = data[bid_price_col].astype(str).str.replace('$', '').str.strip()
    data[bid_price_col] = pd.to_numeric(data[bid_price_col], errors='coerce')

    # Convert 'Current Price' to numeric if applicable
    if has_current_price:
        data['Current Price'] = data[current_price_col].astype(str).str.replace('$', '').str.strip()
        data['Current Price'] = pd.to_numeric(data['Current Price'], errors='coerce')

    # Standardize supplier and incumbent names and strip leading/trailing spaces
    data[supplier_name_col] = data[supplier_name_col].astype(str).str.title().str.strip()
    data[incumbent_col] = data[incumbent_col].astype(str).str.title().str.strip()
    data[facility_col] = data[facility_col].astype(str).str.strip()

    # Calculate Baseline Spend
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]

    # Determine Valid Bids
    data['Valid Bid'] = data[bid_price_col].notna() & (data[bid_price_col] != 0)

    # Apply Exclusion Rules (Case-Insensitive)
    for condition in excluded_conditions:
        if len(condition) != 5:
            logger.error(f"Invalid condition format: {condition}")
            raise ValueError(f"Each excluded condition must have 5 elements, got: {condition}")
        
        supplier, field, logic, value, exclude_all = condition

        # Standardize supplier name for case-insensitive comparison
        supplier_standard = supplier.strip().title()

        if exclude_all:
            # Exclude all bids from the supplier (case-insensitive)
            data = data[data[supplier_name_col].str.lower() != supplier_standard.lower()]
            logger.debug(f"Excluding all bids from supplier '{supplier_standard}'.")
        else:
            # Ensure that the field exists in the data
            if field not in data.columns:
                logger.warning(f"Field '{field}' does not exist in data. Skipping this condition.")
                continue

            # If the field is of string type, standardize it for consistent comparison
            if pd.api.types.is_string_dtype(data[field]):
                data[field] = data[field].astype(str).str.title().str.strip()

            if logic == "Equal to":
                # Exclude bids where supplier matches and field equals the value (case-insensitive for supplier)
                condition_mask = (data[supplier_name_col].str.lower() == supplier_standard.lower()) & (data[field] == value)
                data = data[~condition_mask]
                logger.debug(f"Excluding bids from supplier '{supplier_standard}' where '{field}' == '{value}'.")
            elif logic == "Not equal to":
                # Exclude bids where supplier matches and field does not equal the value
                condition_mask = (data[supplier_name_col].str.lower() == supplier_standard.lower()) & (data[field] != value)
                data = data[~condition_mask]
                logger.debug(f"Excluding bids from supplier '{supplier_standard}' where '{field}' != '{value}'.")
            else:
                logger.warning(f"Unknown logic '{logic}' in condition: {condition}")

    # Filter valid bids after exclusions
    bid_data = data.loc[data['Valid Bid']].copy()
    bid_data = bid_data.sort_values([bid_id_col, bid_price_col])
    best_of_best_excl_list = []
    bid_ids = bid_data[bid_id_col].unique()

    for bid_id in bid_ids:
        bid_rows = bid_data[bid_data[bid_id_col] == bid_id]
        if bid_rows.empty:
            # Handle unallocated bids
            original_bid = data[data[bid_id_col] == bid_id].iloc[0]
            row_dict = {
                'Bid ID': bid_id,
                'Bid ID Split': 'A',
                'Facility': original_bid[facility_col],
                'Incumbent': original_bid[incumbent_col],
                'Baseline Price': original_bid[baseline_price_col],
                'Current Price': np.nan if not has_current_price else original_bid['Current Price'],
                'Bid Volume': original_bid[bid_volume_col],
                'Baseline Spend': original_bid['Baseline Spend'],
                'Awarded Supplier': 'Unallocated',
                'Awarded Supplier Price': np.nan,
                'Awarded Volume': np.nan,
                'Awarded Supplier Spend': np.nan,
                'Awarded Supplier Capacity': np.nan,
                'Baseline Savings': np.nan,
                'Current Price Savings': np.nan if not has_current_price else np.nan
            }
            best_of_best_excl_list.append(row_dict)
            logger.debug(f"No valid bids for Bid ID {bid_id}. Marked as Unallocated.")
            continue

        remaining_volume = bid_rows.iloc[0][bid_volume_col]
        split_index = 'A'

        for _, row in bid_rows.iterrows():
            supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else remaining_volume
            awarded_volume = min(remaining_volume, supplier_capacity)
            baseline_spend = awarded_volume * row[baseline_price_col]
            awarded_spend = awarded_volume * row[bid_price_col]
            baseline_savings = baseline_spend - awarded_spend

            # Calculate Current Price Savings if applicable
            if has_current_price:
                current_price = row['Current Price']
                current_price_savings = (current_price - row[bid_price_col]) * awarded_volume if pd.notna(current_price) else np.nan
            else:
                current_price_savings = np.nan

            row_dict = {
                'Bid ID': row[bid_id_col],
                'Bid ID Split': split_index,
                'Facility': row[facility_col],
                'Incumbent': row[incumbent_col],
                'Baseline Price': row[baseline_price_col],
                'Current Price': row['Current Price'] if has_current_price else np.nan,
                'Bid Volume': awarded_volume,
                'Baseline Spend': baseline_spend,
                'Awarded Supplier': row[supplier_name_col],
                'Awarded Supplier Price': row[bid_price_col],
                'Awarded Volume': awarded_volume,
                'Awarded Supplier Spend': awarded_spend,
                'Awarded Supplier Capacity': supplier_capacity,
                'Baseline Savings': baseline_savings,
                'Current Price Savings': current_price_savings
            }

            best_of_best_excl_list.append(row_dict)
            logger.debug(
                f"Processed Bid ID {bid_id}, Split {split_index}: "
                f"Awarded Volume = {awarded_volume}"
            )
            remaining_volume -= awarded_volume
            if remaining_volume <= 0:
                break
            split_index = chr(ord(split_index) + 1)

    # Create DataFrame from the list of dictionaries
    best_of_best_excl_df = pd.DataFrame(best_of_best_excl_list)

    # Define desired column order
    desired_columns = [
        'Bid ID', 'Bid ID Split', 'Facility', 'Incumbent',
        'Baseline Price'
    ]

    if has_current_price:
        desired_columns.append('Current Price')

    desired_columns.extend([
        'Bid Volume', 'Baseline Spend',
        'Awarded Supplier', 'Awarded Supplier Price',
        'Awarded Volume', 'Awarded Supplier Spend',
        'Awarded Supplier Capacity', 'Baseline Savings'
    ])

    if has_current_price:
        desired_columns.append('Current Price Savings')

    # Reorder columns
    best_of_best_excl_df = best_of_best_excl_df.reindex(columns=desired_columns)

    logger.info("Completed Best of Best Excluding Suppliers analysis.")
    return best_of_best_excl_df

def as_is_excluding_suppliers_analysis(data, column_mapping, excluded_conditions):
    """Perform 'As-Is Excluding Suppliers' analysis with exclusion rules, including Current Price Savings."""
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
    
    # Check if 'Current Price' is mapped and not 'None'
    has_current_price = 'Current Price' in column_mapping and column_mapping['Current Price'] != 'None'
    if has_current_price:
        current_price_col = column_mapping['Current Price']
        data['Current Price'] = data[current_price_col]
    
    # Standardize supplier and incumbent names
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
            if supplier == incumbent:
                if exclude_all:
                    incumbent_excluded = True
                    break
                elif logic == "Equal to" and all_rows[field].iloc[0] == value:
                    incumbent_excluded = True
                    break
                elif logic == "Not equal to" and all_rows[field].iloc[0] != value:
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
                baseline_savings = baseline_spend - awarded_spend

                row_dict = {
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
                    'Baseline Savings': baseline_savings  # Renamed from 'Savings'
                }

                if has_current_price:
                    current_price = data.loc[data[bid_id_col] == bid_id, 'Current Price'].iloc[0]
                    row_dict['Current Price'] = current_price
                    if row_dict['Awarded Supplier Price'] is not None and row_dict['Bid Volume'] is not None:
                        row_dict['Current Price Savings'] = (current_price - row_dict['Awarded Supplier Price']) * row_dict['Bid Volume']
                    else:
                        row_dict['Current Price Savings'] = None

                as_is_excl_list.append(row_dict)
                logger.debug(f"As-Is Excl analysis for Bid ID {bid_id}: Awarded to incumbent.")

                remaining_volume = bid_volume - awarded_volume
                if remaining_volume > 0:
                    # Remaining volume is unallocated
                    row_dict_unallocated = {
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
                        'Baseline Savings': None  # Renamed from 'Savings'
                    }

                    if has_current_price:
                        row_dict_unallocated['Current Price'] = None
                        row_dict_unallocated['Current Price Savings'] = None

                    as_is_excl_list.append(row_dict_unallocated)
                    logger.debug(f"Remaining volume for Bid ID {bid_id} is unallocated after awarding to incumbent.")
            else:
                # Incumbent did not bid or bid is invalid
                row_dict = {
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
                    'Baseline Savings': None,  # Renamed from 'Savings'
                }

                if has_current_price:
                    row_dict['Current Price'] = None
                    row_dict['Current Price Savings'] = None

                as_is_excl_list.append(row_dict)
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
                row_dict = {
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
                    'Baseline Savings': None  # Renamed from 'Savings'
                }

                if has_current_price:
                    row_dict['Current Price'] = current_price
                    row_dict['Current Price Savings'] = None

                as_is_excl_list.append(row_dict)
                logger.debug(f"No valid bids for Bid ID {bid_id} after exclusions. Entire volume is unallocated.")
                continue

            for _, row in valid_bids.iterrows():
                supplier_capacity = row[supplier_capacity_col] if pd.notna(row[supplier_capacity_col]) else remaining_volume
                awarded_volume = min(remaining_volume, supplier_capacity)
                awarded_spend = awarded_volume * row[bid_price_col]
                baseline_spend_allocated = awarded_volume * baseline_price
                baseline_savings = baseline_spend_allocated - awarded_spend

                row_dict = {
                    'Bid ID': row[bid_id_col],
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
                    'Baseline Savings': baseline_savings  # Renamed from 'Savings'
                }

                if has_current_price:
                    current_price = data.loc[data[bid_id_col] == bid_id, 'Current Price'].iloc[0]
                    row_dict['Current Price'] = current_price
                    if row_dict['Awarded Supplier Price'] is not None and row_dict['Bid Volume'] is not None:
                        row_dict['Current Price Savings'] = (current_price - row_dict['Awarded Supplier Price']) * row_dict['Bid Volume']
                    else:
                        row_dict['Current Price Savings'] = None

                as_is_excl_list.append(row_dict)
                logger.debug(f"As-Is Excl analysis for Bid ID {bid_id}, Split {split_index}: Awarded Volume = {awarded_volume} to {row[supplier_name_col]}")

                remaining_volume -= awarded_volume
                if remaining_volume <= 0:
                    break
                split_index = chr(ord(split_index) + 1)

            if remaining_volume > 0:
                # Remaining volume is unallocated
                row_dict_unallocated = {
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
                    'Baseline Savings': None  # Renamed from 'Savings'
                }

                if has_current_price:
                    row_dict_unallocated['Current Price'] = current_price
                    row_dict_unallocated['Current Price Savings'] = None

                as_is_excl_list.append(row_dict_unallocated)
                logger.debug(f"Remaining volume for Bid ID {bid_id} is unallocated after allocating to suppliers.")

    as_is_excl_df = pd.DataFrame(as_is_excl_list)

    # Define desired column order
    desired_columns = [
        'Bid ID', 'Bid ID Split', 'Facility', 'Incumbent',
        'Baseline Price'
    ]

    if has_current_price:
        desired_columns.append('Current Price')

    desired_columns.extend([
        'Bid Volume', 'Baseline Spend',
        'Awarded Supplier', 'Awarded Supplier Price',
        'Awarded Volume', 'Awarded Supplier Spend',
        'Awarded Supplier Capacity', 'Baseline Savings'
    ])

    if has_current_price:
        desired_columns.append('Current Price Savings')

    # Reorder columns
    as_is_excl_df = as_is_excl_df.reindex(columns=desired_columns)

    return as_is_excl_df

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

# Bid Coverage Reporting

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


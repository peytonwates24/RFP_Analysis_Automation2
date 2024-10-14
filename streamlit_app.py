import pandas as pd # comment for git
import streamlit as st
from io import BytesIO

# Utility functions
def normalize_columns(df):
    """Normalize column names to lower case and replace spaces with underscores."""
    df.columns = df.columns.map(str).str.strip().str.lower().str.replace(' ', '_')
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
            df['supplier_name'] = supplier_name
            data_frames.append(df)
        combined_data = pd.concat(data_frames, ignore_index=True)
        return combined_data
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

def load_baseline_data(file_path):
    """Load baseline data from the first sheet of the Excel file."""
    try:
        baseline_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        baseline_data = baseline_data[list(baseline_data.keys())[0]]
        baseline_data = normalize_columns(baseline_data)
        return baseline_data
    except Exception as e:
        st.error(f"An error occurred: {e}")
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
            merged_data = pd.merge(baseline_data, combined_bid_data, on="bid_id", how="left")
            merged_data = consolidate_columns(merged_data)
        except KeyError:
            st.error("'bid_id' column not found in bid data or baseline data.")
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
            column_mapping[col] = st.selectbox(f"Select Column for {col}", df.columns)
    return column_mapping

# Analysis functions
def as_is_analysis(data, column_mapping):
    """Perform 'As-Is' analysis with normalized fields."""
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
    bid_data = data.loc[data[bid_price_col].notna()]
    data['Baseline Spend'] = data[bid_volume_col] * data[baseline_price_col]
    as_is_list = []
    bid_ids = data[bid_id_col].unique()
    for bid_id in bid_ids:
        bid_rows = bid_data[bid_data[bid_id_col] == bid_id]
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
                'As-Is Supplier': 'No Bid from Incumbent',
                'As-Is Supplier Price': None,
                'As-Is Awarded Volume': None,
                'As-Is Supplier Spend': None,
                'As-Is Supplier Capacity': None,
                'As-Is Savings': None
            })
            continue
        remaining_volume = incumbent_bid.iloc[0][bid_volume_col]
        split_index = 'A'
        original_baseline_volume = remaining_volume
        for i, row in incumbent_bid.iterrows():
            supplier_capacity = row[supplier_capacity_col]
            awarded_volume = min(remaining_volume, supplier_capacity)
            if awarded_volume < original_baseline_volume:
                baseline_volume = awarded_volume
            else:
                baseline_volume = original_baseline_volume
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
                'As-Is Supplier': row[supplier_name_col],
                'As-Is Supplier Price': row[bid_price_col],
                'As-Is Awarded Volume': awarded_volume,
                'As-Is Supplier Spend': as_is_spend,
                'As-Is Supplier Capacity': supplier_capacity,
                'As-Is Savings': as_is_savings
            })
            remaining_volume -= awarded_volume
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break
    as_is_df = pd.DataFrame(as_is_list)
    return as_is_df

def as_is_excluding_suppliers_analysis(data, column_mapping, exclude_suppliers):
    """Perform 'As-Is' analysis excluding specific suppliers."""
    data = data[~data[column_mapping['Supplier Name']].isin(exclude_suppliers)]
    return as_is_analysis(data, column_mapping)

def best_of_best_analysis(data, column_mapping):
    """Perform 'Best of Best' analysis with normalized fields."""
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
                'Best of Best Supplier': 'No Bids',
                'Best of Best Supplier Price': None,
                'Best of Best Awarded Volume': None,
                'Best of Best Supplier Spend': None,
                'Best of Best Supplier Capacity': None,
                'Best of Best Savings': None
            })
            continue
        remaining_volume = bid_rows.iloc[0][bid_volume_col]
        split_index = 'A'
        original_baseline_volume = remaining_volume
        for i, row in bid_rows.iterrows():
            supplier_capacity = row[supplier_capacity_col]
            awarded_volume = min(remaining_volume, supplier_capacity)
            if awarded_volume < original_baseline_volume:
                baseline_volume = awarded_volume
            else:
                baseline_volume = original_baseline_volume
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
                'Best of Best Supplier': row[supplier_name_col],
                'Best of Best Supplier Price': row[bid_price_col],
                'Best of Best Awarded Volume': awarded_volume,
                'Best of Best Supplier Spend': best_of_best_spend,
                'Best of Best Supplier Capacity': supplier_capacity,
                'Best of Best Savings': best_of_best_savings
            })
            remaining_volume -= awarded_volume
            if remaining_volume > 0:
                split_index = chr(ord(split_index) + 1)
            else:
                break
    best_of_best_df = pd.DataFrame(best_of_best_list)
    return best_of_best_df


# Main app function
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
                    st.session_state.columns = list(merged_data.columns)
                    st.success("Data Merged Successfully. Please map the columns for analysis.")

    # Column Mapping
    if st.session_state.merged_data is not None:
        required_columns = ['Supplier Name', 'Incumbent', 'Facility', 'Baseline Price', 'Bid Volume', 'Bid Price', 'Supplier Capacity', 'Bid ID']
        st.session_state.column_mapping = auto_map_columns(st.session_state.merged_data, required_columns)

        analyses_to_run = st.multiselect("Select Scenario Analyses to Run", ["As-is", "Best of Best"])

        if st.button("Run Analysis"):
            with st.spinner("Running analysis..."):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    if "As-is" in analyses_to_run:
                        as_is_df = as_is_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                        as_is_df.to_excel(writer, sheet_name='As-Is')

                    if "Best of Best" in analyses_to_run:
                        best_of_best_df = best_of_best_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                        best_of_best_df.to_excel(writer, sheet_name='Best of Best')

                st.download_button(label="Download Analysis Results", data=output, file_name="scenario_analysis_results.xlsx")

if __name__ == "__main__":
    main()

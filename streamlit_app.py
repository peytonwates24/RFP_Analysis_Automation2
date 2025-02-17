import streamlit as st
import pandas as pd
import io
import os
from io import BytesIO
from modules.config import logger, BASE_PROJECTS_DIR, config
from modules.utils import validate_uploaded_file, normalize_columns, run_merge_warnings
from modules.authentication import authenticate_user
from modules.data_loader import load_baseline_data, start_process
from modules.analysis import *
from modules.presentations import *
from modules.projects import get_user_projects, create_project, delete_project
from openpyxl.utils import *
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font
from openpyxl import Workbook
from pptx import Presentation
import bcrypt
 
 
# from supabase import create_client, Client
 
# Initialize Supabase connection.
#@st.cache_resource
#def init_supabase_connection():
    #url = st.secrets["SUPABASE_URL"]
    #key = st.secrets["SUPABASE_KEY"]
   # return create_client(url, key)
 
#supabase = init_supabase_connection()
 
 
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
            <img src=https://scfuturemakers.com/wp-content/uploads/2017/11/Georgia-Pacific_overview_video-854x480-c-default.jpg alt="Logo">
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
    if st.sidebar.button('Dashboards'):
        navigate_to('dashboards')
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
 
 
                            # ---- Run the warning checks ----
                            run_merge_warnings(
                                st.session_state.baseline_data,
                                st.session_state.merged_data,
                                bid_files_suppliers
                            )
 
        else:
            # For Merged Data input method
            # === Merged Data Upload Section ===
            st.header("Upload Merged Data File")
            merged_file = st.file_uploader("Upload Merged Data File", type=["xlsx"], key="merged_data_file")
            merged_sheet = None
            if merged_file:
                try:
                    # Read the merged data file and extract sheet names.
                    excel_merged = pd.ExcelFile(merged_file, engine="openpyxl")
                    merged_sheet = st.selectbox(
                        "Select Sheet",
                        excel_merged.sheet_names,
                        key="merged_sheet_selection"
                    )
                except Exception as e:
                    st.error(f"Error reading merged data file: {e}")

            if merged_file and merged_sheet:
                try:
                    merged_data = pd.read_excel(merged_file, sheet_name=merged_sheet, engine="openpyxl")
                    # Optionally normalize columns using your existing function.
                    merged_data = normalize_columns(merged_data)
                    st.session_state.merged_data = merged_data
                    st.session_state.original_merged_data = merged_data.copy()
                    st.session_state.columns = list(merged_data.columns)
                    st.success("Merged data loaded successfully. Please map the columns for analysis.")

                    # --- Populate Rebate Data from the merged file ---
                    if "rebates" in excel_merged.sheet_names:
                        st.session_state["rebates_data"] = pd.read_excel(merged_file, sheet_name="rebates", engine="openpyxl")
                    else:
                        st.warning("No 'rebates' tab found in the merged data file; default rebate data will be used.")

                    # --- Populate Capacity Data from the merged file ---
                    # Look for a sheet named "capacity" (case-insensitive, trimmed)
                    capacity_sheet = None
                    for sheet in excel_merged.sheet_names:
                        if sheet.strip().lower() == "capacity":
                            capacity_sheet = sheet
                            break
                    if capacity_sheet is not None:
                        capacity_df = pd.read_excel(merged_file, sheet_name=capacity_sheet, engine="openpyxl")
                        if "Grouping" in capacity_df.columns:
                            capacity_df["Grouping"] = capacity_df["Grouping"].astype(str)
                        st.session_state["capacity_data"] = capacity_df
                    else:
                        st.warning("No 'capacity' tab found in the merged data file; default supplier capacity data will be used.")


                except Exception as e:
                    st.error(f"Error loading merged data: {e}")


 
        if st.session_state.merged_data is not None:
 
            st.subheader("Export Merged Data Before Mapping")
 
 
 
            # Let's define which columns are currency vs. generic numeric
            CURRENCY_COLUMNS = ['Baseline Price', 'Current Price', 'Bid Price'] 
            NUMBER_COLUMNS = ['Bid Volume', 'Supplier Capacity']                
 
            # 1) Convert columns to floats and remove floating artifacts
            #    (No forced rounding to fewer decimals, just remove binary noise)
            for col in CURRENCY_COLUMNS + NUMBER_COLUMNS:
                if col in st.session_state.merged_data.columns:
                    st.session_state.merged_data[col] = (
                        st.session_state.merged_data[col]
                        .astype(float)
                        .round(15)  # Enough precision to avoid artifacts like 0.14500000000000002
                    )
 
            # 2) Write the DataFrame to an in-memory Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.merged_data.to_excel(writer, index=False, sheet_name='MergedData')
 
                workbook = writer.book
                worksheet = writer.sheets['MergedData']
 
                # 3a) Apply a currency format to CURRENCY_COLUMNS
                for col_name in CURRENCY_COLUMNS:
                    if col_name in st.session_state.merged_data.columns:
                        col_idx = st.session_state.merged_data.columns.get_loc(col_name) + 1
                        # Rows start at 2 because row=1 is the header
                        for row in range(2, len(st.session_state.merged_data) + 2):
                            cell = worksheet.cell(row=row, column=col_idx)
                            cell.number_format = '$#,##0.00'  # 2 decimal places, comma separators, $ sign
 
                # 3b) Apply a plain numeric format (e.g., 'General') to NUMBER_COLUMNS
                for col_name in NUMBER_COLUMNS:
                    if col_name in st.session_state.merged_data.columns:
                        col_idx = st.session_state.merged_data.columns.get_loc(col_name) + 1
                        for row in range(2, len(st.session_state.merged_data) + 2):
                            cell = worksheet.cell(row=row, column=col_idx)
                            # "General" means Excel will show as a normal number without currency
                            cell.number_format = 'General'
 
            # 4) Retrieve the final Excel bytes
            excel_data = output.getvalue()
 
            # 5) Provide a download button in Streamlit
            st.download_button(
                label="Export Merged Data (Excel)",
                data=excel_data,
                file_name="merged_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
 
 
        # Proceed to Column Mapping if merged_data is available
        if st.session_state.merged_data is not None:
            required_columns = ['Bid ID', 'Incumbent', 'Facility', 'Baseline Price', 'Current Price',
                    'Bid Volume', 'Bid Price', 'Supplier Capacity', 'Supplier Name']
 
 
 
            # Initialize column_mapping if not already initialized
            if 'column_mapping' not in st.session_state:
                st.session_state.column_mapping = {}
 
            # Map columns
            st.write("Map the following columns:")
            for col in required_columns:
                if col == 'Current Price':
                    options = ['None'] + list(st.session_state.merged_data.columns)
                    st.session_state.column_mapping[col] = st.selectbox(
                        f"Select Column for {col}",
                        options,
                        key=f"{col}_mapping"
                    )
                else:
                    st.session_state.column_mapping[col] = st.selectbox(
                        f"Select Column for {col}",
                        st.session_state.merged_data.columns,
                        key=f"{col}_mapping"
                    )
 
            # Note
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
 
                # New Presentation Summaries multi-select
            presentation_options = ["Scenario Summary", "Bid Coverage Summary", "Supplier Comparison Summary"]
            selected_presentations = st.multiselect("Presentation Summaries", options=presentation_options)
            

            # === Conditions Section (Rebates Only) ===
            with st.expander("Conditions", expanded=False):
                st.markdown("### Conditions")
                st.write(
                    "Define optional conditions to be applied in the analysis. "
                    "Check a box to enable a condition. If data from the merged file exists, it will auto-populate."
                )

                # Retrieve supplier options from the mapped supplier column.
                if (
                    "merged_data" in st.session_state and 
                    "column_mapping" in st.session_state and 
                    st.session_state.column_mapping.get("Supplier Name")
                ):
                    mapped_supplier_col = st.session_state.column_mapping.get("Supplier Name")
                    supplier_options_raw = st.session_state.merged_data[mapped_supplier_col].dropna().unique().tolist()
                    supplier_options = [str(s) for s in supplier_options_raw]
                else:
                    supplier_options = []

                # ----- Rebate Condition (Fixed 7-Column Layout) -----
                if st.checkbox("Enable Rebate Condition", key="enable_rebate_condition"):
                    st.markdown("#### Rebate Information")
                    st.write("Review the rebate information. You can add rows as needed.")

                    # Define default data with 7 fixed columns:
                    # Supplier Name, Min Volume, Max Volume, Min Spend, Max Spend, % Rebate, $ Rebate.
                    default_rebate_data = pd.DataFrame({
                        "Supplier Name": [supplier_options[0] if supplier_options else ""],
                        "Min Volume": [0],
                        "Max Volume": [0],
                        "Min Spend": [0],
                        "Max Spend": [0],
                        "% Rebate": [0.0],
                        "$ Rebate": [0.0]
                    })

                    # Initialize the rebates_data in session state if not already present.
                    if "rebates_data" not in st.session_state:
                        st.session_state["rebates_data"] = default_rebate_data.copy()

                    # Define column configuration for the data editor.
                    rebate_column_config = {
                        "Supplier Name": st.column_config.TextColumn("Supplier Name"),
                        "Min Volume": st.column_config.NumberColumn("Min Volume", min_value=0, step=1),
                        "Max Volume": st.column_config.NumberColumn("Max Volume", min_value=0, step=1),
                        "Min Spend": st.column_config.NumberColumn("Min Spend", min_value=0, step=1),
                        "Max Spend": st.column_config.NumberColumn("Max Spend", min_value=0, step=1),
                        "% Rebate": st.column_config.NumberColumn("% Rebate", min_value=0, format="%.2f%%"),
                        "$ Rebate": st.column_config.NumberColumn("$ Rebate", min_value=0, format="$%.2f")
                    }

                    # Render the data editor for rebates (as in your existing code)
                    rebate_df = st.data_editor(
                        st.session_state["rebates_data"],
                        column_config=rebate_column_config,
                        num_rows="dynamic",
                        key="rebate_editor"
                    )
                    st.session_state["rebates_data"] = rebate_df

                    # Add a Save button to lock in the rebate data
                    if st.button("Save Rebate Data", key="save_rebate_data"):
                        # Optionally, you can save the data to a file (e.g., CSV) or a database here.
                        # For now, we'll store it in session_state and display a summary.
                        saved_data = st.session_state["rebates_data"]
                        
                        # Display a success message and a summary per supplier
                        st.success("Rebate data saved successfully!")
                        
                        st.markdown("### Saved Rebate Data Summary")
                        if not saved_data.empty:
                            for index, row in saved_data.iterrows():
                                st.write(
                                    f"**Supplier:** {row['Supplier Name']} | "
                                    f"Min Volume: {row['Min Volume']}, Max Volume: {row['Max Volume']} | "
                                    f"Min Spend: {row['Min Spend']}, Max Spend: {row['Max Spend']} | "
                                    f"% Rebate: {row['% Rebate']} | "
                                    f"$ Rebate: {row['$ Rebate']}"
                                )
                        else:
                            st.write("No rebate data to display.")




 
            # ----- Constraints Section -----
            with st.expander("Constraints", expanded=False):
                st.markdown("### Analysis Constraints")
                st.write("Define optional constraints for the analysis below. Each constraint is inactive unless enabled.")

                # Retrieve Bid ID options from merged data.
                if "merged_data" in st.session_state and "Bid ID" in st.session_state.merged_data.columns:
                    bid_options = st.session_state.merged_data["Bid ID"].dropna().unique().tolist()
                else:
                    bid_options = []

                # Retrieve supplier options from the mapped "Supplier Name" column.
                if (
                    "merged_data" in st.session_state 
                    and "column_mapping" in st.session_state 
                    and st.session_state.column_mapping.get("Supplier Name")
                ):
                    mapped_supplier_col = st.session_state.column_mapping.get("Supplier Name")
                    supplier_options = st.session_state.merged_data[mapped_supplier_col].dropna().unique().tolist()
                else:
                    supplier_options = []

                # 1. Maximum Suppliers Constraint
                if st.checkbox("Enable maximum suppliers constraint", key="enable_max_suppliers"):
                    default_max = len(supplier_options) if supplier_options else 1
                    max_suppliers = st.number_input(
                        "Maximum number of suppliers awarded across all bid IDs",
                        min_value=1,
                        value=default_max,
                        step=1,
                        key="max_suppliers"
                    )

                # 2. Supplier Volume Allocation Constraint
                if st.checkbox("Enable supplier volume allocation constraint", key="enable_supplier_capacity"):
                    st.markdown("#### Supplier Volume Allocation Constraint")
                    st.write("Review and edit the supplier capacity allocations. The table is auto-populated from the 'capacity' sheet if available; otherwise, default values are used. None of the fields are mandatory.")
                    
                    # Retrieve supplier options:
                    # If a capacity sheet was loaded, use its 'Supplier Name' column.
                    if "capacity_data" in st.session_state:
                        supplier_options = st.session_state.capacity_data["Supplier Name"].dropna().unique().tolist()
                    else:
                        # Otherwise, use the mapped supplier names from the merged data.
                        if ("merged_data" in st.session_state and 
                            "column_mapping" in st.session_state and 
                            st.session_state.column_mapping.get("Supplier Name")):
                            mapped_supplier_col = st.session_state.column_mapping.get("Supplier Name")
                            supplier_options = st.session_state.merged_data[mapped_supplier_col].dropna().unique().tolist()
                        else:
                            supplier_options = []
                    
                    # Provide a selector for the grouping column from the merged data headers.
                    if st.session_state.get("merged_data") is not None:
                        selected_grouping_col = st.selectbox(
                            "Select a grouping column for capacity allocation",
                            options=st.session_state.merged_data.columns.tolist(),
                            key="capacity_grouping_selector"
                        )
                        # Pull unique values from the selected grouping column.
                        grouping_values = st.session_state.merged_data[selected_grouping_col].dropna().unique().tolist()
                    else:
                        grouping_values = []
 
                    # Define column configuration with 3 columns: Supplier Name, Capacity, and Grouping.
                    capacity_column_config = {
                        "Supplier Name": st.column_config.SelectboxColumn("Supplier Name", options=supplier_options),
                        "Capacity": st.column_config.NumberColumn("Capacity", min_value=0, step=1),
                        "Grouping": st.column_config.SelectboxColumn("Grouping", options=grouping_values, help="Select a subgrouping value")
                    }
                    
                    # Use capacity data from the capacity sheet if available; otherwise, create default data.
                    if "capacity_data" in st.session_state:
                        default_capacity_data = st.session_state["capacity_data"]
                    else:
                        default_capacity_data = pd.DataFrame({
                            "Supplier Name": [supplier_options[0] if supplier_options else ""],
                            "Capacity": [0],
                            "Grouping": [""]
                        })
                    
                    capacity_df = st.data_editor(
                        default_capacity_data,
                        column_config=capacity_column_config,
                        num_rows="dynamic",
                        key="capacity_editor"
                    )
                    st.session_state["capacity_data"] = capacity_df



                # 3. Maximum Transitions Constraint
                if st.checkbox("Enable maximum transitions constraint", key="enable_max_transitions"):
                    max_transitions = st.number_input(
                        "Maximum transitions from incumbent to new awarded supplier",
                        min_value=0,
                        value=0,
                        step=1,
                        key="max_transitions"
                    )

                # 4. Lock In Supplier Constraint
                if st.checkbox("Enable lock in supplier constraint", key="enable_lock_in"):
                    st.markdown("#### Lock In Supplier")
                    st.write("Specify the Bid IDs that must be awarded to a specific supplier. These Bid IDs will be forced to use the designated supplier in the analysis.")
                    lock_in_column_config = {
                        "Bid ID": st.column_config.SelectboxColumn("Bid ID", options=bid_options),
                        "Locked Supplier": st.column_config.SelectboxColumn("Locked Supplier", options=supplier_options)
                    }
                    default_lock_in = pd.DataFrame({
                        "Bid ID": [bid_options[0] if bid_options else ""],
                        "Locked Supplier": [supplier_options[0] if supplier_options else ""]
                    })
                    lock_in_df = st.data_editor(
                        default_lock_in,
                        column_config=lock_in_column_config,
                        num_rows="dynamic",
                        key="lock_in_editor"
                    )

                # 5. Exclude Suppliers Constraint
                if st.checkbox("Enable supplier exclusion constraint", key="enable_exclusions"):
                    st.markdown("#### Exclude Suppliers")
                    global_exclusions = st.multiselect(
                        "Exclude these suppliers from all bid IDs",
                        options=supplier_options,
                        key="global_exclusions"
                    )
                    st.write("Alternatively, specify per-Bid ID exclusions below:")
                    exclusions_column_config = {
                        "Bid ID": st.column_config.SelectboxColumn("Bid ID", options=bid_options),
                        "Excluded Supplier": st.column_config.SelectboxColumn("Excluded Supplier", options=supplier_options, help="Select a supplier to exclude")
                    }
                    default_exclusions = pd.DataFrame({
                        "Bid ID": [bid_options[0] if bid_options else ""],
                        "Excluded Supplier": [supplier_options[0] if supplier_options else ""]
                    })
                    exclusions_df = st.data_editor(
                        default_exclusions,
                        column_config=exclusions_column_config,
                        num_rows="dynamic",
                        key="exclusions_editor"
                    )

                # 6. Splitting Awarded Supplier Constraint
                allow_splitting = st.checkbox(
                    "Enable splitting of awarded supplier for a bid ID across multiple suppliers",
                    key="allow_splitting"
                )
                if allow_splitting:
                    st.write("Specify the split details. Additional columns for extra suppliers are provided as dropdowns.")
                    split_column_config = {
                        "Bid ID": st.column_config.SelectboxColumn("Bid ID", options=bid_options),
                        "Supplier 1": st.column_config.SelectboxColumn("Supplier 1", options=supplier_options),
                        "Supplier 2": st.column_config.SelectboxColumn("Supplier 2", options=supplier_options),
                        "Supplier 3": st.column_config.SelectboxColumn("Supplier 3", options=supplier_options),
                        "Supplier 4": st.column_config.SelectboxColumn("Supplier 4", options=supplier_options),
                        "Supplier 5": st.column_config.SelectboxColumn("Supplier 5", options=supplier_options)
                    }
                    default_split = pd.DataFrame({
                        "Bid ID": [bid_options[0] if bid_options else ""],
                        "Supplier 1": [supplier_options[0] if supplier_options else ""],
                        "Supplier 2": [supplier_options[0] if supplier_options else ""],
                        "Supplier 3": [supplier_options[0] if supplier_options else ""],
                        "Supplier 4": [supplier_options[0] if supplier_options else ""],
                        "Supplier 5": [supplier_options[0] if supplier_options else ""]
                    })
                    split_df = st.data_editor(
                        default_split,
                        column_config=split_column_config,
                        num_rows="dynamic",
                        key="split_editor"
                    )

            # === Scenario Optimizer Section ===
            with st.expander("Scenario Optimizer", expanded=False):
                if st.checkbox("Enable Scenario Optimizer", key="enable_scenario_optimizer"):
                    st.markdown("### Scenario Optimizer")
                    st.write("Configure scenario optimization settings. (This section is for interface purposes only.)")
                    
                    # --- Grouping Column Selector (outside the form) ---
                    if st.session_state.get("merged_data") is not None:
                        grouping_columns = st.session_state.merged_data.columns.tolist()
                        selected_grouping = st.selectbox(
                            "Select Grouping Column (this determines the available subgroup values)",
                            options=grouping_columns,
                            key="rule_grouping_outside"
                        )
                        # Immediately calculate unique values from the selected column.
                        grouping_scope_options = st.session_state.merged_data[selected_grouping].dropna().unique().tolist()
                    else:
                        selected_grouping = ""
                        grouping_scope_options = []
                    
                    # --- Rule Input Form ---
                    with st.form(key="scenario_optimizer_form"):
                        # Capacity Scope: Global or Per Item.
                        capacity_scope = st.radio(
                            "Select Capacity Scope",
                            options=["Global", "Per Item"],
                            key="capacity_scope"
                        )
                        
                        # Rule Type and Operator in two columns.
                        col1, col2 = st.columns(2)
                        with col1:
                            rule_type = st.selectbox(
                                "Rule Type",
                                options=["% of Volume Awarded", "# of Facilities Awarded"],
                                key="rule_type"
                            )
                        with col2:
                            operator = st.selectbox(
                                "Operator",
                                options=["At Most", "At least", "Equal to"],
                                key="rule_operator"
                            )
                        
                        # Manual rule input.
                        rule_input = st.number_input(
                            "Rule Input (%)",
                            min_value=0.0,
                            max_value=100.0,
                            value=0.0,
                            step=0.1,
                            key="rule_input"
                        )
                        
                        # Grouping Scope, now based on the grouping column selected outside.
                        selected_grouping_scope = st.selectbox(
                            "Select Grouping Scope",
                            options=grouping_scope_options,
                            key="rule_grouping_scope"
                        )
                        
                        # Supplier Scope: from the mapped Supplier Name column.
                        if (
                            "merged_data" in st.session_state and 
                            "column_mapping" in st.session_state and 
                            st.session_state.column_mapping.get("Supplier Name")
                        ):
                            mapped_supplier_col = st.session_state.column_mapping.get("Supplier Name")
                            supplier_scope_options = st.session_state.merged_data[mapped_supplier_col].dropna().unique().tolist()
                        else:
                            supplier_scope_options = []
                        selected_supplier_scope = st.selectbox(
                            "Select Supplier Scope",
                            options=supplier_scope_options,
                            key="rule_supplier_scope"
                        )
                        
                        submitted = st.form_submit_button("Add Rule")
                        if submitted:
                            new_rule = {
                                "Capacity Scope": capacity_scope,
                                "Rule Type": rule_type,
                                "Operator": operator,
                                "Rule Input (%)": rule_input,
                                "Grouping Column": st.session_state.get("rule_grouping_outside", ""),
                                "Grouping Scope": selected_grouping_scope,
                                "Supplier Scope": selected_supplier_scope
                            }
                            if "scenario_rules" not in st.session_state:
                                st.session_state["scenario_rules"] = []
                            st.session_state["scenario_rules"].append(new_rule)
                            st.success("Rule added.")
                            # Optionally, you can remove the rerun if not desired.
                            # st.experimental_rerun()
                    
                    # ---- Display Added Rules with Descriptions ----
                    st.markdown("#### Added Rules")
                    if "scenario_rules" in st.session_state and st.session_state["scenario_rules"]:
                        for idx, rule in enumerate(st.session_state["scenario_rules"]):
                            # Build a descriptive sentence based on the rule.
                            if rule.get("Capacity Scope", "").strip().lower() == "global":
                                scope_desc = "Globally"
                            else:
                                scope_desc = "Per item"
                            
                            if rule["Rule Type"] == "% of Volume Awarded":
                                description = (f"{scope_desc}, {rule['Operator']} {rule['Rule Input (%)']}% of Volume Awarded "
                                               f"for {rule['Grouping Scope']} is awarded to Supplier {rule['Supplier Scope']}.")
                            elif rule["Rule Type"] == "# of Facilities Awarded":
                                description = (f"{scope_desc}, {rule['Operator']} {rule['Rule Input (%)']} facilities in "
                                               f"{rule['Grouping Scope']} are awarded to Supplier {rule['Supplier Scope']}.")
                            else:
                                description = (f"{scope_desc}, {rule['Operator']} {rule['Rule Input (%)']}% rule on {rule['Rule Type']} "
                                               f"in {rule['Grouping Scope']} for Supplier {rule['Supplier Scope']}.")
                            
                            st.write(f"**Rule {idx+1}:** {description}")
                            if st.button(f"Remove Rule {idx+1}", key=f"remove_rule_{idx}"):
                                st.session_state["scenario_rules"].pop(idx)
                                st.success("Rule removed. Please refresh the page if necessary.")
                    else:
                        st.write("No rules added yet.")


            # Exclusion rules for Best of Best Excluding Suppliers
            if "Best of Best Excluding Suppliers" in analyses_to_run:
                with st.expander("Configure Exclusion Rules for Best of Best Excluding Suppliers"):
                    st.header("Exclusion Rules for Best of Best Excluding Suppliers")
 
                    # Select Supplier to Exclude
                    supplier_name = st.selectbox(
                        "Select Supplier to Exclude",
                        st.session_state.merged_data['Awarded Supplier'].dropna().unique(),
                        key="supplier_name_excl_bob"
                    )
 
                    # Select Field for Rule
                    field = st.selectbox(
                        "Select Field for Rule",
                        st.session_state.merged_data.columns.drop('Awarded Supplier'),  # Exclude 'Awarded Supplier' if necessary
                        key="field_excl_bob"
                    )
 
                    # Select Logic
                    logic = st.selectbox(
                        "Select Logic (Equal to or Not equal to)",
                        ["Equal to", "Not equal to"],
                        key="logic_excl_bob"
                    )
 
                    # Select Value based on chosen field
                    unique_values = st.session_state.merged_data[field].dropna().unique()
 
                    if unique_values.size > 0:
                        value = st.selectbox(
                            "Select Value",
                            unique_values,
                            key="value_excl_bob"
                        )
                    else:
                        value = st.selectbox(
                            "Select Value",
                            options=[0],  # You can change 0 to any default value you prefer
                            index=0,
                            key="value_excl_bob"
                        )
 
 
                    # Checkbox to exclude all Bid IDs from the supplier
                    exclude_all = st.checkbox("Exclude from all Bid IDs", key="exclude_all_excl_bob")
 
                    # Button to add exclusion rule
                    if st.button("Add Exclusion Rule", key="add_excl_bob"):
                        if 'exclusions_bob' not in st.session_state:
                            st.session_state.exclusions_bob = []
                        # Append the exclusion rule as a tuple
                        st.session_state.exclusions_bob.append((supplier_name, field, logic, value, exclude_all))
                        logger.debug(f"Added exclusion rule for BOB Excl Suppliers: {supplier_name}, {field}, {logic}, {value}, Exclude All: {exclude_all}")
                        st.success("Exclusion rule added successfully!")
 
                    # Button to clear all exclusion rules
                    if st.button("Clear Exclusion Rules", key="clear_excl_bob"):
                        st.session_state.exclusions_bob = []
                        logger.debug("Cleared all exclusion rules for BOB Excl Suppliers.")
                        st.success("All exclusion rules cleared.")
 
                    # Display current exclusion rules
                    if 'exclusions_bob' in st.session_state and st.session_state.exclusions_bob:
                        st.write("**Current Exclusion Rules for Best of Best Excluding Suppliers:**")
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
 
 
            if "Scenario Summary" in selected_presentations:
                with st.expander("Configure Grouping for Scenario Summary Slides"):
                    st.header("Scenario Summary Grouping")
 
                    grouping_options = st.session_state.merged_data.columns.tolist()
                    # This will store the selected grouping in st.session_state["scenario_detail_grouping"]
                    st.selectbox("Group by for Scenario Detail", grouping_options, key="scenario_detail_grouping")
 
                    # Just call st.toggle with a key. Do not assign to st.session_state again.
                    st.toggle("Include Sub-Scenario Summaries?", key="scenario_sub_summaries_on")
 
                    if st.session_state["scenario_sub_summaries_on"]:
                        scenario_summary_fields = st.session_state.merged_data.columns.tolist()
                        st.selectbox("Scenario Summaries Selections", scenario_summary_fields, key="scenario_summary_selections")
 
                        # st.pills with a key will automatically store the selected values in st.session_state["sub_summary_selections"]
                        st.pills(
                            "Select scenario sub-summaries",
                            st.session_state.merged_data[st.session_state.scenario_summary_selections].unique(),
                            selection_mode="multi",
                            key="sub_summary_selections"
                        )
 
 
 
            if "Bid Coverage Summary" in selected_presentations:
                with st.expander("Configure Grouping for Bid Coverage Slides"):
                    st.header("Bid Coverage Grouping")
              
                    # Select group by field
                    bid_coverage_slides_grouping = st.selectbox("Group by", st.session_state.merged_data.columns, key="bid_coverage_slides_grouping")
 
            if "Supplier Comparison Summary" in selected_presentations:
                with st.expander("Configure Grouping for Supplier Comparison Summary"):
                    st.header("Supplier Comparison Summary")
              
                    # Select group by field
                    supplier_comparison_summary_grouping = st.selectbox("Group by", st.session_state.merged_data.columns, key="supplier_comparison_summary_grouping")
 
            if st.button("Run Analysis"):
                with st.spinner("Running analysis..."):
                    # 1. Define required analyses for Scenario Summary with correct casing
                    REQUIRED_ANALYSES_FOR_SCENARIO_SUMMARY = [
                        "As-Is",
                        "Best of Best",
                        "Best of Best Excluding Suppliers",
                        "As-Is Excluding Suppliers"
                    ]
 
                    # 2. Initialize validation flag and message
                    is_valid_selection = True
                    validation_message = ""
 
                    # 3. Check if "Scenario Summary" is selected
                    if "Scenario Summary" in selected_presentations:
                        # Check if at least one required analysis is selected
                        if not any(analysis in analyses_to_run for analysis in REQUIRED_ANALYSES_FOR_SCENARIO_SUMMARY):
                            is_valid_selection = False
                            validation_message = (
                                "⚠️ **Error:** To generate the 'Scenario Summary' presentation, please select at least one of the following Excel analyses: "
                                + ", ".join(REQUIRED_ANALYSES_FOR_SCENARIO_SUMMARY) + "."
                            )
 
                    # 4. Display validation messages and stop execution if invalid
                    if not is_valid_selection:
                        st.error(validation_message)
                        st.stop()  # Prevent further execution
 
                    # 5. Informative message for Scenario Summary dependencies
                    if "Scenario Summary" in selected_presentations:
                        st.info(
                            "📌 The 'Scenario Summary' presentation requires at least one of the following Excel analyses to be selected: "
                            + ", ".join(REQUIRED_ANALYSES_FOR_SCENARIO_SUMMARY) + "."
                        )
 
                    # 6. Proceed with analysis and presentation generation
                    excel_output = BytesIO()
                    ppt_output = BytesIO()
                    ppt_data = None
                    prs = None
 
                    # Determine if we need to produce an Excel file at all:
                    # We need Excel if we are running any analyses or if Scenario Summary is required.
                    need_excel = bool(analyses_to_run) or ("Scenario Summary" in selected_presentations)
 
                    try:
                        if need_excel:
                            # Perform the Excel-based analyses only if required
                            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                                baseline_data = st.session_state.baseline_data
                                original_merged_data = st.session_state.original_merged_data
 
                                # --- As-Is Analysis ---
                                if "As-Is" in analyses_to_run:
                                    as_is_df = as_is_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                                    as_is_df = add_missing_bid_ids(as_is_df, original_merged_data, st.session_state.column_mapping, 'As-Is')
                                    as_is_df.to_excel(writer, sheet_name='#As-Is', index=False)
                                    if 'Baseline Savings' in as_is_df.columns:
                                        logger.info("[CHECKPOINT] As-Is 'Baseline Savings' sample: %s", as_is_df['Baseline Savings'].head(5).tolist())
                                    else:
                                        logger.warning("As-Is DF has no 'Baseline Savings' column. Columns = %s", as_is_df.columns.tolist())
                                    logger.info("As-Is analysis completed.")
 
                                # --- Best of Best Analysis ---
                                if "Best of Best" in analyses_to_run:
                                    best_of_best_df = best_of_best_analysis(st.session_state.merged_data, st.session_state.column_mapping)
                                    best_of_best_df = add_missing_bid_ids(best_of_best_df, original_merged_data, st.session_state.column_mapping, 'Best of Best')
                                    best_of_best_df.to_excel(writer, sheet_name='#Best of Best', index=False)
                                    if 'Baseline Savings' in best_of_best_df.columns:
                                         logger.info("[CHECKPOINT] Best-of-Best 'Baseline Savings' sample: %s", best_of_best_df['Baseline Savings'].head(5).tolist())
                                    logger.info("Best of Best analysis completed.")
 
                                # --- Best of Best Excluding Suppliers Analysis ---
                                if "Best of Best Excluding Suppliers" in analyses_to_run:
                                    # Retrieve exclusion rules from session state, or use an empty list
                                    exclusions_list_bob = st.session_state.exclusions_bob if 'exclusions_bob' in st.session_state else []
                                   
                                    # Ensure column names are stripped of leading/trailing spaces
                                    st.session_state.merged_data.columns = st.session_state.merged_data.columns.str.strip()
                                   
                                    # Call the updated best_of_best_excluding_suppliers function with column_mapping
                                    try:
                                        best_of_best_excl_df = best_of_best_excluding_suppliers(
                                            data=st.session_state.merged_data,
                                            column_mapping=st.session_state.column_mapping,
                                            excluded_conditions=exclusions_list_bob
                                        )
                                    except ValueError as ve:
                                        st.error(f"Error in Best of Best Excluding Suppliers Analysis: {ve}")
                                        logger.error(f"Best of Best Excluding Suppliers Analysis failed: {ve}")
                                    else:
                                        # Call add_missing_bid_ids with column_mapping
                                        try:
                                            best_of_best_excl_df = add_missing_bid_ids(
                                                best_of_best_excl_df,
                                                original_merged_data,
                                                st.session_state.column_mapping,
                                                'BOB Excl Suppliers'
                                            )
                                        except Exception as e:
                                            st.error(f"Error in adding missing Bid IDs: {e}")
                                            logger.error(f"Adding Missing Bid IDs failed: {e}")
                                        else:
                                            # Export the result to Excel
                                            try:
                                                best_of_best_excl_df.to_excel(writer, sheet_name='#BOB Excl Suppliers', index=False)
                                                logger.info("Best of Best Excluding Suppliers analysis completed successfully.")
                                                st.success("Best of Best Excluding Suppliers analysis completed and exported to Excel.")
                                            except Exception as e:
                                                st.error(f"Error exporting Best of Best Excluding Suppliers Analysis to Excel: {e}")
                                                logger.error(f"Exporting to Excel failed: {e}")
 
                                # --- As-Is Excluding Suppliers Analysis ---
                                if "As-Is Excluding Suppliers" in analyses_to_run:
                                    exclusions_list_ais = st.session_state.exclusions_ais if 'exclusions_ais' in st.session_state else []
                                    as_is_excl_df = as_is_excluding_suppliers_analysis(
                                        st.session_state.merged_data,
                                        st.session_state.column_mapping,
                                        exclusions_list_ais
                                    )
                                    as_is_excl_df = add_missing_bid_ids(
                                        as_is_excl_df,
                                        original_merged_data,
                                        st.session_state.column_mapping,
                                        'As-Is Excl Suppliers'
                                    )
                                    as_is_excl_df.to_excel(writer, sheet_name='#As-Is Excl Suppliers', index=False)
                                    logger.info("As-Is Excluding Suppliers analysis completed.")
 
                                # --- Bid Coverage Report Processing ---
                                if "Bid Coverage Report" in analyses_to_run:
                                    variations = st.session_state.bid_coverage_variations if 'bid_coverage_variations' in st.session_state else []
                                    group_by_field = st.session_state.bid_coverage_group_by if 'bid_coverage_group_by' in st.session_state else st.session_state.merged_data.columns[0]
                                    if variations:
                                        bid_coverage_reports = bid_coverage_report(
                                            st.session_state.merged_data,
                                            st.session_state.column_mapping,
                                            variations,
                                            group_by_field
                                        )
 
                                        # Initialize startrow for Supplier Coverage sheet
                                        supplier_coverage_startrow = 0
 
                                        for report_name, report_df in bid_coverage_reports.items():
                                            if "Supplier Coverage" in report_name:
                                                sheet_name = "Supplier Coverage"
                                                if sheet_name not in writer.sheets:
                                                    report_df.to_excel(writer, sheet_name=sheet_name, startrow=supplier_coverage_startrow, index=False)
                                                    supplier_coverage_startrow += len(report_df) + 2
                                                else:
                                                    worksheet = writer.sheets[sheet_name]
                                                    worksheet.cell(row=supplier_coverage_startrow + 1, column=1, value=report_name)
                                                    supplier_coverage_startrow += 1
                                                    report_df.to_excel(writer, sheet_name=sheet_name, startrow=supplier_coverage_startrow, index=False)
                                                    supplier_coverage_startrow += len(report_df) + 2
                                                logger.info(f"{report_name} added to sheet '{sheet_name}'.")
                                            else:
                                                sheet_name_clean = report_name.replace(" ", "_")
                                                if len(sheet_name_clean) > 31:
                                                    sheet_name_clean = sheet_name_clean[:31]
                                                report_df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                                                logger.info(f"{report_name} generated and added to Excel.")
                                    else:
                                        st.warning("No Bid Coverage Report variations selected.")
 
                                # --- Customizable Analysis Processing ---
                                if "Customizable Analysis" in analyses_to_run:
                                    from modules.analysis import run_customizable_analysis_processing  # or wherever you place it
 
                                    try:
                                        run_customizable_analysis_processing(
                                            st=st,
                                            writer=writer,
                                            customizable_analysis=customizable_analysis,
                                            logger=logger,
                                            pd=pd,
                                            workbook=writer.book,
                                            data=st.session_state.merged_data,
                                            column_mapping=st.session_state.column_mapping,
                                            grouping_column_raw=st.session_state.grouping_column_raw,
                                            grouping_column_mapped=st.session_state.grouping_column_mapped,
                                            grouping_column_already_mapped=st.session_state.grouping_column_already_mapped
                                        )
                                    except Exception as e:
                                        st.error(f"An error occurred while running Customizable Analysis: {e}")
                                        logger.error(f"Error in Customizable Analysis block: {e}")
 
 
                        if need_excel:
                            excel_output.seek(0)
                            excel_data = excel_output.getvalue()
                            st.session_state.excel_data = excel_data
                        else:
                            # No Excel data generated
                            st.session_state.excel_data = None
                            excel_data = None
 
                        # If we needed Excel data for Scenario Summary, we read it now:
                        scenario_sheets_loaded = False
                        if "Scenario Summary" in selected_presentations and need_excel:
                            try:
                                scenario_excel_file = pd.ExcelFile(BytesIO(excel_output.getvalue()))
                                scenario_sheet_names = [sheet_name for sheet_name in scenario_excel_file.sheet_names if sheet_name.startswith('#')]
                                scenario_dataframes = {}
                                for sheet_name in scenario_sheet_names:
                                    df = pd.read_excel(scenario_excel_file, sheet_name=sheet_name)
                                            # Log a sample of "Savings" to see if it changed after saving/reading.
                                    if 'Baseline Savings' in df.columns:
                                        logger.info("[CHECKPOINT] After re-reading '%s': 'Baseline Savings' sample: %s",
                                                    sheet_name, df['Baseline Savings'].head(5).tolist())
                                    else:
                                        logger.warning("Sheet '%s' has no 'Baseline Savings' column. Columns = %s", sheet_name, df.columns.tolist())
                                    scenario_dataframes[sheet_name] = df
                                scenario_sheets_loaded = True
                            except Exception as e:
                                st.error(f"Failed to read the generated Excel file for scenario summary: {e}")
                                logger.error(f"Failed to read the generated Excel file for scenario summary: {e}")
                                st.stop()
 
                        # Generate PowerPoint File if presentations are selected
                        if "Scenario Summary" in selected_presentations:
                            try:
                                # Construct the template file path
                                script_dir = os.path.dirname(os.path.abspath(__file__))
                                template_file_path = os.path.join(script_dir, 'Slide template.pptx')
 
                                # Read the Excel file into a pandas ExcelFile object
                                excel_file = pd.ExcelFile(BytesIO(excel_data))
 
                                # Use 'original_merged_data' from your existing variables
                                # Ensure that 'original_merged_data' is available in this scope
                                if 'original_merged_data' in globals() or 'original_merged_data' in locals():
                                    original_df = original_merged_data.copy()
                                else:
                                    st.error("The 'original_merged_data' DataFrame is not available.")
                                    original_df = None  # Set to None to handle the error later
 
                                # Retrieve the grouping field selected by the user
                                scenario_detail_grouping = st.session_state.get('scenario_detail_grouping', None)
                                scenario_sub_summaries_on = st.session_state.get("scenario_sub_summaries_on", False)
                                scenario_summary_selections = st.session_state.get("scenario_summary_selections", None)
                                sub_summaries_list = st.session_state.get("sub_summary_selections", [])
 
                                # Get list of sheet names starting with '#'
                                scenario_sheet_names = [sheet_name for sheet_name in excel_file.sheet_names if sheet_name.startswith('#')]
 
                                scenario_dataframes = {}
                                for sheet_name in scenario_sheet_names:
                                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
 
                                    # Ensure 'Bid ID' exists in df
                                    if 'Bid ID' not in df.columns:
                                        st.error(f"'Bid ID' is not present in the scenario data for '{sheet_name}'. Skipping this sheet.")
                                        continue  # Skip this scenario

                                    # Merge scenario_detail_grouping if required
                                    if scenario_detail_grouping and scenario_detail_grouping not in df.columns:
                                        # Attempt to pull grouping from original_df, but ensuring one row per Bid ID
                                        if original_df is not None and 'Bid ID' in original_df.columns:
                                            if scenario_detail_grouping in original_df.columns:
                                                # 1) Make sure both have str Bid ID
                                                df['Bid ID'] = df['Bid ID'].astype(str)
                                                original_df['Bid ID'] = original_df['Bid ID'].astype(str)
                                                # 2) Build a unique_map from original_df to pick the FIRST non-blank grouping row
                                                unique_map = (
                                                    original_df.copy()
                                                    .assign(
                                                        _blank=lambda x: (
                                                            x[scenario_detail_grouping].isna() 
                                                            | (x[scenario_detail_grouping] == "")
                                                        )
                                                    )
                                                    # Sort so that non-blanks (_blank=False) come first, blanks go last
                                                    .sort_values("_blank")
                                                    # Then group by Bid ID, keep only the first row
                                                    .groupby("Bid ID", as_index=False)
                                                    .first()
                                                )
                                                # 3) Merge that single row per Bid ID
                                                df = df.merge(unique_map[['Bid ID', scenario_detail_grouping]], on='Bid ID', how='left')
                                                if scenario_detail_grouping not in df.columns:
                                                    st.error(f"Failed to merge the grouping field '{scenario_detail_grouping}' into '{sheet_name}'. Skipping this scenario.")
                                                    continue
                                            else:
                                                st.warning(
                                                    f"The selected grouping field '{scenario_detail_grouping}' is not in 'original_merged_data'. "
                                                    "No detail slides will be created for this scenario."
                                                )
                                                # It's okay to proceed without details if user still wants scenario summary slides
                                        else:
                                            st.warning(
                                                "The 'original_merged_data' is not available or 'Bid ID' is missing in 'original_merged_data'. "
                                                f"Cannot merge grouping for '{sheet_name}'."
                                            )
                                            # We can still proceed without scenario details
 
                                    # Merge scenario_summary_selections if sub-summaries are on and the column not present
                                    if scenario_sub_summaries_on and scenario_summary_selections:
                                        if scenario_summary_selections not in df.columns:
                                            # Attempt to merge from original_df
                                            if original_df is not None and 'Bid ID' in original_df.columns:
                                                if scenario_summary_selections in original_df.columns:
                                                    df['Bid ID'] = df['Bid ID'].astype(str)
                                                    original_df['Bid ID'] = original_df['Bid ID'].astype(str)
                                                    df = df.merge(
                                                        original_df[['Bid ID', scenario_summary_selections]], 
                                                        on='Bid ID', 
                                                        how='left'
                                                    )
                                                    if scenario_summary_selections not in df.columns:
                                                        st.warning(
                                                            f"Failed to merge the sub-summary field '{scenario_summary_selections}' into '{sheet_name}'. "
                                                            "Sub-summaries may not be created."
                                                        )
                                                else:
                                                    st.warning(
                                                        f"The sub-summary field '{scenario_summary_selections}' is not in 'original_merged_data'. "
                                                        f"No sub-summaries for '{sheet_name}'."
                                                    )
                                            else:
                                                st.warning("The 'original_merged_data' is not available or 'Bid ID' missing for merging sub-summary selections.")

                                    scenario_dataframes[sheet_name] = df
 
                                if not scenario_dataframes:
                                    st.error("No valid scenario dataframes were created. Please check your data.")
                                    ppt_data = None  # Ensure ppt_data is set to None if generation fails
                                else:
                                    # Generate the presentation (this now includes sub-summaries if toggled on)

                                    for sn, df in scenario_dataframes.items():
                                        if 'Baseline Savings' in df.columns:
                                            logger.info("[CHECKPOINT] Pre-PPT '%s': 'Baseline Savings' sample: %s",
                                                        sn, df['Baseline Savings'].head(5).tolist())
                                        else:
                                            logger.warning("[CHECKPOINT] Pre-PPT '%s' has no 'Baseline Savings' col. Columns: %s",
                                                        sn, df.columns.tolist())

                                    prs = create_scenario_summary_presentation(scenario_dataframes, template_file_path)
 
                                    if not prs:
                                        st.error("Failed to generate Scenario Summary presentation.")
                                        ppt_data = None  # Ensure ppt_data is set to None if generation fails
                                    else:
                                        # Save the presentation to BytesIO
                                        prs.save(ppt_output)
                                        ppt_data = ppt_output.getvalue()
                            except Exception as e:
                                st.error(f"An error occurred while generating the presentation: {e}")
                                logger.error(f"Error generating presentation: {e}")
                                ppt_data = None  # Ensure ppt_data is set to None if generation fails
                        else:
                            ppt_data = None  # No presentations selected



                            # --- Supplier Comparison Summary Presentation ---
                        if "Supplier Comparison Summary" in selected_presentations:
                            try:
                                supplier_comparison_summary_grouping = st.session_state.get('supplier_comparison_summary_grouping', None)
                                if not supplier_comparison_summary_grouping:
                                    st.error("Please select a grouping field for the Supplier Comparison Summary.")
                                else:
                                    # Make a copy of merged_data for supplier comparison summary
                                    if 'merged_data' not in st.session_state:
                                        st.error("No merged data available for Supplier Comparison Summary.")
                                    else:
                                        sc_df = st.session_state.merged_data.copy()
                       
                                        # If no existing presentation (prs) has been created yet, create one
                                        if prs is None:
                                            script_dir = os.path.dirname(os.path.abspath(__file__))
                                            template_file_path = os.path.join(script_dir, 'Slide template.pptx')
                                            prs = Presentation(template_file_path)
                       
                                        # Add the Supplier Comparison Summary slide
                                        prs = create_supplier_comparison_summary_slide(prs, sc_df, supplier_comparison_summary_grouping)
                       
                                        # Re-save presentation with new slides
                                        ppt_output = BytesIO()
                                        prs.save(ppt_output)
                                        ppt_output.seek(0)
                                        ppt_data = ppt_output.getvalue()
                       
                            except Exception as e:
                                st.error(f"An error occurred while generating the Supplier Comparison Summary presentation: {e}")
                                logger.error(f"Error generating Supplier Comparison Summary presentation: {e}")
 
                        # --- Bid Coverage Summary Presentation ---
                        if "Bid Coverage Summary" in selected_presentations:
                            try:
                                bid_coverage_slides_grouping = st.session_state.get('bid_coverage_slides_grouping', None)
                                if not bid_coverage_slides_grouping:
                                    st.error("Please select a grouping field for the Bid Coverage slides.")
                                else:
                                    bc_df = st.session_state.merged_data.copy()
                                    if prs is None:
                                        script_dir = os.path.dirname(os.path.abspath(__file__))
                                        template_file_path = os.path.join(script_dir, 'Slide template.pptx')
                                        prs = Presentation(template_file_path)
 
                                    prs = create_bid_coverage_summary_slides(prs, bc_df, bid_coverage_slides_grouping)
                                    ppt_output = BytesIO()
                                    prs.save(ppt_output)
                                    ppt_output.seek(0)
                                    ppt_data = ppt_output.getvalue()
                            except Exception as e:
                                st.error(f"An error occurred while generating the Bid Coverage Summary presentation: {e}")
                                logger.error(f"Error generating Bid Coverage Summary presentation: {e}")
 
                    except Exception as e:
                        st.error(f"An error occurred during analysis: {e}")
                        logger.error(f"Error during analysis: {e}")
                        st.stop()
 
                    # If we never needed Excel, excel_data is None. Otherwise, we have excel_output.
                    if need_excel:
                        excel_output.seek(0)
                        excel_data = excel_output.getvalue()
                        st.session_state.excel_data = excel_data
                    else:
                        # No Excel data generated
                        st.session_state.excel_data = None
 
                    st.session_state.ppt_data = ppt_data
 
                    st.success("Analysis completed successfully. Please download your files below.")
 
                    # Excel download button (only if excel_data is generated)
                    if st.session_state.excel_data:
                        st.download_button(
                            label="Download Analysis Results (Excel)",
                            data=st.session_state.excel_data,
                            file_name="scenario_analysis_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        logger.info("Analysis results prepared for download.")
                    else:
                        # If user selected no excel analyses and no scenario summary, no excel is expected
                        logger.info("No Excel analysis results available for download.")
                   
                    # PowerPoint download button (only if data is available)
                    if st.session_state.ppt_data:
                        st.download_button(
                            label="Download Presentation (PowerPoint)",
                            data=st.session_state.ppt_data,
                            file_name="presentation_summary.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        logger.info("Presentation prepared for download.")
                    else:
                        # If no ppt_data is available, either user didn't select presentations
                        # or required conditions for certain presentations were not met.
                        logger.info("No presentation available for download.")
 
 
 
 
 
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
 
    elif section == 'dashboards':
        st.title('Dashboards')
#        # Insert file info into Supabase
# #        if uploaded_file is not None:
#             file_name = uploaded_file.name
#             file_size = len(uploaded_file.getvalue())  # size in bytes
 
#             try:
#                 # Optionally, upload to Supabase Storage
#                 upload_response = supabase.storage.from_("uploads").upload(file_name, uploaded_file.read())
#                 if upload_response.status_code == 200:
#                     st.success(f"File '{file_name}' uploaded to storage!")
#                     logger.info(f"File '{file_name}' uploaded to Supabase Storage.")
 
#                     # Insert metadata into 'uploaded_files' table
#                     db_response = supabase.table("uploaded_files").insert({
#                         "filename": file_name,
#                         "file_size": file_size
#                     }).execute()
 
#                     if db_response.status_code == 201:
#                         st.success(f"File '{file_name}' recorded in the database!")
#                         logger.info(f"File '{file_name}' recorded in Supabase database.")
#                     else:
#                         st.error(f"Failed to record '{file_name}' in the database. Status Code: {db_response.status_code}")
#                         logger.error(f"Failed to record '{file_name}' in Supabase database. Status Code: {db_response.status_code}")
#                 else:
#                     st.error(f"Failed to upload '{file_name}' to storage. Status Code: {upload_response.status_code}")
#                     logger.error(f"Failed to upload '{file_name}' to Supabase Storage. Status Code: {upload_response.status_code}")
#             except Exception as e:
#                 st.error(f"An error occurred while uploading the file: {e}")
#                 logger.error(f"Error uploading file '{file_name}' to Supabase: {e}")
 
#         # Fetch previously uploaded files
#         try:
#             response = supabase.table("uploaded_files").select("*").order("id", ascending=False).execute()
#             rows = response.data
#             if rows:
#                 for row in rows:
#                     # Generate a signed URL for file download
#                     signed_url = supabase.storage.from_("uploads").create_signed_url(row['filename'], 3600).data['signedURL']
#                     st.write(f"**Filename:** {row['filename']} | [Download]({signed_url}) | **Size:** {row['file_size']} bytes | **Uploaded:** {row['upload_time']}")
#             else:
#                 st.write("No files uploaded yet.")
#         except Exception as e:
#             st.error(f"Failed to fetch uploaded files: {e}")
#             logger.error(f"Error fetching uploaded files from Supabase: {e}")
 
 
    elif section == 'about':
        st.title("About")
 
        # Step 1: Read the entire markdown file
        with open("docs/report_documentation.md", "r", encoding="utf-8") as f:
            doc_text = f.read()
 
 
        # Step 2: Identify a consistent pattern in your markdown headings.
        # For example, if each report section starts with "##" followed by the report name,
        # we can split on that pattern or use a more robust parsing approach.
        #
        # In this example, let's assume your markdown is structured with distinct
        # second-level headings (##) for each report type, like:
        #
        # ## "As-Is" Report
        # (content...)
        #
        # ## "As-Is + Exclusions" Report
        # (content...)
        #
        # and so forth.
       
        # Split the text by '## ' to separate each section
        sections = doc_text.split('## ')
        # The first split might be text before the first "##", so we can ignore it if empty
        sections = [sec.strip() for sec in sections if sec.strip()]
 
        # Each element in 'sections' should now start with something like:
        # '"As-Is" Report\n\n**What it does:** ...'
        # We can map each section to a title and content by splitting on the first newline.
        report_dict = {}
        for sec in sections:
            # Split on the first newline to separate the title line from the content
            lines = sec.split('\n', 1)
            title_line = lines[0].strip()
            content = lines[1].strip() if len(lines) > 1 else ""
            report_dict[title_line] = content
 
        # Step 3: Create tabs. The keys in report_dict should match your reports.
        # Extract just the titles you're interested in, ensuring they match your markdown headings.
        report_titles = [
            'Reports Overview',
            'As-Is',
            'As-Is + Exclusions',
            'Best of Best',
            'Best of Best + Exclusions',
            'Customizable Analysis',
            'Bid Coverage Report'
        ]
       
        st.markdown(
            """
            <style>
            /* Allow the tabs container to scroll horizontally */
            div[data-baseweb="tab-list"] {
            overflow-x: auto;
            white-space: nowrap;
            }
            </style>
            """,
            unsafe_allow_html=True
        )
 
        tab_objects = st.tabs(report_titles)
 
        # Step 4: Display each report’s content in the corresponding tab dynamically
        for tab, title in zip(tab_objects, report_titles):
            with tab:
                # Safely access the dictionary (if there's a mismatch, provide a fallback)
                content = report_dict.get(title, "*Content not found.*")
                # Write the content as markdown
                st.markdown("## " + title)
                st.markdown(content)
 
    else:
        st.title('Home')
        st.write("This section is under construction.")
 
if __name__ == '__main__':
    main()                                      


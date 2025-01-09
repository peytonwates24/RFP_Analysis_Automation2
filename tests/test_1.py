# tests/test_1.py

import pytest
import pandas as pd
from io import BytesIO
from pathlib import Path
from openpyxl import Workbook

# Import necessary functions from your analysis module
# Adjust the import path according to your project structure
from modules.analysis import *
from modules.config import *
from modules.utils import *
from modules.data_loader import *
from modules.analysis import *
from modules.presentations import *
from openpyxl.utils import *

@pytest.mark.parametrize("upload_method", ["merged", "separate"])
def test_run_analysis(upload_method, get_test_data_dir, setup_logging):
    """
    Test the analysis pipeline with both merged data and separate bid & baseline files.
    """
    logger = setup_logging

    # Define paths based on upload method
    if upload_method == "merged":
        merged_file = get_test_data_dir / 'merged' / 'test_merged_data.xlsx'
        expected_excel_path = Path('tests/test_outputs/merged/expected_excel_output_merged.xlsx')
        expected_ppt = Path('tests/test_outputs/merged/expected_powerpoint_output_merged.pptx')
    elif upload_method == "separate":
        separate_data_dir = get_test_data_dir / 'separate'
        bid_files = [
            separate_data_dir / 'test_bid_file1.xlsx',
            separate_data_dir / 'test_bid_file2.xlsx',
            separate_data_dir / 'test_bid_file3.xlsx',
            separate_data_dir / 'test_bid_file4.xlsx',
            separate_data_dir / 'test_bid_file5.xlsx'
        ]
        baseline_file = separate_data_dir / 'test_baseline.xlsx'
        expected_excel_path = Path('tests/test_outputs/separate/expected_excel_output_separate.xlsx')
        expected_ppt = Path('tests/test_outputs/separate/expected_powerpoint_output_separate.pptx')
    else:
        pytest.fail(f"Unknown upload method: {upload_method}")

    # Check if the necessary files exist
    if upload_method == "merged":
        assert merged_file.exists(), f"Merged test file {merged_file} does not exist."
        logger.info(f"Found merged data file: {merged_file}")
    elif upload_method == "separate":
        for file in bid_files + [baseline_file]:
            assert file.exists(), f"Required test file {file} does not exist."
            logger.info(f"Found required file: {file}")

    # Load data based on upload method
    if upload_method == "merged":
        merged_df = pd.read_excel(merged_file, engine='openpyxl')
        logger.debug(f"Merged DataFrame columns before normalization: {merged_df.columns.tolist()}")
    elif upload_method == "separate":
        bid_dfs = [pd.read_excel(bid_file, engine='openpyxl') for bid_file in bid_files]
        merged_df = pd.concat(bid_dfs, ignore_index=True)
        baseline_df = pd.read_excel(baseline_file, engine='openpyxl')
        merged_df = pd.merge(
            baseline_df,
            merged_df,
            on='Bid ID',
            how='outer',
            suffixes=('_baseline', '_bid')
        )
        logger.debug(f"Merged DataFrame columns after merge: {merged_df.columns.tolist()}")

    # Combine suffixed columns for 'separate' upload method
    if upload_method == "separate":
        merged_df.loc[:, 'Facility'] = merged_df['Facility_baseline'].combine_first(merged_df['Facility_bid'])
        merged_df.loc[:, 'Bid Volume'] = merged_df['Bid Volume_baseline'].combine_first(merged_df['Bid Volume_bid'])
        # Drop the suffixed columns
        merged_df.drop(['Facility_baseline', 'Facility_bid', 'Bid Volume_baseline', 'Bid Volume_bid'], axis=1, inplace=True)
        logger.debug(f"Columns after combining suffixed columns: {merged_df.columns.tolist()}")

    # Apply column mapping
    merged_df = normalize_columns(merged_df)
    column_mapping = {
        'Bid ID': 'Bid ID',
        'Facility': 'Facility',
        'Incumbent': 'Incumbent',
        'Bid Volume': 'Bid Volume',
        'Baseline Price': 'Baseline Price',
        'Current Price': 'Current Price',
        'Bid Price': 'Bid Price',
        'Supplier Capacity': 'Bid Supplier Capacity',  # Changed key
        'Supplier Name': 'Bid Supplier Name'          # Changed key
    }

    # Debugging: Check columns after normalization
    logger.debug(f"Columns after normalization: {merged_df.columns.tolist()}")

    # Ensure all required columns are present
    required_columns = list(column_mapping.values())
    missing_columns = [col for col in required_columns if col not in merged_df.columns]
    assert not missing_columns, f"Missing columns in merged data: {missing_columns}"
    logger.info("All required columns are present in merged data.")

    # Add 'Awarded Supplier' automatically
    merged_df.loc[:, 'Awarded Supplier'] = merged_df['Bid Supplier Name']
    logger.info("Added 'Awarded Supplier' column.")

    # Apply exclusion rules
    exclusions_bob = [("Supplier 1", "Business Group", "Equal to", "Group 1", True)]
    exclusions_ais = [("Supplier 5", "Product Type", "Equal to", "A", True)]

    # Run analyses
    as_is_df = as_is_analysis(merged_df, column_mapping)
    as_is_df = add_missing_bid_ids(as_is_df, merged_df, column_mapping, 'As-Is')  # Removed 'logger'

    best_of_best_df = best_of_best_analysis(merged_df, column_mapping)
    best_of_best_df = add_missing_bid_ids(best_of_best_df, merged_df, column_mapping, 'Best of Best')  # Removed 'logger'

    best_of_best_excl_df = best_of_best_excluding_suppliers(
        data=merged_df,
        column_mapping=column_mapping,
        excluded_conditions=exclusions_bob
    )
    best_of_best_excl_df = add_missing_bid_ids(
        best_of_best_excl_df,
        merged_df,
        column_mapping,
        'BOB Excl Suppliers'  # Removed 'logger'
    )

    as_is_excl_df = as_is_excluding_suppliers_analysis(
        merged_df,
        column_mapping,
        exclusions_ais
    )
    as_is_excl_df = add_missing_bid_ids(
        as_is_excl_df,
        merged_df,
        column_mapping,
        'As-Is Excl Suppliers'  # Removed 'logger'
    )

    bid_coverage_variations = ["Competitiveness Report", "Supplier Coverage", "Facility Coverage"]
    group_by_field = "Product Type"
    bid_coverage_reports = bid_coverage_report(
        merged_df,
        column_mapping,
        bid_coverage_variations,
        group_by_field
    )

    customizable_grouping_column = "Product Type"
    customizable_df = customizable_analysis(merged_df, column_mapping)

    # Generate Excel output
    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
        as_is_df.to_excel(writer, sheet_name='#As-Is', index=False)
        best_of_best_df.to_excel(writer, sheet_name='#Best of Best', index=False)
        best_of_best_excl_df.to_excel(writer, sheet_name='#BOB Excl Suppliers', index=False)
        as_is_excl_df.to_excel(writer, sheet_name='#As-Is Excl Suppliers', index=False)
        for report_name, report_df in bid_coverage_reports.items():
            sheet_name = report_name.replace(" ", "_")[:31]  # Excel sheet name limit
            report_df.to_excel(writer, sheet_name=sheet_name, index=False)
        customizable_df.to_excel(writer, sheet_name='Customizable Template', index=False)

    # Read the generated Excel for comparison
    generated_excel = pd.read_excel(BytesIO(excel_output.getvalue()), sheet_name=None)

    # Define output directory for saving generated and expected files
    output_dir = Path('tests/test_outputs') / upload_method
    output_dir.mkdir(parents=True, exist_ok=True)

    # Save the generated Excel for manual inspection
    generated_excel_save_path = output_dir / f'generated_excel_output_{upload_method}.xlsx'
    with pd.ExcelWriter(generated_excel_save_path, engine='openpyxl') as writer:
        for sheet_name, df in generated_excel.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    logger.info(f"Generated Excel saved to {generated_excel_save_path}")

    # Check if expected Excel exists
    if not expected_excel_path.exists():
        # If not, save the current generated as expected and skip the test
        with pd.ExcelWriter(expected_excel_path, engine='openpyxl') as writer:
            for sheet_name, df in generated_excel.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"Expected Excel created at {expected_excel_path}")
        pytest.skip("Expected Excel file was missing and has been created. Please rerun the tests.")
    else:
        # Read the expected Excel
        expected_excel = pd.read_excel(expected_excel_path, sheet_name=None)

        # Compare Excel sheets
        try:
            for sheet_name, expected_df in expected_excel.items():
                assert sheet_name in generated_excel, f"Sheet '{sheet_name}' is missing in the generated Excel."
                generated_df = generated_excel[sheet_name]
                # Compare DataFrames
                pd.testing.assert_frame_equal(
                    generated_df.reset_index(drop=True),
                    expected_df.reset_index(drop=True),
                    check_dtype=False,
                    check_like=True,
                    atol=1e-2,  # Allow small numerical differences
                    obj=f"Sheet '{sheet_name}' does not match expected output."
                )
                logger.info(f"Sheet '{sheet_name}' matches expected output.")
        except AssertionError as e:
            # Save the differing sheet for review
            differing_sheet_path = output_dir / f'diff_{sheet_name}_{upload_method}.xlsx'
            with pd.ExcelWriter(differing_sheet_path, engine='openpyxl') as writer:
                generated_df.to_excel(writer, sheet_name='Generated', index=False)
                expected_df.to_excel(writer, sheet_name='Expected', index=False)
            logger.error(str(e))
            logger.info(f"Differing sheet saved to {differing_sheet_path}")
            pytest.fail(str(e))

    # Optionally, handle PowerPoint comparisons similarly
    # This part is not implemented here, but you can add similar logic for PowerPoint files

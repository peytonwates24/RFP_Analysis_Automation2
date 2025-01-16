# tests/test_merged_data.py

import pytest
import pandas as pd
from pathlib import Path
from modules.utils import normalize_columns
from modules.analysis import (
    as_is_analysis,
    best_of_best_analysis,
    best_of_best_excluding_suppliers,
    as_is_excluding_suppliers_analysis,
    bid_coverage_report,
    customizable_analysis,
    add_missing_bid_ids
)
from modules.presentations import *  # if needed
from modules.data_loader import *    # if needed
from io import BytesIO
from pptx import Presentation


@pytest.mark.merged
def test_merged_data(get_test_data_dir, setup_logging):
    """
    Test analysis pipeline using merged data upload method.
    """
    logger = setup_logging
    upload_method = "merged"
    merged_file = get_test_data_dir / 'merged' / 'test_merged_data.xlsx'
    expected_excel_path = Path('tests/test_outputs/merged/expected_excel_output_merged.xlsx')
    expected_ppt = Path('tests/test_outputs/merged/expected_powerpoint_output_merged.pptx')

    # ----------------------------------------------------------------
    # 1. Check if the merged data file exists
    # ----------------------------------------------------------------
    assert merged_file.exists(), f"Merged test file {merged_file} does not exist."
    logger.info(f"Found merged data file: {merged_file}")

    # ----------------------------------------------------------------
    # 2. Load merged data
    # ----------------------------------------------------------------
    merged_df = pd.read_excel(merged_file, engine='openpyxl')
    logger.debug(f"Merged DataFrame columns before normalization: {merged_df.columns.tolist()}")

    # ----------------------------------------------------------------
    # 3. Apply column mapping / normalization
    # ----------------------------------------------------------------
    merged_df = normalize_columns(merged_df)
    column_mapping = {
        'Bid ID': 'Bid ID',
        'Facility': 'Facility',
        'Incumbent': 'Incumbent',
        'Bid Volume': 'Bid Volume',
        'Baseline Price': 'Baseline Price',
        'Current Price': 'Current Price',
        'Bid Price': 'Bid Price',
        'Supplier Capacity': 'Bid Supplier Capacity',  # changed key
        'Supplier Name': 'Bid Supplier Name'           # changed key
    }

    logger.debug(f"Columns after normalization: {merged_df.columns.tolist()}")

    # Ensure all required columns are present
    required_columns = list(column_mapping.values())
    missing_columns = [col for col in required_columns if col not in merged_df.columns]
    assert not missing_columns, f"Missing columns in merged data: {missing_columns}"
    logger.info("All required columns are present in merged data.")

    # Add 'Awarded Supplier' automatically
    merged_df.loc[:, 'Awarded Supplier'] = merged_df['Bid Supplier Name']
    logger.info("Added 'Awarded Supplier' column.")

    # ----------------------------------------------------------------
    # 4. Exclusions and analyses
    # ----------------------------------------------------------------
    exclusions_bob = [("Supplier 1", "Business Group", "Equal to", "Group 1", True)]
    exclusions_ais = [("Supplier 5", "Product Type", "Equal to", "A", True)]

    # As-Is
    as_is_df = as_is_analysis(merged_df, column_mapping)
    as_is_df = add_missing_bid_ids(as_is_df, merged_df, column_mapping, 'As-Is')

    # Best of Best
    best_of_best_df = best_of_best_analysis(merged_df, column_mapping)
    best_of_best_df = add_missing_bid_ids(best_of_best_df, merged_df, column_mapping, 'Best of Best')

    # Best of Best Excl
    best_of_best_excl_df = best_of_best_excluding_suppliers(
        data=merged_df,
        column_mapping=column_mapping,
        excluded_conditions=exclusions_bob
    )
    best_of_best_excl_df = add_missing_bid_ids(
        best_of_best_excl_df,
        merged_df,
        column_mapping,
        'BOB Excl Suppliers'
    )

    # As-Is Excl
    as_is_excl_df = as_is_excluding_suppliers_analysis(merged_df, column_mapping, exclusions_ais)
    as_is_excl_df = add_missing_bid_ids(as_is_excl_df, merged_df, column_mapping, 'As-Is Excl Suppliers')

    # Bid Coverage
    bid_coverage_variations = ["Competitiveness Report", "Supplier Coverage", "Facility Coverage"]
    bid_coverage_reports = bid_coverage_report(
        merged_df,
        column_mapping,
        bid_coverage_variations,
        group_by_field='Product Type'
    )

    # Customizable
    customizable_df = customizable_analysis(merged_df, column_mapping)

    # ----------------------------------------------------------------
    # 5. Generate Excel output (in-memory)
    # ----------------------------------------------------------------
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

    # Read the generated Excel from memory for comparison
    generated_excel = pd.read_excel(BytesIO(excel_output.getvalue()), sheet_name=None)

    # ----------------------------------------------------------------
    # 6. Output directory & stale diff cleanup
    # ----------------------------------------------------------------
    output_dir = Path('tests/test_outputs') / upload_method
    output_dir.mkdir(parents=True, exist_ok=True)

    # Remove any stale diff files from previous runs
    for diff_file in output_dir.glob("diff_*_*.xlsx"):
        try:
            diff_file.unlink()
            logger.info(f"Removed old diff file: {diff_file}")
        except OSError as e:
            logger.warning(f"Could not remove {diff_file}: {e}")

    # ----------------------------------------------------------------
    # 7. Save the newly generated Excel for manual inspection
    # ----------------------------------------------------------------
    generated_excel_save_path = output_dir / f'generated_excel_output_{upload_method}.xlsx'
    with pd.ExcelWriter(generated_excel_save_path, engine='openpyxl') as writer:
        for sheet_name, df in generated_excel.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    logger.info(f"Generated Excel saved to {generated_excel_save_path}")

    # ----------------------------------------------------------------
    # 8. Check if expected Excel exists & compare
    # ----------------------------------------------------------------
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

        # Compare each sheet
        try:
            for sheet_name, expected_df in expected_excel.items():
                assert sheet_name in generated_excel, f"Sheet '{sheet_name}' is missing in the generated Excel."
                gen_df = generated_excel[sheet_name]

                # Compare DataFrames
                pd.testing.assert_frame_equal(
                    gen_df.reset_index(drop=True),
                    expected_df.reset_index(drop=True),
                    check_dtype=False,
                    check_like=True,
                    atol=1e-2,  # small numerical tolerance
                    obj=f"Sheet '{sheet_name}' does not match expected output."
                )
                logger.info(f"Sheet '{sheet_name}' matches expected output.")

        except AssertionError as e:
            # Save the differing sheet for review
            differing_sheet_path = output_dir / f'diff_{sheet_name}_{upload_method}.xlsx'
            with pd.ExcelWriter(differing_sheet_path, engine='openpyxl') as writer:
                gen_df.to_excel(writer, sheet_name='Generated', index=False)
                expected_df.to_excel(writer, sheet_name='Expected', index=False)
            logger.error(str(e))
            logger.info(f"Differing sheet saved to {differing_sheet_path}")
            pytest.fail(str(e))

# tests/test_1.py

import pytest
import pandas as pd
from pathlib import Path
from modules.utils import *
from modules.analysis import *
from modules.presentations import *
from modules.data_loader import *
from io import BytesIO
from pptx import Presentation

@pytest.mark.separate  # or whichever mark you prefer for test_1
def test_run_analysis_separate(get_test_data_dir, setup_logging):
    """
    Test the analysis pipeline with separate bid & baseline files.
    """
    logger = setup_logging
    upload_method = "separate"
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

    # Check if the necessary files exist
    for file in bid_files + [baseline_file]:
        assert file.exists(), f"Required test file {file} does not exist."
        logger.info(f"Found required file: {file}")

    # Load data
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

    # Combine suffixed columns
    merged_df.loc[:, 'Facility'] = merged_df['Facility_baseline'].combine_first(merged_df['Facility_bid'])
    merged_df.loc[:, 'Bid Volume'] = merged_df['Bid Volume_baseline'].combine_first(merged_df['Bid Volume_bid'])
    merged_df.drop(['Facility_baseline', 'Facility_bid', 'Bid Volume_baseline', 'Bid Volume_bid'], axis=1, inplace=True)
    logger.debug(f"Columns after combining suffixed columns: {merged_df.columns.tolist()}")

    # Column mapping
    merged_df = normalize_columns(merged_df)
    column_mapping = {
        'Bid ID': 'Bid ID',
        'Facility': 'Facility',
        'Incumbent': 'Incumbent',
        'Bid Volume': 'Bid Volume',
        'Baseline Price': 'Baseline Price',
        'Current Price': 'Current Price',
        'Bid Price': 'Bid Price',
        'Supplier Capacity': 'Bid Supplier Capacity',
        'Supplier Name': 'Bid Supplier Name'
    }

    # Check columns
    logger.debug(f"Columns after normalization: {merged_df.columns.tolist()}")
    required_columns = list(column_mapping.values())
    missing_columns = [col for col in required_columns if col not in merged_df.columns]
    assert not missing_columns, f"Missing columns in merged data: {missing_columns}"
    logger.info("All required columns are present in merged data.")

    # Add 'Awarded Supplier' automatically
    merged_df.loc[:, 'Awarded Supplier'] = merged_df['Bid Supplier Name']
    logger.info("Added 'Awarded Supplier' column.")

    # Exclusions & analyses
    exclusions_bob = [("Supplier 1", "Business Group", "Equal to", "Group 1", True)]
    exclusions_ais = [("Supplier 5", "Product Type", "Equal to", "A", True)]

    as_is_df = as_is_analysis(merged_df, column_mapping)
    as_is_df = add_missing_bid_ids(as_is_df, merged_df, column_mapping, 'As-Is')

    best_of_best_df = best_of_best_analysis(merged_df, column_mapping)
    best_of_best_df = add_missing_bid_ids(best_of_best_df, merged_df, column_mapping, 'Best of Best')

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

    as_is_excl_df = as_is_excluding_suppliers_analysis(
        merged_df,
        column_mapping,
        exclusions_ais
    )
    as_is_excl_df = add_missing_bid_ids(
        as_is_excl_df,
        merged_df,
        column_mapping,
        'As-Is Excl Suppliers'
    )

    bid_coverage_variations = ["Competitiveness Report", "Supplier Coverage", "Facility Coverage"]
    coverage_reports = bid_coverage_report(
        merged_df,
        column_mapping,
        bid_coverage_variations,
        group_by_field="Product Type"
    )

    customizable_df = customizable_analysis(merged_df, column_mapping)

    # Generate Excel
    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
        as_is_df.to_excel(writer, sheet_name='#As-Is', index=False)
        best_of_best_df.to_excel(writer, sheet_name='#Best of Best', index=False)
        best_of_best_excl_df.to_excel(writer, sheet_name='#BOB Excl Suppliers', index=False)
        as_is_excl_df.to_excel(writer, sheet_name='#As-Is Excl Suppliers', index=False)
        for report_name, report_df in coverage_reports.items():
            sheet_name = report_name.replace(" ", "_")[:31]
            report_df.to_excel(writer, sheet_name=sheet_name, index=False)
        customizable_df.to_excel(writer, sheet_name='Customizable Template', index=False)

    generated_excel = pd.read_excel(BytesIO(excel_output.getvalue()), sheet_name=None)

    # Output dir & stale diff cleanup
    output_dir = Path('tests/test_outputs') / upload_method
    output_dir.mkdir(parents=True, exist_ok=True)

    # Remove stale diff files if any
    for diff_file in output_dir.glob("diff_*_*.xlsx"):
        try:
            diff_file.unlink()
            logger.info(f"Removed old diff file: {diff_file}")
        except OSError as e:
            logger.warning(f"Could not remove {diff_file}: {e}")

    # Save generated Excel
    generated_excel_save_path = output_dir / f'generated_excel_output_{upload_method}.xlsx'
    with pd.ExcelWriter(generated_excel_save_path, engine='openpyxl') as writer:
        for sheet_name, df in generated_excel.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    logger.info(f"Generated Excel saved to {generated_excel_save_path}")

    # Compare to expected
    if not expected_excel_path.exists():
        with pd.ExcelWriter(expected_excel_path, engine='openpyxl') as writer:
            for sheet_name, df in generated_excel.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"Expected Excel created at {expected_excel_path}")
        pytest.skip("Expected Excel file was missing; created. Please rerun.")

    else:
        expected_excel = pd.read_excel(expected_excel_path, sheet_name=None)
        try:
            for sheet_name, expected_df in expected_excel.items():
                assert sheet_name in generated_excel, f"Sheet '{sheet_name}' missing in generated output."
                gen_df = generated_excel[sheet_name]
                pd.testing.assert_frame_equal(
                    gen_df.reset_index(drop=True),
                    expected_df.reset_index(drop=True),
                    check_dtype=False,
                    check_like=True,
                    atol=1e-2,
                    obj=f"Sheet '{sheet_name}' does not match expected output."
                )
                logger.info(f"Sheet '{sheet_name}' matches expected output.")
        except AssertionError as e:
            diff_file = output_dir / f'diff_{sheet_name}_{upload_method}.xlsx'
            with pd.ExcelWriter(diff_file, engine='openpyxl') as writer:
                gen_df.to_excel(writer, sheet_name='Generated', index=False)
                expected_df.to_excel(writer, sheet_name='Expected', index=False)
            logger.error(str(e))
            logger.info(f"Differing sheet saved to {diff_file}")
            pytest.fail(str(e))

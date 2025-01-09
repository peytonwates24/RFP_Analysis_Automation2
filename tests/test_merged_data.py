# tests/test_merged_data.py

import pytest
import pandas as pd
from pathlib import Path
from modules.utils import normalize_columns
from modules.analysis import *
from modules.presentations import *
from modules.data_loader import *
from io import BytesIO
from pptx import Presentation

@pytest.mark.merged
def test_merged_data(get_test_data_dir, setup_logging):
    """
    Test analysis pipeline using merged data upload method.
    """
    logger = setup_logging
    merged_data_file = get_test_data_dir / 'merged' / 'test_merged_data.xlsx'
    expected_excel = get_test_data_dir / 'merged' / 'expected_excel_output_merged.xlsx'
    expected_ppt = get_test_data_dir / 'merged' / 'expected_powerpoint_output_merged.pptx'

    # Check if the merged data file exists
    assert merged_data_file.exists(), f"Merged test file {merged_data_file} does not exist."
    logger.info(f"Found merged data file: {merged_data_file}")

    # Load merged data
    merged_df = pd.read_excel(merged_data_file, engine='openpyxl')
    logger.debug(f"Merged DataFrame columns: {merged_df.columns.tolist()}")

    # Apply column mapping
    merged_df = normalize_columns(merged_df)
    logger.info("Applied column normalization.")
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
    merged_df['Awarded Supplier'] = merged_df['Bid Supplier Name']
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

    generated_excel = pd.read_excel(BytesIO(excel_output.getvalue()), sheet_name=None)
    expected_excel = pd.read_excel(expected_excel, sheet_name=None)

    # Compare Excel sheets
    for sheet_name, expected_df in expected_excel.items():
        assert sheet_name in generated_excel, f"Sheet '{sheet_name}' is missing in the generated Excel."
        generated_df = generated_excel[sheet_name]
        # Compare DataFrames
        try:
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
            logger.error(str(e))
            pytest.fail(str(e))

    # Generate PowerPoint output
    prs = Presentation()
    # Create slides based on analyses
    prs = create_scenario_summary_presentation(generated_excel, 'Business Group')  # Adjust as per your function
    prs = create_bid_coverage_summary_slides(prs, merged_df, 'Product Type')
    prs = create_supplier_comparison_summary_slide(prs, merged_df, 'Business Group')

    ppt_output = BytesIO()
    prs.save(ppt_output)
    generated_ppt = Presentation(BytesIO(ppt_output.getvalue()))
    expected_ppt_presentation = Presentation(expected_ppt)

    # Compare PowerPoint slides
    assert len(generated_ppt.slides) == len(expected_ppt_presentation.slides), "Number of slides in PowerPoint does not match expected."

    for i, (gen_slide, exp_slide) in enumerate(zip(generated_ppt.slides, expected_ppt_presentation.slides)):
        # Compare slide titles
        gen_title = None
        exp_title = None
        for shape in gen_slide.shapes:
            if hasattr(shape, "text") and shape.text:
                gen_title = shape.text
                break
        for shape in exp_slide.shapes:
            if hasattr(shape, "text") and shape.text:
                exp_title = shape.text
                break
        assert gen_title == exp_title, f"Slide {i+1} title mismatch: expected '{exp_title}', got '{gen_title}'."
        logger.info(f"Slide {i+1} title matches expected: '{gen_title}'.")

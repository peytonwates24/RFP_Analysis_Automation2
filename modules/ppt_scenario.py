# ppt_scenario.py

import os
import pandas as pd
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from modules.presentations import *
from pptx import Presentation
from pptx.enum.chart import XL_LABEL_POSITION



LIGHT_BLUE = RGBColor(173, 216, 230)  

from typing import Optional, Dict

def create_scenario_summary_slides(
    prs: Presentation,
    scenario_dataframes: Dict[str, pd.DataFrame],
    scenario_detail_grouping: Optional[str],
    title_suffix: str = "",
    create_details: bool = True,
    logger=logger
) -> Presentation:
    logger.info("→ Entering create_scenario_summary_slides()")
    scenario_keys = list(scenario_dataframes.keys())
    scenarios    = [k.lstrip("#") for k in scenario_keys]
    logger.info(f"Found scenarios: {scenarios!r}")

    scenarios_per_slide = 3
    total_slides       = (len(scenarios) + scenarios_per_slide - 1) // scenarios_per_slide
    logger.info(f"Total slides to create: {total_slides}")

    # find the Blank layout once (by name), fallback to index 6 if necessary
    blank_layout_idx = next(
        (i for i, lay in enumerate(prs.slide_layouts) if lay.name.lower() == "blank"),
        6
    )

    scenario_index = 0
    for slide_num in range(1, total_slides + 1):
        logger.info(f"--- Building slide {slide_num}/{total_slides} ---")

        # always add a fresh blank slide
        slide = prs.slides.add_slide(prs.slide_layouts[blank_layout_idx])
        logger.info("Added new blank slide")

        # header & labels
        logger.info("Adding header + row labels")
        add_header(slide, slide_num, title_suffix)
        add_row_labels(slide)

        scenario_names       = []
        ast_savings_list     = []
        current_savings_list = []

        for col_idx in range(scenarios_per_slide):
            if scenario_index >= len(scenarios):
                logger.info("No more scenarios to place on this slide")
                break

            name = scenarios[scenario_index]
            key  = scenario_keys[scenario_index]
            df   = scenario_dataframes[key]
            logger.info(f"Processing scenario #{scenario_index}: '{name}' (key={key})")

            # sanitize and prepare
            df = process_scenario_dataframe(df, scenario_detail_grouping, require_grouping=create_details)
            df = df.fillna(0).replace([float('inf'), float('-inf')], 0)
            logger.info(f"  DataFrame columns after processing: {df.columns.tolist()}")

            scen_dict = _prepare_scenario_dict(df, name)
            logger.info(f"  scen_dict: {scen_dict}")

            # place content
            logger.info(f"  add_scenario_content at col {col_idx+1}")
            add_scenario_content(slide, scen_dict, scenario_position=col_idx + 1)

            logger.info(f"  add_small_bar_chart at col {col_idx+1}")
            add_small_bar_chart(
                slide,
                {
                    "savings_values": [df["AST Savings"].sum(), df["Current Savings"].sum()],
                    "rebate_values":  [df["Rebate Savings"].sum()] * 2
                },
                scenario_position=col_idx + 1
            )

            # collect values
            ast = float(df["AST Savings"].sum())
            cur = float(df["Current Savings"].sum())
            scenario_names.append(name)
            ast_savings_list.append(ast)
            current_savings_list.append(cur)
            logger.info(f"  Collected for chart: AST={ast}, Current={cur}")

            # optional detail
            if create_details:
                logger.info(f"  Adding detail slide for '{name}'")
                add_scenario_detail_slide(
                    prs=prs,
                    df=df,
                    scenario_name=name,
                    template_slide_layout_index=blank_layout_idx,
                    scenario_detail_grouping=scenario_detail_grouping
                )

            scenario_index += 1

        logger.info(
            f"Adding clustered chart on slide {slide_num}: "
            f"names={scenario_names}, ASTs={ast_savings_list}, Currents={current_savings_list}"
        )
        add_chart(slide, scenario_names, ast_savings_list, current_savings_list)

    logger.info("← Exiting create_scenario_summary_slides()")
    return prs





def add_header(slide, slide_number: int, title_suffix: str):
    """A large header at the top of the slide (left-aligned, large text)."""
    left = Inches(0)
    top = Inches(0.01)
    width = Inches(12.5)
    height = Inches(0.5)
    header = slide.shapes.add_textbox(left, top, width, height)
    header_tf = header.text_frame
    header_tf.vertical_anchor = MSO_ANCHOR.TOP

    p = header_tf.paragraphs[0]
    p.text = f"Scenario Summary — Slide {slide_number}{title_suffix}"
    p.font.size = Pt(25)
    p.font.name = "Calibri"
    p.font.color.rgb = RGBColor(0, 51, 153)  # Dark Blue
    p.font.bold = True

def add_row_labels(slide):
    """
    Labels on the left side. 
    Changed the 4th label to 'RFP Savings % (Baseline|Current)' to reflect 
    that we're now showing both baseline and current in a single text field.
    """
    row_labels = [
        'Scenario Name',
        'Description',
        '# of Suppliers (% spend)',
        'RFP Savings % (Baseline|Current)',
        'Total Value Opportunity ($MM)',
        'Key Considerations'
    ]
    top_positions = [
        Inches(0.8),
        Inches(1.3),
        Inches(1.9),
        Inches(3.2),  # RFP Savings row
        Inches(3.8),
        Inches(4.8)
    ]
    left = Inches(0.58)
    width = Inches(1.7)
    height = Inches(0.3)

    for label_text, top in zip(row_labels, top_positions):
        label_box = slide.shapes.add_textbox(left, top, width, height)
        tf = label_box.text_frame
        p = tf.paragraphs[0]
        p.text = label_text
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.name = "Calibri"
        p.font.color.rgb = RGBColor(0, 51, 153)  # Dark Blue
        label_box.fill.background()
        label_box.line.fill.background()

def add_scenario_content(slide, scenario_dict: dict, scenario_position: int):
    """
    Places all text fields for one scenario column on the summary slide.

    Parameters:
    - slide: pptx Slide object
    - scenario_dict: dict with keys:
        "scenario_name", "description", "num_suppliers_text",
        "rfp_savings_text", "value_opportunity_text", "key_considerations"
    - scenario_position: 1, 2 or 3 (which column on the slide)
    """
    logger.info(f"    → add_scenario_content(position={scenario_position})")

    # Compute left offset for this column
    base_left = Inches(2.8) + (scenario_position - 1) * Inches(3.1)
    logger.info(f"      • base_left = {base_left}")

    # Row-top positions
    row_tops = {
        "scenario_name": Inches(0.8),
        "description":   Inches(1.3),
        "num_suppliers": Inches(1.9),
        "rfp_savings":   Inches(3.2),
        "value_oppty":   Inches(3.8),
        "key_cons":      Inches(4.8),
    }

    # 1) Scenario name
    name = scenario_dict.get("scenario_name", "")
    logger.info(f"      • adding scenario_name='{name}'")
    sn_box = slide.shapes.add_textbox(base_left, row_tops["scenario_name"], Inches(2.5), Inches(0.35))
    sn_tf = sn_box.text_frame
    p = sn_tf.paragraphs[0]
    p.text = name
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.name = "Calibri"
    p.font.color.rgb = RGBColor(0, 51, 153)

    # 2) Description
    desc = scenario_dict.get("description", "")
    logger.info(f"      • adding description='{desc}'")
    desc_box = slide.shapes.add_textbox(base_left, row_tops["description"], Inches(2.5), Inches(0.5))
    desc_tf = desc_box.text_frame
    p = desc_tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(12)
    p.font.name = "Calibri"
    p.alignment = PP_ALIGN.CENTER

    # 3) # of Suppliers
    sup_txt = scenario_dict.get("num_suppliers_text", "")
    logger.info(f"      • adding num_suppliers_text='{sup_txt}'")
    sup_box = slide.shapes.add_textbox(base_left, row_tops["num_suppliers"], Inches(2.5), Inches(1.0))
    sup_tf = sup_box.text_frame
    p = sup_tf.paragraphs[0]
    p.text = sup_txt
    p.font.size = Pt(12)
    p.font.name = "Calibri"
    p.alignment = PP_ALIGN.CENTER

    # 4) RFP Savings % (Baseline | Current)
    rfp_txt = scenario_dict.get("rfp_savings_text", "")
    logger.info(f"      • adding rfp_savings_text='{rfp_txt}'")
    rfp_box = slide.shapes.add_textbox(base_left, row_tops["rfp_savings"], Inches(2.5), Inches(0.5))
    rfp_tf = rfp_box.text_frame
    p = rfp_tf.paragraphs[0]
    p.text = rfp_txt
    p.font.size = Pt(12)
    p.font.name = "Calibri"
    p.alignment = PP_ALIGN.CENTER

    # 5) Total Value Opportunity
    val_txt = scenario_dict.get("value_opportunity_text", "")
    logger.info(f"      • adding value_opportunity_text='{val_txt}'")
    val_box = slide.shapes.add_textbox(base_left, row_tops["value_oppty"], Inches(2.5), Inches(0.5))
    val_tf = val_box.text_frame
    p = val_tf.paragraphs[0]
    p.text = val_txt
    p.font.size = Pt(12)
    p.font.name = "Calibri"
    p.alignment = PP_ALIGN.CENTER

    # 6) Key Considerations
    key_txt = scenario_dict.get("key_considerations", "")
    logger.info(f"      • adding key_considerations='{key_txt}'")
    key_box = slide.shapes.add_textbox(base_left, row_tops["key_cons"], Inches(2.5), Inches(0.8))
    key_tf = key_box.text_frame
    p = key_tf.paragraphs[0]
    p.text = key_txt
    p.font.size = Pt(12)
    p.font.name = "Calibri"
    p.alignment = PP_ALIGN.LEFT

    logger.info("    ← done add_scenario_content")



# ────────────────────────────────────────────────────────────────
# BAR-CHART helper
# ────────────────────────────────────────────────────────────────
from pptx.dml.color import RGBColor

# ------------------------------------------------------------
# Tiny stacked bar-chart (AST | Current with rebate overlay)
# ------------------------------------------------------------
def add_small_bar_chart(slide, scen_dict: dict, scenario_position: int):
    """
    Draws a tiny 2-category stacked column chart at the bottom of a scenario column.

    Parameters:
    - slide: pptx Slide object
    - scen_dict: dict with keys "savings_values" (list of two floats) and "rebate_values" (list of two floats)
    - scenario_position: 1, 2 or 3 (which column on the slide)
    """
    logger.info(f"    → add_small_bar_chart(position={scenario_position})")

    savings_vals = scen_dict["savings_values"]
    rebate_vals  = scen_dict["rebate_values"]
    logger.info(f"      • savings_vals={savings_vals}, rebate_vals={rebate_vals}")

    categories = ["AST", "Current"]
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Savings", savings_vals)
    chart_data.add_series("Rebates", rebate_vals)

    # compute chart position
    left  = Inches(2.8) + (scenario_position - 1) * Inches(3.1)
    top   = Inches(5.7)
    width = Inches(2.5)
    height= Inches(1.2)
    logger.info(f"      • chart bounds L={left}, T={top}, W={width}, H={height}")

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        left, top, width, height,
        chart_data
    ).chart

    # color the two series
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1  # dark blue
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2  # light blue

    # remove axes & legend
    chart.category_axis.visible = False
    chart.value_axis.visible    = False
    chart.has_legend            = False

    # add data labels only on the Savings series
    for idx, point in enumerate(chart.series[0].points):
        val = savings_vals[idx]
        point.data_label.number_format_is_linked = False
        point.data_label.number_format = '"$"#,##0.0,,"MM"'
        point.data_label.position = XL_LABEL_POSITION.OUTSIDE_END
        logger.info(f"      • data label for point[{idx}] = {val}")

    logger.info("    ← done add_small_bar_chart")



# ──────────────────────────────────────────────────────────────
# Helper: convert one scenario DataFrame → dict for the slides
# ──────────────────────────────────────────────────────────────
def _prepare_scenario_dict(df: pd.DataFrame, scen_name: str) -> dict:
    # …
    # 1) Ensure our *new* columns are present
    required_cols = [
        "AST Savings", "Current Savings",
        "Rebate Savings", "AST Baseline Spend"
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = 0.0

    # 2) Sum from the post-processed columns
    total_ast   = float(df["AST Savings"].sum())
    total_cur   = float(df["Current Savings"].sum())
    total_reb   = float(df["Rebate Savings"].sum())
    total_base  = float(df["AST Baseline Spend"].sum())

    # 3) Percent calculations
    pct_ast   = (total_ast / total_base) if total_base else 0.0
    pct_cur   = (total_cur / total_base) if total_base else 0.0

    # 4) Build the dictionary exactly the same shape as before
    scen_dict = {
        "scenario_name":         scen_name.lstrip("#"),
        "description":           "",   # still blank by default
        "num_suppliers_text":    "",   # optional
        "rfp_savings_text":      f"{pct_ast:.1%} | {pct_cur:.1%}",
        "value_opportunity_text":f"${(total_ast + total_reb)/1e6:,.1f} MM",
        "key_considerations":    "",

        # values for the little bar-chart
        "savings_values": [total_ast, total_cur],
        "rebate_values":  [total_reb, total_reb],
    }
    return scen_dict



def process_scenario_dataframe(df, scenario_detail_grouping, require_grouping=True):
    """
    Processes the scenario DataFrame to ensure required columns are present.
    If require_grouping=False, we do not insist on the scenario_detail_grouping column.

    Parameters:
    - df: DataFrame to process.
    - scenario_detail_grouping: The grouping column name used for scenario details.
    - require_grouping: boolean, if True scenario_detail_grouping must be included, otherwise it's optional.

    Returns:
    - df: Processed DataFrame with required columns added if necessary.
    """

    # 1) Strip whitespace from column names
    df.columns = df.columns.str.strip()

    # 2) Rename any generic columns
    df = df.rename(columns={
        'Awarded Supplier': 'Awarded Supplier Name'
    })

    # 3) Normalize awarded‐supplier price field
    if 'Discounted Awarded Supplier Price' in df.columns:
        df = df.rename(columns={'Discounted Awarded Supplier Price': 'Awarded Supplier Price'})
    elif 'Original Awarded Supplier Price' in df.columns:
        df = df.rename(columns={'Original Awarded Supplier Price': 'Awarded Supplier Price'})

    # 4) Normalize savings column names
    if 'Baseline Savings' in df.columns:
        df = df.rename(columns={'Baseline Savings': 'AST Savings'})
    if 'Current Price Savings' in df.columns:
        df = df.rename(columns={'Current Price Savings': 'Current Savings'})

    # Base expected columns
    expected_columns = [
        'Awarded Supplier Name',
        'Awarded Supplier Spend',
        'AST Savings',
        'Current Savings',
        'AST Baseline Spend',
        'Current Baseline Spend'
    ]
    if require_grouping and scenario_detail_grouping:
        expected_columns.append(scenario_detail_grouping)

    # 5) Awarded Supplier Spend
    if 'Awarded Supplier Spend' not in df.columns:
        if 'Awarded Supplier Price' in df.columns and 'Awarded Volume' in df.columns:
            df['Awarded Supplier Spend'] = df['Awarded Supplier Price'] * df['Awarded Volume']
        else:
            df['Awarded Supplier Spend'] = 0

    # 6) AST Baseline Spend
    if 'AST Baseline Spend' not in df.columns:
        if 'Baseline Spend' in df.columns:
            df['AST Baseline Spend'] = df['Baseline Spend']
        elif 'Baseline Price' in df.columns and 'Bid Volume' in df.columns:
            df['AST Baseline Spend'] = df['Baseline Price'] * df['Bid Volume']
        else:
            df['AST Baseline Spend'] = 0

    # 7) AST Savings
    if 'AST Savings' not in df.columns:
        df['AST Savings'] = df['AST Baseline Spend'] - df['Awarded Supplier Spend']

    # 8) Current Baseline Spend
    if 'Current Baseline Spend' not in df.columns:
        if 'Current Price' in df.columns and 'Bid Volume' in df.columns:
            df['Current Baseline Spend'] = df['Current Price'] * df['Bid Volume']
        else:
            df['Current Baseline Spend'] = df['AST Baseline Spend']

    # 9) Current Savings
    if 'Current Savings' not in df.columns:
        df['Current Savings'] = df['Current Baseline Spend'] - df['Awarded Supplier Spend']

    # 10) Ensure Awarded Volume exists
    if 'Awarded Volume' not in df.columns:
        if 'Bid Volume' in df.columns:
            df['Awarded Volume'] = df['Bid Volume']
        else:
            df['Awarded Volume'] = 0

    # 11) Ensure Incumbent & Bid ID exist
    if 'Incumbent' not in df.columns:
        df['Incumbent'] = 'Unknown'
    if 'Bid ID' not in df.columns:
        df['Bid ID'] = df.index

    # ── NEW: ensure 'Bid Volume' exists (needed by detail slides)
    if 'Bid Volume' not in df.columns:
        if 'Awarded Volume' in df.columns:
            df['Bid Volume'] = df['Awarded Volume']
        else:
            df['Bid Volume'] = 0

    # 12) Backfill any missing expected columns
    for col in expected_columns:
        if col not in df.columns:
            if col == scenario_detail_grouping and require_grouping:
                raise ValueError(f"Required grouping column '{scenario_detail_grouping}' not found in data.")
            df[col] = 0

    return df



def create_scenario_detail_slides(prs, scenario_dataframes, scenario_detail_grouping, logger=logger):
    """
    After the summary slides are done, call this to append one detail slide per scenario,
    grouping each scenario’s DataFrame by `scenario_detail_grouping`.
    """
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.dml.color import RGBColor
    from collections import OrderedDict

    for sheet_key, df in scenario_dataframes.items():
        scenario_name = sheet_key.lstrip("#")
        # ensure the grouping column exists
        if scenario_detail_grouping not in df.columns:
            logger.warning(f"Skipping detail slides for '{scenario_name}'—no '{scenario_detail_grouping}' column.")
            continue

        # make sure numeric & spend columns are present
        df = process_scenario_dataframe(df, scenario_detail_grouping, require_grouping=True)

        # add one detail slide per scenario
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        for shape in list(slide.shapes):
            shape.element.getparent().remove(shape.element)

        add_scenario_detail_slide(
            prs=prs,
            df=df,
            scenario_name=scenario_name,
            template_slide_layout_index=1,
            scenario_detail_grouping=scenario_detail_grouping
        )
    return prs

# ppt_scenario.py

import os
import pandas as pd

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from modules.presentations import *


LIGHT_BLUE = RGBColor(173, 216, 230)  

def create_scenario_summary_slides(
    prs,
    scenario_data_list,
    title="Scenario Summary"
):
    """
    Creates 'Scenario Summary' slides. Each slide can hold up to 3 scenarios
    side-by-side in columns. Each scenario displays:

      - Scenario Name
      - Description
      - #Suppliers (% spend)
      - RFP Savings % (Baseline|Current)
      - Total Value Opportunity ($MM)
      - Key Considerations
      - A two-series bar chart at the bottom (e.g. baseline savings vs current savings)

    :param prs: pptx.Presentation object (already created or loaded from a template).
    :param scenario_data_list: List of scenario dictionaries, each describing how to fill the slide.
       For each scenario, define keys:
         {
           "scenario_name": str,
           "description": str,
           "num_suppliers_text": str,       # e.g. "4 - Supplier2(60%), Supplier1(30%)..."
           "rfp_savings_text": str,         # e.g. "7% | 12%" for Baseline vs Current
           "value_opportunity_text": str,   # e.g. "$0.3MM"
           "key_considerations": str,       # e.g. "Key considerations..."
           
           # For the bottom chart: 2 bars => [ total_baseline_savings, total_current_savings ]
           "bar_chart_values": [float, float],  
           # e.g. [ 300000.0, 100000.0 ] => 0.3MM vs 0.1MM
           "bar_chart_labels": [ "Baseline", "Current" ]
         }
    :param title: The large header at the top of each slide. Defaults to "Scenario Summary".

    :return: updated pptx.Presentation object with new slides added.
    """

    def add_header(slide, slide_title):
        """A large header at the top of the slide (left-aligned, large text)."""
        left = Inches(0)
        top = Inches(0.01)
        width = Inches(12.5)
        height = Inches(0.5)
        header = slide.shapes.add_textbox(left, top, width, height)
        header_tf = header.text_frame
        header_tf.vertical_anchor = MSO_ANCHOR.TOP

        p = header_tf.paragraphs[0]
        p.text = slide_title
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

    def add_scenario_content(slide, scenario_dict, scenario_position):
        """
        scenario_dict keys:
         - "scenario_name"
         - "description"
         - "num_suppliers_text"
         - "rfp_savings_text" (like "7% | 12%")
         - "value_opportunity_text"
         - "key_considerations"
         - "bar_chart_values"
         - "bar_chart_labels"

        scenario_position = 1, 2, or 3 => which column on the slide
        """
        base_left = Inches(2.8) + (scenario_position - 1)*Inches(3.1)

        # Each row's top
        row_top_positions = {
            "scenario_name":  Inches(0.8),
            "description":    Inches(1.3),
            "num_suppliers":  Inches(1.9),
            "rfp_savings":    Inches(3.2),
            "value_oppty":    Inches(3.8),
            "key_cons":       Inches(4.8)
        }

        # Scenario Name
        sn_box = slide.shapes.add_textbox(base_left, row_top_positions["scenario_name"], Inches(2.5), Inches(0.35))
        sn_tf = sn_box.text_frame
        p = sn_tf.paragraphs[0]
        p.text = scenario_dict.get("scenario_name", "")
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.name = "Calibri"
        p.font.color.rgb = RGBColor(0, 51, 153)

        # Description
        desc_box = slide.shapes.add_textbox(base_left, row_top_positions["description"], Inches(2.5), Inches(0.5))
        desc_tf = desc_box.text_frame
        p = desc_tf.paragraphs[0]
        p.text = scenario_dict.get("description", "")
        p.font.size = Pt(12)
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER

        # #Suppliers
        sup_box = slide.shapes.add_textbox(base_left, row_top_positions["num_suppliers"], Inches(2.5), Inches(1.0))
        sup_tf = sup_box.text_frame
        p = sup_tf.paragraphs[0]
        p.text = scenario_dict.get("num_suppliers_text", "")
        p.font.size = Pt(12)
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER

        # RFP Savings (Baseline|Current) => e.g. "7% | 12%"
        sav_box = slide.shapes.add_textbox(base_left, row_top_positions["rfp_savings"], Inches(2.5), Inches(0.5))
        sav_tf = sav_box.text_frame
        p = sav_tf.paragraphs[0]
        p.text = scenario_dict.get("rfp_savings_text", "")
        p.font.size = Pt(12)
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER

        # total value
        tv_box = slide.shapes.add_textbox(base_left, row_top_positions["value_oppty"], Inches(2.5), Inches(0.5))
        tv_tf = tv_box.text_frame
        p = tv_tf.paragraphs[0]
        p.text = scenario_dict.get("value_opportunity_text", "")
        p.font.size = Pt(12)
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER

        # Key Considerations
        kc_box = slide.shapes.add_textbox(base_left, row_top_positions["key_cons"], Inches(2.5), Inches(0.8))
        kc_tf = kc_box.text_frame
        p = kc_tf.paragraphs[0]
        p.text = scenario_dict.get("key_considerations", "")
        p.font.size = Pt(12)
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.LEFT

    # ────────────────────────────────────────────────────────────────
    # BAR-CHART helper
    # ────────────────────────────────────────────────────────────────
    from pptx.dml.color import RGBColor
    LIGHT_BLUE = RGBColor(173, 216, 230)   # rebate segment colour

    # ------------------------------------------------------------
    # Tiny stacked bar-chart (AST | Current with rebate overlay)
    # ------------------------------------------------------------
    def add_small_bar_chart(slide, scen_dict: dict, scenario_position: int):
        """
        Draws a 2-category stacked column chart at the bottom of a scenario column:

            – savings_values : [AST $, Current $]
            – rebate_values  : [rebate $, rebate $]   (stacked on both)

        Colors:
            • savings part  → Theme ACCENT_1 (dark blue)
            • rebate part   → Theme ACCENT_2 (light blue)
        """
        savings_vals = scen_dict["savings_values"]
        rebate_vals  = scen_dict["rebate_values"]
        categories   = ["AST", "Current"]

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series("Savings", savings_vals)
        chart_data.add_series("Rebates", rebate_vals)

        # ── position (matches earlier layout) ────────────────────
        left  = Inches(2.8) + (scenario_position - 1) * Inches(3.1)
        top   = Inches(5.7)
        width = Inches(2.5)
        height= Inches(1.2)

        chart = (
            slide.shapes
            .add_chart(XL_CHART_TYPE.COLUMN_STACKED, left, top, width, height, chart_data)
            .chart
        )

        # color series
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1  # dark blue

        chart.series[1].format.fill.solid()
        chart.series[1].format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2  # light blue

        # legend & axes
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False

        chart.category_axis.visible = False
        chart.value_axis.visible    = False

        # data labels → show $MM on Savings only (top of stack)
        for point, val in zip(chart.series[0].points, savings_vals):
            point.data_label.number_format_is_linked = False
            point.data_label.number_format = '"$"#,##0.0,,"MM"'
            point.data_label.position = XL_LABEL_POSITION.OUTSIDE_END




    # ────────────────────────────────────────────────────────────────
    # MAIN LOGIC — assemble all slides
    # ────────────────────────────────────────────────────────────────
    scenarios_per_slide = 3
    total_slides = (len(scenario_data_list) + scenarios_per_slide - 1) // scenarios_per_slide
    scenario_index = 0

    for slide_idx in range(total_slides):
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        # Remove template placeholders (we place everything manually)
        for shape in list(slide.shapes):
            shape.element.getparent().remove(shape.element)

        add_header(slide, title)
        add_row_labels(slide)

        for col_idx in range(scenarios_per_slide):
            if scenario_index >= len(scenario_data_list):
                break

            scenario_dict = scenario_data_list[scenario_index]

            # text content
            add_scenario_content(slide, scenario_dict, scenario_position=col_idx + 1)
            # stacked bar-chart
            add_small_bar_chart(slide, scenario_dict, scenario_position=col_idx + 1)

            scenario_index += 1

    return prs


# ──────────────────────────────────────────────────────────────
# Helper: convert one scenario DataFrame → dict for the slides
# ──────────────────────────────────────────────────────────────
def _prepare_scenario_dict(df: "pd.DataFrame", scen_name: str) -> dict:
    """
    Expected numeric columns in *df* (will be auto-created as 0.0 if absent):

        • 'Baseline Savings'        – savings vs baseline (AST)
        • 'Current Price Savings'   – savings vs current price
        • 'Rebate Savings'          – rebate-driven savings $
        • 'Baseline Spend'          – total baseline spend

    The function returns a dictionary that the updated
    `create_scenario_summary_slides()` knows how to consume.
    """

    # ------------------------------------------------------------------
    # 1) Ensure all required columns exist so .sum() never errors
    # ------------------------------------------------------------------
    required_cols = [
        "Baseline Savings", "Current Price Savings",
        "Rebate Savings", "Baseline Spend"
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = 0.0

    # ------------------------------------------------------------------
    # 2) Aggregate scenario-level numbers
    # ------------------------------------------------------------------
    total_baseline_save = float(df["Baseline Savings"].sum())          # AST savings
    total_current_save  = float(df["Current Price Savings"].sum())     # Current-price savings
    total_rebates       = float(df["Rebate Savings"].sum())            # rebate $ (single bucket)
    total_base_spend    = float(df["Baseline Spend"].sum())

    pct_ast    = (total_baseline_save / total_base_spend) if total_base_spend else 0.0
    pct_current= (total_current_save  / total_base_spend) if total_base_spend else 0.0

    # ------------------------------------------------------------------
    # 3) Build the dictionary
    # ------------------------------------------------------------------
    scen_dict = {
        # ── text fields shown in the table ────────────────────────────
        "scenario_name":        scen_name.lstrip("#"),          # strip the leading '#'
        "description":          "",                             # fill later as desired
        "num_suppliers_text":   "",                             # optional
        "rfp_savings_text":     f"{pct_ast:.1%} | {pct_current:.1%}",  # AST | Current
        "value_opportunity_text": f"${(total_baseline_save + total_rebates)/1e6:,.1f} MM",
        "key_considerations":   "",

        # ── numeric series for the tiny stacked-column chart ─────────
        # 2 categories → 0 = AST, 1 = Current
        "savings_values": [total_baseline_save, total_current_save],
        "rebate_values":  [total_rebates,      total_rebates]   # same rebate stacked on both
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

    # Ensure columns are stripped
    df.columns = df.columns.str.strip()

    # Map existing columns to required columns
    df = df.rename(columns={
        'Awarded Supplier': 'Awarded Supplier Name'
        # Add more mappings if necessary
    })

    # Base expected columns
    expected_columns = [
        'Awarded Supplier Name',
        'Awarded Supplier Spend',
        'AST Savings',
        'Current Savings',
        'AST Baseline Spend',
        'Current Baseline Spend'
    ]

    # Only include scenario_detail_grouping if require_grouping is True
    if require_grouping and scenario_detail_grouping:
        expected_columns.append(scenario_detail_grouping)

    # Calculate 'Awarded Supplier Spend' if not present
    if 'Awarded Supplier Spend' not in df.columns:
        if 'Awarded Supplier Price' in df.columns and 'Awarded Volume' in df.columns:
            df['Awarded Supplier Spend'] = df['Awarded Supplier Price'] * df['Awarded Volume']
        else:
            df['Awarded Supplier Spend'] = 0

    # 'AST Baseline Spend'
    if 'AST Baseline Spend' not in df.columns:
        if 'Baseline Spend' in df.columns:
            df['AST Baseline Spend'] = df['Baseline Spend']
        elif 'Baseline Price' in df.columns and 'Bid Volume' in df.columns:
            df['AST Baseline Spend'] = df['Baseline Price'] * df['Bid Volume']
        else:
            df['AST Baseline Spend'] = 0

    # 'AST Savings'
    if 'AST Savings' not in df.columns:
        df['AST Savings'] = df['AST Baseline Spend'] - df['Awarded Supplier Spend']

    # 'Current Baseline Spend'
    if 'Current Baseline Spend' not in df.columns:
        if 'Current Price' in df.columns and 'Bid Volume' in df.columns:
            df['Current Baseline Spend'] = df['Current Price'] * df['Bid Volume']
        else:
            df['Current Baseline Spend'] = df['AST Baseline Spend']

    # 'Current Savings'
    if 'Current Savings' not in df.columns:
        df['Current Savings'] = df['Current Baseline Spend'] - df['Awarded Supplier Spend']

    # Ensure 'Awarded Volume' exists
    if 'Awarded Volume' not in df.columns:
        if 'Bid Volume' in df.columns:
            df['Awarded Volume'] = df['Bid Volume']
        else:
            df['Awarded Volume'] = 0

    # Ensure 'Incumbent' exists
    if 'Incumbent' not in df.columns:
        df['Incumbent'] = 'Unknown'

    # Ensure 'Bid ID' exists
    if 'Bid ID' not in df.columns:
        df['Bid ID'] = df.index

    # Check expected columns
    for col in expected_columns:
        if col not in df.columns:
            if col == scenario_detail_grouping and require_grouping:
                # If grouping is required but not found, raise error
                raise ValueError(f"Required grouping column '{scenario_detail_grouping}' not found in data.")
            else:
                # For non-grouping columns, just fill with 0 if not found
                df[col] = 0

    return df
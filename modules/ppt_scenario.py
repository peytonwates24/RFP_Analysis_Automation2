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

    def add_small_bar_chart(slide, scenario_dict, scenario_position):
        """
        We place a small bar chart at the bottom of each scenario column. 
        Here, bar_chart_values might look like [ total_baseline_savings, total_current_savings ]
        with bar_chart_labels e.g. [ "Baseline", "Current" ].
        """
        bar_vals = scenario_dict.get("bar_chart_values", [])
        bar_labels = scenario_dict.get("bar_chart_labels", [])
        if not bar_vals or not bar_labels:
            return

        chart_data = CategoryChartData()
        scenario_nm = scenario_dict.get("scenario_name", "Scenario")
        chart_data.categories = [scenario_nm]

        # Add one series per label => 2 bars per scenario
        for label, val in zip(bar_labels, bar_vals):
            chart_data.add_series(label, [val])

        left = Inches(2.8) + (scenario_position - 1)*Inches(3.1)
        top  = Inches(5.7)
        width= Inches(2.5)
        height=Inches(1.2)

        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            left, top, width, height,
            chart_data
        )
        chart = chart_shape.chart

        chart.has_legend = True
        chart.legend.include_in_layout = False
        chart.legend.position = XL_LEGEND_POSITION.RIGHT

        # Hide axis lines to keep it minimal
        cat_axis = chart.category_axis
        cat_axis.visible = False
        val_axis = chart.value_axis
        val_axis.visible = False

    # ---------- MAIN LOGIC -----------
    scenarios_per_slide = 3
    total_slides = (len(scenario_data_list) + scenarios_per_slide - 1) // scenarios_per_slide
    scenario_index = 0

    for slide_idx in range(total_slides):
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        # Remove placeholders (since we do custom placements)
        for shape in list(slide.shapes):
            shape.element.getparent().remove(shape.element)

        add_header(slide, title)
        add_row_labels(slide)

        for col_idx in range(scenarios_per_slide):
            if scenario_index >= len(scenario_data_list):
                break

            scenario_dict = scenario_data_list[scenario_index]
            add_scenario_content(slide, scenario_dict, scenario_position=(col_idx+1))
            add_small_bar_chart(slide, scenario_dict, scenario_position=(col_idx+1))

            scenario_index += 1

    return prs



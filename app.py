import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="WIN-SQUARE | Drawing & Spec Extractor Engine",
    page_icon="🪟",
    layout="wide",
)

st.title("🪟 BOQ Drawing & Specification Extraction Engine")
st.write(
    "Upload BOQ / Quotation Excel file to extract window drawings, dimensions,"
    " and specifications into a formatted sheet."
)

# -------------------------------------------------------------
# File Upload & Input Section
# -------------------------------------------------------------
st.subheader("1. Upload Quotation / BOQ Excel File")
uploaded_file = st.file_uploader(
    "Upload BOQ Excel File (.xlsx)", type=["xlsx", "xlsm"]
)

st.subheader("2. Extraction Mode")
mode = st.radio(
    "Select Extraction Preference:",
    ["Extract All Items Automatically", "Filter by Custom Window Codes"],
    horizontal=True,
)

custom_codes_input = ""
if mode == "Filter by Custom Window Codes":
  custom_codes_input = st.text_input(
      "Enter Window Codes (comma separated, e.g., D1, W3, W4)",
      value="D1, D2, W1",
  )

submit_btn = st.button("🚀 Process & Generate Output Excel", type="primary")


# -------------------------------------------------------------
# Processing Engine
# -------------------------------------------------------------
def generate_measurements_from_boq(file_bytes, target_codes=None):
  wb_source = openpyxl.load_workbook(file_bytes, data_only=True)

  wb_out = openpyxl.Workbook()
  ws_out = wb_out.active
  ws_out.title = "MEASUREMENTS"
  ws_out.views.sheetView[0].showGridLines = True

  # Styling definitions
  font_black_bold = Font(name="Calibri", size=11, bold=True, color="000000")
  font_red_bold = Font(name="Calibri", size=11, bold=True, color="FF0000")
  fill_grey = PatternFill(
      start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
  )
  fill_yellow = PatternFill(
      start_color="FFFF00", end_color="FFFF00", fill_type="solid"
  )

  thin_border = Border(
      left=Side(style="thin", color="000000"),
      right=Side(style="thin", color="000000"),
      top=Side(style="thin", color="000000"),
      bottom=Side(style="thin", color="000000"),
  )

  align_center = Alignment(horizontal="center", vertical="center")
  align_left = Alignment(horizontal="left", vertical="center")

  # Default header info extraction
  quote_ref = "WIN-QT-00000416"
  project_name = "PROJECT"
  quote_date = ""

  first_ws = wb_source[wb_source.sheetnames[0]]
  header_cell_val = str(first_ws.cell(row=6, column=1).value or "")
  if "Quote No." in header_cell_val:
    ws_out.cell(row=2, column=1, value=header_cell_val).font = font_black_bold
  else:
    ws_out.cell(
        row=2,
        column=1,
        value=(
            f"Quote No. : {quote_ref} / Project : {project_name} / Date :"
            f" {quote_date}"
        ),
    ).font = font_black_bold

  curr_row = 4
  extracted_count = 0

  for sheetname in wb_source.sheetnames:
    src_ws = wb_source[sheetname]
    images = getattr(src_ws, "_images", [])

    for r in range(1, src_ws.max_row + 1):
      c1_val = str(src_ws.cell(r, 1).value or "").strip()

      if c1_val == "Code :":
        code = str(src_ws.cell(r, 5).value or "").strip()

        # Check target filter if custom codes mode selected
        if target_codes and code.upper() not in target_codes:
          continue

        size_str = str(
            src_ws.cell(r, 12).value or src_ws.cell(r, 13).value or ""
        )
        name = str(src_ws.cell(r + 1, 5).value or "").strip()
        prof_sys = str(
            src_ws.cell(r + 1, 12).value or src_ws.cell(r + 1, 13).value or ""
        ).strip()
        location = str(src_ws.cell(r + 2, 5).value or "").strip()
        glass = str(
            src_ws.cell(r + 2, 12).value or src_ws.cell(r + 2, 13).value or ""
        ).strip()

        # Extract Width & Height
        w_val, h_val = 0, 0
        if "W =" in size_str and ";" in size_str:
          try:
            parts = size_str.split(";")
            w_val = float(parts[0].split("=")[1].strip())
            h_val = float(parts[1].split("=")[1].strip())
          except Exception:
            pass

        # Profile, Accessories, & Remarks Descriptions
        prof_desc = ""
        acc_desc = ""
        remarks_desc = ""

        for offset in range(10, 22):
          val_8 = str(src_ws.cell(r + offset, 8).value or "")
          val_13 = str(
              src_ws.cell(r + offset, 13).value
              or src_ws.cell(r + offset, 14).value
              or ""
          )
          if any(
              k in val_8
              for k in [
                  "Outer :",
                  "Interlock",
                  "Casement",
                  "Sld",
                  "Flymesh",
                  "Mullion",
              ]
          ):
            prof_desc = val_8
          if any(
              k in val_13
              for k in ["Roller :", "Handle", "Friction", "Lockable", "Locking"]
          ):
            acc_desc = val_13

        for search_r in range(r, min(r + 25, src_ws.max_row + 1)):
          cell_rem = str(src_ws.cell(search_r, 8).value or "")
          if "Remarks :" in cell_rem:
            remarks_desc = cell_rem.replace("Remarks :", "").strip()

        # Image Extraction (Matching anchored row position)
        matched_img = None
        for img in images:
          anchor = img.anchor
          if hasattr(anchor, "_from"):
            img_row = anchor._from.row + 1
            if r <= img_row <= r + 22 and img_row > 5:
              matched_img = img
              break

        # -----------------------------------------------------
        # WRITE BLOCK TO OUTPUT SHEET
        # -----------------------------------------------------
        # Code, Name, Location
        ws_out.cell(row=curr_row, column=1, value="Code :").font = (
            font_black_bold
        )
        ws_out.merge_cells(
            start_row=curr_row, start_column=2, end_row=curr_row, end_column=4
        )
        ws_out.cell(row=curr_row, column=2, value=code)

        ws_out.cell(row=curr_row + 1, column=1, value="Name :").font = (
            font_black_bold
        )
        ws_out.merge_cells(
            start_row=curr_row + 1,
            start_column=2,
            end_row=curr_row + 1,
            end_column=4,
        )
        ws_out.cell(row=curr_row + 1, column=2, value=name)

        ws_out.cell(row=curr_row + 2, column=1, value="Location :").font = (
            font_black_bold
        )
        ws_out.merge_cells(
            start_row=curr_row + 2,
            start_column=2,
            end_row=curr_row + 2,
            end_column=4,
        )
        ws_out.cell(row=curr_row + 2, column=2, value=location)

        # Size, Profile System, Glass
        ws_out.cell(row=curr_row, column=5, value="Size :").font = (
            font_black_bold
        )
        ws_out.cell(
            row=curr_row + 1, column=5, value="Profile System :"
        ).font = font_black_bold
        ws_out.cell(row=curr_row + 2, column=5, value="Glass :").font = (
            font_black_bold
        )

        ws_out.merge_cells(
            start_row=curr_row + 1,
            start_column=8,
            end_row=curr_row + 1,
            end_column=11,
        )
        ws_out.cell(row=curr_row + 1, column=8, value=prof_sys)

        ws_out.merge_cells(
            start_row=curr_row + 2,
            start_column=8,
            end_row=curr_row + 2,
            end_column=11,
        )
        ws_out.cell(row=curr_row + 2, column=8, value=glass)

        for r_offset in range(3):
          for c_idx in range(1, 12):
            ws_out.cell(
                row=curr_row + r_offset, column=c_idx
            ).border = thin_border

        # Dimensions Box (WIDTH, HEIGHT, SQFT)
        ws_out.merge_cells(
            start_row=curr_row + 5,
            start_column=5,
            end_row=curr_row + 5,
            end_column=7,
        )
        ws_out.merge_cells(
            start_row=curr_row + 6,
            start_column=5,
            end_row=curr_row + 6,
            end_column=7,
        )
        ws_out.merge_cells(
            start_row=curr_row + 7,
            start_column=5,
            end_row=curr_row + 7,
            end_column=7,
        )

        c_w_lbl = ws_out.cell(row=curr_row + 5, column=5, value="WIDTH")
        c_h_lbl = ws_out.cell(row=curr_row + 6, column=5, value="HEIGHT")
        c_s_lbl = ws_out.cell(row=curr_row + 7, column=5, value="SQFT")

        for cl in [c_w_lbl, c_h_lbl, c_s_lbl]:
          cl.font = font_black_bold
          cl.fill = fill_grey

        ws_out.merge_cells(
            start_row=curr_row + 5,
            start_column=8,
            end_row=curr_row + 5,
            end_column=9,
        )
        ws_out.merge_cells(
            start_row=curr_row + 6,
            start_column=8,
            end_row=curr_row + 6,
            end_column=9,
        )
        ws_out.merge_cells(
            start_row=curr_row + 7,
            start_column=8,
            end_row=curr_row + 7,
            end_column=9,
        )

        cw = ws_out.cell(row=curr_row + 5, column=8, value=w_val)
        cw.number_format = "0"
        ch = ws_out.cell(row=curr_row + 6, column=8, value=h_val)
        ch.number_format = "0"

        r_w = curr_row + 5
        r_h = curr_row + 6
        c_sq = ws_out.cell(
            row=curr_row + 7,
            column=8,
            value=f"=ROUND((H{r_w}/304.8)*(H{r_h}/304.8), 2)",
        )
        c_sq.number_format = "0.00"

        for r_offset in range(5, 8):
          for c_idx in range(5, 10):
            ws_out.cell(
                row=curr_row + r_offset, column=c_idx
            ).border = thin_border

        # Profile & Accessories Specs Table
        ws_out.merge_cells(
            start_row=curr_row + 11,
            start_column=5,
            end_row=curr_row + 11,
            end_column=7,
        )
        ws_out.merge_cells(
            start_row=curr_row + 11,
            start_column=8,
            end_row=curr_row + 11,
            end_column=9,
        )

        ws_out.cell(row=curr_row + 11, column=5, value="Profile").font = (
            font_black_bold
        )
        ws_out.cell(row=curr_row + 11, column=5).fill = fill_grey
        ws_out.cell(row=curr_row + 11, column=8, value="Accessories").font = (
            font_black_bold
        )
        ws_out.cell(row=curr_row + 11, column=8).fill = fill_grey

        ws_out.cell(row=curr_row + 12, column=5, value=prof_desc)
        ws_out.cell(row=curr_row + 12, column=8, value=acc_desc)

        # Remarks Block
        ws_out.cell(row=curr_row + 20, column=5, value="Remarks :").font = (
            font_red_bold
        )
        ws_out.cell(row=curr_row + 20, column=5).fill = fill_yellow

        c_rem = ws_out.cell(row=curr_row + 21, column=5, value=remarks_desc)
        c_rem.fill = fill_yellow
        c_rem.font = font_red_bold

        # Add Drawing Image safely via BytesIO Stream
        if matched_img:
          try:
            img_bytes = io.BytesIO(matched_img._data())
            img_copy = openpyxl.drawing.image.Image(img_bytes)
            ws_out.add_image(img_copy, f"A{curr_row + 5}")
          except Exception:
            pass

        curr_row += 25
        extracted_count += 1

  # Column Width Adjustments
  col_widths = {
      "A": 12,
      "B": 18,
      "C": 18,
      "D": 18,
      "E": 16,
      "F": 12,
      "G": 12,
      "H": 25,
      "I": 25,
      "J": 15,
      "K": 15,
  }
  for col_letter, width in col_widths.items():
    ws_out.column_dimensions[col_letter].width = width

  output = io.BytesIO()
  wb_out.save(output)
  output.seek(0)
  return output, extracted_count


# -------------------------------------------------------------
# Trigger Action
# -------------------------------------------------------------
if submit_btn:
  if not uploaded_file:
    st.error("Please upload a Quotation / BOQ Excel file first!")
  else:
    with st.spinner("Extracting Drawings & Building Output Sheet..."):
      target_codes = None
      if mode == "Filter by Custom Window Codes" and custom_codes_input.strip():
        target_codes = [
            c.strip().upper()
            for c in custom_codes_input.split(",")
            if c.strip()
        ]

      file_bytes = io.BytesIO(uploaded_file.read())
      uploaded_file.seek(0)

      excel_out, count = generate_measurements_from_boq(
          file_bytes, target_codes
      )

      if count > 0:
        st.success(f"Successfully extracted {count} Window/Door Block(s)!")
        filename = f"MEASUREMENTS_{uploaded_file.name}"
        st.download_button(
            label="📥 Download Generated Excel File",
            data=excel_out,
            file_name=filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning("No matching Window Codes or Drawing Blocks found.")

import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="WIN-SQUARE | Measurement Sheet Generator",
    page_icon="🪟",
    layout="wide",
)

st.title("🪟 BOQ Horizontal Measurement Sheet Engine")
st.write(
    "Upload Quotation / BOQ Excel file to convert it into the exact horizontal"
    " MEASUREMENT table layout with drawings."
)

# -------------------------------------------------------------
# 1. File Upload
# -------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Quotation / BOQ Excel File (.xlsx)", type=["xlsx", "xlsm"]
)
submit_btn = st.button("🚀 Generate MEASUREMENT Sheet", type="primary")


# -------------------------------------------------------------
# 2. Horizontal Measurement Generator Function
# -------------------------------------------------------------
def generate_horizontal_measurement_sheet(file_bytes):
  wb_source = openpyxl.load_workbook(file_bytes, data_only=True)

  wb_out = openpyxl.Workbook()
  ws_out = wb_out.active
  ws_out.title = "MEASUREMENT"
  ws_out.views.sheetView[0].showGridLines = True

  # Styles
  header_fill = PatternFill(
      start_color="A6A085", end_color="A6A085", fill_type="solid"
  )  # Olive/Grey fill as in screenshot
  font_header = Font(name="Arial", size=10, bold=True, color="000000")
  font_data = Font(name="Arial", size=10, bold=False, color="000000")
  font_bold = Font(name="Arial", size=10, bold=True, color="000000")

  thin_border = Border(
      left=Side(style="thin", color="000000"),
      right=Side(style="thin", color="000000"),
      top=Side(style="thin", color="000000"),
      bottom=Side(style="thin", color="000000"),
  )

  align_center = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )
  align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

  # ---------------------------------------------------------
  # Create Table Headers (Rows 1 & 2)
  # ---------------------------------------------------------
  ws_out.row_dimensions[1].height = 28
  ws_out.row_dimensions[2].height = 20

  headers_m = [
      ("A1:A2", "S.NO"),
      ("B1:B2", "FLAT NO"),
      ("C1:C2", "WINDOW CODE"),
      ("D1:D2", "WINDOW TYPE"),
      ("E1:E2", "LOCATION"),
      ("F1:F2", "PROFILE"),
      ("G1:H1", "Dimension in mm"),
      ("I1:I2", "Area in Sq. Ft"),
      ("J1:J2", "Mosquito Mesh"),
      ("K1:K2", "Lock Type"),
      ("L1:M1", "Glass / FLYMESH"),
      ("N1:N2", "Remarks"),
  ]

  for cell_range, text in headers_m:
    if ":" in cell_range:
      ws_out.merge_cells(cell_range)
      top_left_cell = ws_out[cell_range.split(":")[0]]
      top_left_cell.value = text
    else:
      ws_out[cell_range].value = text

  # Sub-headers in Row 2
  ws_out["G2"] = "Width"
  ws_out["H2"] = "Height"
  ws_out["L2"] = "Thickness"
  ws_out["M2"] = "Spece."

  # Format Headers
  for r in range(1, 3):
    for c in range(1, 15):
      cell = ws_out.cell(row=r, column=c)
      cell.fill = header_fill
      cell.font = font_header
      cell.alignment = align_center
      cell.border = thin_border

  # ---------------------------------------------------------
  # Extract Data From Source BOQ
  # ---------------------------------------------------------
  curr_row = 3
  s_no = 1

  for sheetname in wb_source.sheetnames:
    src_ws = wb_source[sheetname]
    images = getattr(src_ws, "_images", [])

    for r in range(1, src_ws.max_row + 1):
      c1_val = str(src_ws.cell(r, 1).value or "").strip()

      if c1_val == "Code :":
        code = str(src_ws.cell(r, 5).value or "").strip()
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
        w_val, h_val = "", ""
        if "W =" in size_str and ";" in size_str:
          try:
            parts = size_str.split(";")
            w_val = float(parts[0].split("=")[1].strip())
            h_val = float(parts[1].split("=")[1].strip())
          except Exception:
            pass

        # Extract Mesh, Lock, Thickness, Spec, Remarks
        mesh_type = "NO"
        lock_type = ""
        glass_thick = ""
        glass_spec = glass
        remarks_desc = ""

        # Parse details
        for offset in range(10, 22):
          v_mesh = str(src_ws.cell(r + offset, 8).value or "")
          if "MeshType :" in v_mesh:
            mesh_type = v_mesh.replace("MeshType :", "").strip() or "NO"

          v_lock = str(
              src_ws.cell(r + offset, 13).value
              or src_ws.cell(r + offset, 14).value
              or ""
          )
          if "Locking :" in v_lock:
            lock_type = v_lock.replace("Locking :", "").strip()

        for search_r in range(r, min(r + 25, src_ws.max_row + 1)):
          cell_rem = str(src_ws.cell(search_r, 8).value or "")
          if "Remarks :" in cell_rem:
            remarks_desc = cell_rem.replace("Remarks :", "").strip()

        # Parse Glass Thickness (e.g., 5 MM, 11.52mm)
        if glass:
          thick_match = re.search(
              r"(\d+(\.\d+)?\s*(mm|MM)?)", glass, re.IGNORECASE
          )
          if thick_match:
            glass_thick = thick_match.group(0)

        # Drawing Image Matching
        matched_img = None
        for img in images:
          anchor = img.anchor
          if hasattr(anchor, "_from"):
            img_row = anchor._from.row + 1
            if r <= img_row <= r + 22 and img_row > 5:
              matched_img = img
              break

        # -----------------------------------------------------
        # Write Data Row
        # -----------------------------------------------------
        ws_out.row_dimensions[curr_row].height = 24

        ws_out.cell(row=curr_row, column=1, value=s_no)  # S.NO
        ws_out.cell(row=curr_row, column=2, value="")  # FLAT NO
        ws_out.cell(row=curr_row, column=3, value=code)  # WINDOW CODE
        ws_out.cell(row=curr_row, column=4, value=name)  # WINDOW TYPE
        ws_out.cell(row=curr_row, column=5, value=location)  # LOCATION
        ws_out.cell(row=curr_row, column=6, value=prof_sys)  # PROFILE
        ws_out.cell(row=curr_row, column=7, value=w_val)  # Width
        ws_out.cell(row=curr_row, column=8, value=h_val)  # Height

        # Area in Sq. Ft Formula
        c_sq = ws_out.cell(
            row=curr_row,
            column=9,
            value=f"=ROUND((G{curr_row}/304.8)*(H{curr_row}/304.8), 2)",
        )
        c_sq.number_format = "0.00"

        ws_out.cell(row=curr_row, column=10, value=mesh_type)  # Mosquito Mesh
        ws_out.cell(row=curr_row, column=11, value=lock_type)  # Lock Type
        ws_out.cell(row=curr_row, column=12, value=glass_thick)  # Thickness
        ws_out.cell(row=curr_row, column=13, value=glass_spec)  # Spece.
        ws_out.cell(row=curr_row, column=14, value=remarks_desc)  # Remarks

        # Format Row Cells
        for c in range(1, 15):
          cell = ws_out.cell(row=curr_row, column=c)
          cell.font = font_data
          cell.border = thin_border
          if c in [1, 2, 3, 7, 8, 9, 10, 11, 12]:
            cell.alignment = align_center
          else:
            cell.alignment = align_left

        # Add Drawing Image in the rows below
        if matched_img:
          try:
            img_bytes = io.BytesIO(matched_img._data())
            img_copy = openpyxl.drawing.image.Image(img_bytes)
            ws_out.add_image(img_copy, f"C{curr_row + 2}")
          except Exception:
            pass

        # Give gap for drawing image height (approx 16 rows)
        curr_row += 18
        s_no += 1

  # Column Widths Adjustment
  col_widths = {
      "A": 6,
      "B": 10,
      "C": 16,
      "D": 22,
      "E": 18,
      "F": 18,
      "G": 10,
      "H": 10,
      "I": 14,
      "J": 15,
      "K": 15,
      "L": 12,
      "M": 25,
      "N": 25,
  }
  for col_letter, width in col_widths.items():
    ws_out.column_dimensions[col_letter].width = width

  output = io.BytesIO()
  wb_out.save(output)
  output.seek(0)
  return output, s_no - 1


# -------------------------------------------------------------
# 3. Execution
# -------------------------------------------------------------
if submit_btn:
  if not uploaded_file:
    st.error("Please upload a Quotation / BOQ Excel file first!")
  else:
    with st.spinner("Converting into Horizontal MEASUREMENT Sheet..."):
      file_bytes = io.BytesIO(uploaded_file.read())
      uploaded_file.seek(0)

      excel_out, count = generate_horizontal_measurement_sheet(file_bytes)

      if count > 0:
        st.success(f"Successfully generated MEASUREMENT sheet with {count} items!")
        out_filename = f"MEASUREMENT_{uploaded_file.name}"
        st.download_button(
            label="📥 Download Generated MEASUREMENT Sheet",
            data=excel_out,
            file_name=out_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning("No valid items found in the uploaded file.")

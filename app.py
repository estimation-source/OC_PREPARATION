import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# PDF Text Extraction Sathi
try:
  import pypdf
except ImportError:
  pypdf = None

st.set_page_config(
    page_title="WIN-SQUARE | BOQ & PDF Measurement Generator",
    page_icon="🪟",
    layout="wide",
)

st.title("🪟 BOQ Excel & PDF Measurement Extractor")
st.write(
    "Upload Quotation Excel (.xlsx) OR PDF (.pdf) file to automatically convert"
    " into the horizontal MEASUREMENT sheet layout."
)

# -------------------------------------------------------------
# 1. File Upload Section (Excel + PDF Support)
# -------------------------------------------------------------
st.subheader("1. Upload Quotation File (.xlsx / .pdf)")
uploaded_file = st.file_uploader(
    "Upload File (.xlsx, .xlsm, .pdf)", type=["xlsx", "xlsm", "pdf"]
)
submit_btn = st.button("🚀 Generate MEASUREMENT Sheet", type="primary")


# -------------------------------------------------------------
# 2. PDF Parsing Logic
# -------------------------------------------------------------
def parse_pdf_quotation(file_bytes):
  """PDF Madhun Window Codes, Specs ani Dimensions Read Karnyasathi."""
  if pypdf is None:
    st.error(
        "PDF processing library 'pypdf' nahiye. Kripaya `pip install pypdf`"
        " kara."
    )
    return []

  reader = pypdf.PdfReader(file_bytes)
  full_text = ""
  for page in reader.pages:
    full_text += page.extract_text() + "\n"

  extracted_items = []

  # Regex pattern to find Window Codes & Details from PDF text
  code_blocks = re.split(r"(Code\s*:\s*)", full_text, flags=re.IGNORECASE)

  for i in range(1, len(code_blocks), 2):
    block_text = code_blocks[i] + code_blocks[i + 1]

    code_match = re.search(r"Code\s*:\s*([^\n\r]+)", block_text)
    name_match = re.search(r"Name\s*:\s*([^\n\r]+)", block_text)
    loc_match = re.search(r"Location\s*:\s*([^\n\r]+)", block_text)
    size_match = re.search(
        r"Size\s*:\s*W\s*=\s*([\d\.]+);\s*H\s*=\s*([\d\.]+)", block_text
    )
    prof_match = re.search(r"Profile System\s*:\s*([^\n\r]+)", block_text)
    glass_match = re.search(r"Glass\s*:\s*([^\n\r]+)", block_text)

    code = code_match.group(1).strip() if code_match else ""
    name = name_match.group(1).strip() if name_match else ""
    location = loc_match.group(1).strip() if loc_match else ""
    w_val = float(size_match.group(1)) if size_match else ""
    h_val = float(size_match.group(2)) if size_match else ""
    prof_sys = prof_match.group(1).strip() if prof_match else ""
    glass = glass_match.group(1).strip() if glass_match else ""

    if code:
      extracted_items.append({
          "code": code,
          "name": name,
          "location": location,
          "w_val": w_val,
          "h_val": h_val,
          "prof_sys": prof_sys,
          "glass": glass,
          "mesh": "NO",
          "lock": "MULTI POINT",
          "remarks": "",
          "image": None,
      })

  return extracted_items


# -------------------------------------------------------------
# 3. Excel Processing Logic
# -------------------------------------------------------------
def parse_excel_quotation(file_bytes):
  """Excel WorkBook Madhun Data ani Drawing Images Extract Karnyasathi."""
  wb_source = openpyxl.load_workbook(file_bytes, data_only=True)
  extracted_items = []

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

        w_val, h_val = "", ""
        if "W =" in size_str and ";" in size_str:
          try:
            parts = size_str.split(";")
            w_val = float(parts[0].split("=")[1].strip())
            h_val = float(parts[1].split("=")[1].strip())
          except Exception:
            pass

        mesh_type, lock_type, remarks_desc = "NO", "", ""

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

        matched_img = None
        for img in images:
          anchor = img.anchor
          if hasattr(anchor, "_from"):
            img_row = anchor._from.row + 1
            if r <= img_row <= r + 22 and img_row > 5:
              matched_img = img
              break

        extracted_items.append({
            "code": code,
            "name": name,
            "location": location,
            "w_val": w_val,
            "h_val": h_val,
            "prof_sys": prof_sys,
            "glass": glass,
            "mesh": mesh_type,
            "lock": lock_type,
            "remarks": remarks_desc,
            "image": matched_img,
        })

  return extracted_items


# -------------------------------------------------------------
# 4. Final Sheet Builder (Creates MEASUREMENT Sheet)
# -------------------------------------------------------------
def build_measurement_excel(items):
  wb_out = openpyxl.Workbook()
  ws_out = wb_out.active
  ws_out.title = "MEASUREMENT"
  ws_out.views.sheetView[0].showGridLines = True

  header_fill = PatternFill(
      start_color="A6A085", end_color="A6A085", fill_type="solid"
  )
  font_header = Font(name="Arial", size=10, bold=True, color="000000")
  font_data = Font(name="Arial", size=10, bold=False, color="000000")

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

  # Headers
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
      ws_out[cell_range.split(":")[0]].value = text
    else:
      ws_out[cell_range].value = text

  ws_out["G2"] = "Width"
  ws_out["H2"] = "Height"
  ws_out["L2"] = "Thickness"
  ws_out["M2"] = "Spece."

  for r in range(1, 3):
    for c in range(1, 15):
      cell = ws_out.cell(row=r, column=c)
      cell.fill = header_fill
      cell.font = font_header
      cell.alignment = align_center
      cell.border = thin_border

  curr_row = 3
  s_no = 1

  for item in items:
    ws_out.row_dimensions[curr_row].height = 24

    ws_out.cell(row=curr_row, column=1, value=s_no)
    ws_out.cell(row=curr_row, column=2, value="")
    ws_out.cell(row=curr_row, column=3, value=item["code"])
    ws_out.cell(row=curr_row, column=4, value=item["name"])
    ws_out.cell(row=curr_row, column=5, value=item["location"])
    ws_out.cell(row=curr_row, column=6, value=item["prof_sys"])
    ws_out.cell(row=curr_row, column=7, value=item["w_val"])
    ws_out.cell(row=curr_row, column=8, value=item["h_val"])

    # Formula
    c_sq = ws_out.cell(
        row=curr_row,
        column=9,
        value=f"=ROUND((G{curr_row}/304.8)*(H{curr_row}/304.8), 2)",
    )
    c_sq.number_format = "0.00"

    # Thickness regex parse
    glass_thick = ""
    if item["glass"]:
      m_thick = re.search(
          r"(\d+(\.\d+)?\s*(mm|MM)?)", item["glass"], re.IGNORECASE
      )
      if m_thick:
        glass_thick = m_thick.group(0)

    ws_out.cell(row=curr_row, column=10, value=item["mesh"])
    ws_out.cell(row=curr_row, column=11, value=item["lock"])
    ws_out.cell(row=curr_row, column=12, value=glass_thick)
    ws_out.cell(row=curr_row, column=13, value=item["glass"])
    ws_out.cell(row=curr_row, column=14, value=item["remarks"])

    for c in range(1, 15):
      cell = ws_out.cell(row=curr_row, column=c)
      cell.font = font_data
      cell.border = thin_border
      if c in [1, 2, 3, 7, 8, 9, 10, 11, 12]:
        cell.alignment = align_center
      else:
        cell.alignment = align_left

    # Drawing Image Add kara (fakta Excel upload sathi)
    if item["image"]:
      try:
        img_bytes = io.BytesIO(item["image"]._data())
        img_copy = openpyxl.drawing.image.Image(img_bytes)
        ws_out.add_image(img_copy, f"C{curr_row + 2}")
      except Exception:
        pass

    curr_row += 18
    s_no += 1

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
  return output, len(items)


# -------------------------------------------------------------
# 5. Process File
# -------------------------------------------------------------
if submit_btn:
  if not uploaded_file:
    st.error("Please upload an Excel or PDF Quotation file first!")
  else:
    file_bytes = io.BytesIO(uploaded_file.read())
    uploaded_file.seek(0)

    with st.spinner("Processing file & extracting items..."):
      if uploaded_file.name.endswith(".pdf"):
        extracted_items = parse_pdf_quotation(file_bytes)
      else:
        extracted_items = parse_excel_quotation(file_bytes)

      if extracted_items:
        excel_out, count = build_measurement_excel(extracted_items)
        st.success(
            f"Successfully processed {uploaded_file.name} ({count} items"
            " extracted)!"
        )

        out_filename = f"MEASUREMENT_{uploaded_file.name.rsplit('.', 1)[0]}.xlsx"
        st.download_button(
            label="📥 Download MEASUREMENT Excel Sheet",
            data=excel_out,
            file_name=out_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning("No valid Window items found in the uploaded file.")

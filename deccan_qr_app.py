#!/usr/bin/env python3
"""
Deccan Enterprises — QR Generator (Web App)
===========================================
A simple browser front-end for the QR generator. The end user never
touches Python or a command line: they type Job Card + MR + Qty, can add
more rows with a button, click Generate, preview the QR codes, and
download everything as a ZIP plus an Excel manifest.

Same output as the command-line script:
  * One QR PNG per insulator, Insulator ID printed below it.
  * QR payload = JobCard|MR|Serial  (a scan returns Insulator ID + MR).
  * Excel manifest with columns: JobCard, MRNo, InsulatorID.
  * Folder layout inside the ZIP:
        JA266-009/
            R126605-005/JA266-009-00001.png ...
            R126605-006/...
            JA266-009_manifest.csv
            JA266-009_manifest.xlsx

RUN (after install):
    streamlit run deccan_qr_app.py

INSTALL ONCE:
    py -3 -m pip install streamlit qrcode pillow openpyxl
"""

import io
import os
import csv
import zipfile
from datetime import datetime

import streamlit as st
import qrcode
from qrcode.constants import ERROR_CORRECT_Q
from PIL import Image, ImageDraw, ImageFont

# ===== CAPTION SIZE — the ONLY knob for how big the ID text is =========
CAPTION_FONT_SIZE = 24
# =======================================================================


# ----------------------------------------------------------------------
# QR + caption logic (identical to the command-line generator)
# ----------------------------------------------------------------------
def _load_font(size):
    candidates = [
        "DejaVuSans-Bold.ttf",
        "arialbd.ttf",
        r"C:\Windows\Fonts\arialbd.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def make_qr_image(data, box_size=20, border=2):
    qr = qrcode.QRCode(version=None, error_correction=ERROR_CORRECT_Q,
                       box_size=box_size, border=border)
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def label_with_caption(code_img, caption, target_px=320, font_size=CAPTION_FONT_SIZE):
    code_img = code_img.resize((target_px, target_px), Image.NEAREST)
    font = _load_font(font_size)

    measure = ImageDraw.Draw(Image.new("RGB", (10, 10), "white"))
    bbox = measure.textbbox((0, 0), caption, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    top_off = bbox[1]

    pad_above = 16
    pad_below = 20
    cap_h = text_h + pad_above + pad_below
    canvas_w = max(target_px, text_w + 8)

    canvas = Image.new("RGB", (canvas_w, target_px + cap_h), "white")
    qr_x = (canvas_w - target_px) // 2
    canvas.paste(code_img, (qr_x, 0))
    draw = ImageDraw.Draw(canvas)
    draw.text(((canvas_w - text_w) / 2, target_px + pad_above - top_off),
              caption, fill="black", font=font)
    return canvas


def build_outputs(jobcard, mr_jobs, start=1):
    """Generate all QR images + manifests in memory.

    Returns: (rows, images, xlsx_bytes, csv_bytes, zip_bytes)
      rows   : list of dicts (InsulatorID, JobCard, MRNo, QRPayload)
      images : list of (mr, insulator_id, PIL.Image) for preview
    """
    jobcard = jobcard.strip().upper()
    rows = []
    images = []
    n = start

    for mr, qty in mr_jobs:
        mr = mr.strip().upper()
        for _ in range(qty):
            serial = f"{jobcard}-{n:05d}"
            payload = f"{jobcard}|{mr}|{serial}"           # ID + MR in the QR
            img = label_with_caption(make_qr_image(payload), serial)
            images.append((mr, serial, img))
            rows.append({
                "InsulatorID": serial, "JobCard": jobcard, "MRNo": mr,
                "QRPayload": payload,
                "GeneratedOn": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            n += 1

    # --- Excel manifest (JobCard / MRNo / InsulatorID) ---
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "Manifest"
    ws.append(["JobCard", "MRNo", "InsulatorID"])
    for r in rows:
        ws.append([r["JobCard"], r["MRNo"], r["InsulatorID"]])
    for col_letter, width in [("A", 14), ("B", 18), ("C", 22)]:
        ws.column_dimensions[col_letter].width = width
    xlsx_buf = io.BytesIO()
    wb.save(xlsx_buf)
    xlsx_bytes = xlsx_buf.getvalue()

    # --- CSV manifest (full columns, like the script) ---
    csv_buf = io.StringIO()
    w = csv.DictWriter(csv_buf, fieldnames=list(rows[0].keys()))
    w.writeheader(); w.writerows(rows)
    csv_bytes = csv_buf.getvalue().encode("utf-8")

    # --- ZIP with the JobCard/MR/image.png folder layout ---
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for mr, serial, img in images:
            png_buf = io.BytesIO()
            img.save(png_buf, format="PNG")
            zf.writestr(f"{jobcard}/{mr}/{serial}.png", png_buf.getvalue())
        zf.writestr(f"{jobcard}/{jobcard}_manifest.xlsx", xlsx_bytes)
        zf.writestr(f"{jobcard}/{jobcard}_manifest.csv", csv_bytes)
    zip_bytes = zip_buf.getvalue()

    return rows, images, xlsx_bytes, csv_bytes, zip_bytes


# ----------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------
st.set_page_config(page_title="Deccan QR Generator", page_icon="🔳", layout="centered")

st.title("Deccan Enterprises — QR Generator")
st.caption("Enter a Job Card, then one or more MR numbers with quantities. "
           "Click Generate to create the QR codes and the Excel manifest.")

# --- Job card ---
jobcard = st.text_input("Job Card No.", placeholder="e.g. JA266-009").strip()

st.markdown("##### MR numbers & quantities")

# Dynamic MR rows held in session state
if "mr_rows" not in st.session_state:
    st.session_state.mr_rows = [{"mr": "", "qty": 10}]


def add_row():
    st.session_state.mr_rows.append({"mr": "", "qty": 10})


def remove_row(i):
    if len(st.session_state.mr_rows) > 1:
        st.session_state.mr_rows.pop(i)


# Render each MR row
for i, row in enumerate(st.session_state.mr_rows):
    c1, c2, c3 = st.columns([5, 2, 1])
    st.session_state.mr_rows[i]["mr"] = c1.text_input(
        "MR No.", value=row["mr"], key=f"mr_{i}",
        placeholder="e.g. R126605-005", label_visibility="collapsed" if i else "visible",
    )
    st.session_state.mr_rows[i]["qty"] = c2.number_input(
        "Qty", min_value=1, max_value=100000, value=int(row["qty"]), step=1,
        key=f"qty_{i}", label_visibility="collapsed" if i else "visible",
    )
    # remove button (disabled when only one row remains)
    if i == 0:
        c3.markdown("&nbsp;")  # spacer to align with the labelled row
    c3.button("🗑", key=f"del_{i}", on_click=remove_row, args=(i,),
              disabled=(len(st.session_state.mr_rows) == 1),
              help="Remove this row")

st.button("➕ Add another MR", on_click=add_row)

# --- starting serial (optional, advanced) ---
with st.expander("Advanced options"):
    start_serial = st.number_input(
        "Starting serial number", min_value=1, value=1, step=1,
        help="The first insulator serial. Leave at 1 unless continuing a previous batch.",
    )

st.divider()

# --- Generate ---
if st.button("Generate QR codes", type="primary", use_container_width=True):
    # validate
    if not jobcard:
        st.error("Please enter a Job Card No.")
        st.stop()

    mr_jobs = []
    for row in st.session_state.mr_rows:
        mr = (row["mr"] or "").strip()
        qty = int(row["qty"])
        if mr and qty > 0:
            mr_jobs.append((mr, qty))

    if not mr_jobs:
        st.error("Please enter at least one MR No. with a quantity.")
        st.stop()

    total = sum(q for _, q in mr_jobs)
    with st.spinner(f"Generating {total} QR code(s)…"):
        rows, images, xlsx_bytes, csv_bytes, zip_bytes = build_outputs(
            jobcard, mr_jobs, start=int(start_serial)
        )

    st.success(f"Generated {len(rows)} QR code(s) across {len(mr_jobs)} MR number(s).")

    # store results so download buttons don't trigger a regenerate
    st.session_state.result = {
        "jobcard": jobcard.upper(),
        "rows": rows, "images": images,
        "xlsx": xlsx_bytes, "csv": csv_bytes, "zip": zip_bytes,
    }

# --- Show results + downloads (persist across download clicks) ---
res = st.session_state.get("result")
if res:
    jobcard = res["jobcard"]

    st.markdown("### Download")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇ All QR codes + manifest (ZIP)",
        data=res["zip"],
        file_name=f"{jobcard}_QR_codes.zip",
        mime="application/zip",
        use_container_width=True,
    )
    d2.download_button(
        "⬇ Excel manifest only (.xlsx)",
        data=res["xlsx"],
        file_name=f"{jobcard}_manifest.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("### Manifest preview")
    st.dataframe(
        [{"JobCard": r["JobCard"], "MRNo": r["MRNo"], "InsulatorID": r["InsulatorID"]}
         for r in res["rows"]],
        use_container_width=True, hide_index=True,
    )

    st.markdown("### QR preview")
    st.caption("Showing the first few QR codes. Download the ZIP for all of them.")
    preview = res["images"][:6]
    cols = st.columns(3)
    for idx, (mr, serial, img) in enumerate(preview):
        buf = io.BytesIO(); img.save(buf, format="PNG")
        cols[idx % 3].image(buf.getvalue(), caption=f"{mr}", use_container_width=True)
    if len(res["images"]) > len(preview):
        st.caption(f"… and {len(res['images']) - len(preview)} more in the ZIP.")

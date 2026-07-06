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
  * Excel manifest with columns: JobCard, MRNo, SKU, InsulatorID.
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
# Kept at the proportion you approved; the whole label is rendered at a
# higher resolution (QR_TARGET_PX) so the text is crisp and clearly
# readable. To make the ID text bigger/smaller RELATIVE to the QR,
# change CAPTION_FONT_SIZE. To make the whole image bigger, raise
# QR_TARGET_PX (and CAPTION_FONT_SIZE by the same factor).
QR_TARGET_PX = 640
CAPTION_FONT_SIZE = 93   # large ID text under the QR for shop-floor readability
# =======================================================================


# ----------------------------------------------------------------------
# QR + caption logic (identical to the command-line generator)
# ----------------------------------------------------------------------
def _load_font(size):
    """Load a bold TrueType font at the requested size.

    Looks FIRST for the DejaVuSans-Bold.ttf bundled next to this script
    (so it works identically on Streamlit Cloud, where system fonts may
    be missing). Falls back to common system fonts, then — only as a last
    resort — Pillow's tiny bitmap default. If we ever hit that last
    resort, FONT_FALLBACK is set True so the app can warn the user.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "DejaVuSans-Bold.ttf"),   # bundled with the app
        "DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "arialbd.ttf",
        r"C:\Windows\Fonts\arialbd.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    global FONT_FALLBACK
    FONT_FALLBACK = True
    return ImageFont.load_default()


FONT_FALLBACK = False


def make_qr_image(data, box_size=20, border=2):
    qr = qrcode.QRCode(version=None, error_correction=ERROR_CORRECT_Q,
                       box_size=box_size, border=border)
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def label_with_caption(code_img, caption, sku="", target_px=QR_TARGET_PX, font_size=CAPTION_FONT_SIZE):
    """QR with the Insulator ID printed large below it, and (optionally)
    the SKU on a second centered line beneath the ID. The image widens
    if the text is wider than the QR, so nothing is clipped."""
    code_img = code_img.resize((target_px, target_px), Image.NEAREST)
    font = _load_font(font_size)

    measure = ImageDraw.Draw(Image.new("RGB", (10, 10), "white"))

    def dims(text):
        b = measure.textbbox((0, 0), text, font=font)
        return (b[2] - b[0], b[3] - b[1], b[1])   # width, height, top offset

    lines = [caption]
    if sku:
        lines.append(sku)

    line_dims = [dims(t) for t in lines]
    max_w = max(w for w, _, _ in line_dims)
    line_gap = int(font_size * 0.25)
    pad_above = 16
    pad_below = 22

    total_text_h = sum(h for _, h, _ in line_dims) + line_gap * (len(lines) - 1)
    cap_h = total_text_h + pad_above + pad_below
    canvas_w = max(target_px, max_w + 8)

    canvas = Image.new("RGB", (canvas_w, target_px + cap_h), "white")
    qr_x = (canvas_w - target_px) // 2
    canvas.paste(code_img, (qr_x, 0))
    draw = ImageDraw.Draw(canvas)

    y = target_px + pad_above
    for text, (w, h, top_off) in zip(lines, line_dims):
        draw.text(((canvas_w - w) / 2, y - top_off), text, fill="black", font=font)
        y += h + line_gap
    return canvas


def build_outputs(jobcard, sku, mr_jobs, start=1):
    """Generate all QR images + manifests in memory.

    One SKU applies to the whole job card (not per MR).

    Returns: (rows, images, xlsx_bytes, csv_bytes, zip_bytes)
      rows   : list of dicts (InsulatorID, JobCard, MRNo, SKU, QRPayload)
      images : list of (mr, insulator_id, PIL.Image) for preview
    """
    jobcard = jobcard.strip().upper()
    sku = (sku or "").strip()
    rows = []
    images = []
    n = start

    for mr, qty in mr_jobs:
        mr = mr.strip().upper()
        for _ in range(qty):
            serial = f"{jobcard}-{n:05d}"
            payload = f"{jobcard}|{mr}|{serial}"           # ID + MR in the QR
            img = label_with_caption(make_qr_image(payload), serial, sku=sku)
            images.append((mr, serial, img))
            rows.append({
                "InsulatorID": serial, "JobCard": jobcard, "MRNo": mr,
                "SKU": sku,
                "QRPayload": payload,
                "GeneratedOn": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            n += 1

    # --- Excel manifest (JobCard / MRNo / SKU / InsulatorID) ---
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "Manifest"
    ws.append(["JobCard", "MRNo", "SKU", "InsulatorID"])
    for r in rows:
        ws.append([r["JobCard"], r["MRNo"], r["SKU"], r["InsulatorID"]])
    for col_letter, width in [("A", 14), ("B", 18), ("C", 16), ("D", 22)]:
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

# Font self-test: render a throwaway label so _load_font runs once. If it
# had to use the tiny bitmap fallback, warn the operator immediately so a
# small-text problem is never a silent surprise.
_ = label_with_caption(make_qr_image("TEST"), "TEST")
if FONT_FALLBACK:
    st.warning(
        "The label font (DejaVuSans-Bold.ttf) was not found, so the "
        "Insulator ID text will be small. Make sure DejaVuSans-Bold.ttf "
        "is uploaded to the app's GitHub repo, next to deccan_qr_app.py."
    )

# --- Job card + SKU (one SKU per job card) ---
jc_col, sku_col = st.columns(2)
jobcard = jc_col.text_input("Job Card No.", placeholder="e.g. JA266-009").strip()
sku = sku_col.text_input("SKU No.", placeholder="e.g. 1098 A3",
                         help="One SKU applies to the whole job card.").strip()

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
    if not sku:
        st.error("Please enter the SKU No. for this job card.")
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
            jobcard, sku, mr_jobs, start=int(start_serial)
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
        [{"JobCard": r["JobCard"], "MRNo": r["MRNo"], "SKU": r["SKU"], "InsulatorID": r["InsulatorID"]}
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

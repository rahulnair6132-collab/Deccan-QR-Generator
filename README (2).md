# Deccan QR Generator — Web App

A browser front-end for the QR generator. End users type a Job Card,
add MR numbers with quantities, click **Generate**, then download a ZIP
of QR images (organised by `JobCard/MR/`) plus the Excel manifest. No
Python or command line needed once it is running/hosted.

The QR payload is `JobCard|MR|Serial`, so a scan returns the Insulator
ID **and** the MR number — same as the command-line script.

---

## Option A — Run it on one PC (simplest)

On the machine that will host it (only this machine needs Python):

```
py -3 -m pip install -r requirements.txt
py -3 -m streamlit run deccan_qr_app.py
```

A browser tab opens at `http://localhost:8501`. Anyone on the **same PC**
can use it there.

To let others on the **office network** use it from their own browsers
(no install on their side), run:

```
py -3 -m streamlit run deccan_qr_app.py --server.address 0.0.0.0
```

then share `http://<this-PC-IP>:8501` (find the IP with `ipconfig`).
Keep this PC on and the window open while others use it.

---

## Option B — Host it free on Streamlit Community Cloud (best for end users)

So nobody needs Python or your PC switched on:

1. Put `deccan_qr_app.py` and `requirements.txt` in a GitHub repo.
2. Go to https://share.streamlit.io , sign in with GitHub, click
   **New app**, pick the repo, set the main file to `deccan_qr_app.py`.
3. Deploy. You get a public URL (e.g. `https://deccan-qr.streamlit.app`)
   that anyone can open in a browser and use directly.

---

## How users use it

1. Enter the **Job Card No.** (e.g. `JA266-009`).
2. Enter an **MR No.** and **Qty**. Click **➕ Add another MR** for more.
3. Click **Generate QR codes**.
4. Download:
   - **All QR codes + manifest (ZIP)** — folder layout `JA266-009/<MR>/<ID>.png`
     plus the Excel and CSV manifests.
   - **Excel manifest only (.xlsx)** — the 3 columns: JobCard, MRNo, InsulatorID.

Serials run continuously across all MR rows in a single generate
(e.g. MR1 qty 10 → 1–10, MR2 qty 5 → 11–15).

---

## Tuning the caption size

In `deccan_qr_app.py`, near the top:

```python
CAPTION_FONT_SIZE = 24
```

Bigger number = bigger ID text under the QR. 24 matches the reference label.

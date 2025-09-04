import os
import io
import zipfile
from datetime import datetime
from dateutil.relativedelta import relativedelta

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json

import streamlit as st
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ===============================
# CONFIGURATION
# ===============================
SERVICE_ACCOUNT_FILE = r"C:\\Users\\skm\\Desktop\\CALIBRATION REPORT AUTOMATION\\data\\service_account.json"
SHEET_ID = "1jgqN9pNWVKsH2gDVCsSmkMpYiyE0_y-vqaxrqt4B8xg"
LOGO_PATH = r"NPL_LOGO.png"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ===============================
# GSPREAD AUTH (Secrets friendly)
# ===============================
if "service_account" in st.secrets:
    svc = st.secrets["service_account"]
    if isinstance(svc, str):
        service_account_info = json.loads(svc)
    else:
        service_account_info = svc
    try:
        gc = gspread.service_account_from_dict(service_account_info, scopes=SCOPES)
    except Exception:
        creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
        gc = gspread.authorize(creds)
else:
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)

# ===============================
# LOAD DATA
# ===============================
sh = gc.open_by_key(SHEET_ID)
df = pd.DataFrame(sh.worksheet("Form Responses 1").get_all_records())
instrument_list = pd.DataFrame(sh.worksheet("InstrumentList").get_all_records())
master_instrument_list = pd.DataFrame(sh.worksheet("MASTERINSTRUMENTLIST").get_all_records())

df.columns = df.columns.str.strip()
instrument_list.columns = instrument_list.columns.str.strip()
master_instrument_list.columns = master_instrument_list.columns.str.strip()

# ===============================
# HELPERS
# ===============================
def _safe_strip(series, default=""):
    try:
        return series.astype(str).str.strip()
    except Exception:
        return series.fillna(default)

for must_col in ["Instrument Tag", "Master Serial No", "Engineer Name", "Remarks"]:
    if must_col in df.columns:
        df[must_col] = _safe_strip(df[must_col], "")

timestamp_col = next((c for c in df.columns if "timestamp" in c.lower()), None)

def to_float_or_none(val):
    try:
        if val in [None, ""]:
            return None
        return float(str(val).strip())
    except Exception:
        return None

def get_calibration_date(row):
    if timestamp_col and timestamp_col in row:
        dt = pd.to_datetime(row[timestamp_col], errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime()
    return datetime.now()

def draw_dual_border(canvas, doc):
    canvas.saveState()
    w, h = A4
    outer, inner = 12, 18
    canvas.setLineWidth(1.2)
    canvas.rect(outer, outer, w - 2*outer, h - 2*outer)
    canvas.setLineWidth(0.6)
    canvas.rect(inner, inner, w - 2*inner, h - 2*inner)
    canvas.restoreState()

def fmt(v, ndigits=4):
    return "" if v is None else str(round(v, ndigits))

# ===============================
# PDF GENERATION
# ===============================
def generate_pdf(row, inst, master, logo_path=None):
    inst_tag = row.get("Instrument Tag", "")
    master_serial = row.get("Master Serial No", "")
    inst_type = str(inst.get("INST TYPE", inst.get("Type", ""))).strip().upper()

    calib_dt = get_calibration_date(row)
    due_dt = calib_dt + relativedelta(years=1)

    min_val = to_float_or_none(inst.get("Min Range", 0)) or 0.0
    max_val = to_float_or_none(inst.get("Max Range", 0)) or 0.0
    if min_val > max_val:
        min_val, max_val = max_val, min_val
    unit = inst.get("Unit", "")

    span = max_val - min_val
    desired_values_up = [round(min_val + p*span, 4) for p in [0,0.25,0.5,0.75,1.0]]
    desired_values_dn = desired_values_up[::-1]
    desired_mA = [round(4 + (val - min_val) / span * 16, 3) if span != 0 else None for val in desired_values_up]

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    normal.fontSize = 9
    wrap_style = ParagraphStyle("wrap", fontSize=9, alignment=1)
    story = []

    # Header
    logo = Image(logo_path, width=100, height=100) if logo_path and os.path.exists(logo_path) else Paragraph("", normal)
    header_table = Table(
        [[logo, Paragraph(
            "<para align='center'><b>NABHA POWER LTD.</b><br/>"
            "<b>2X 700 M.W RAJPURA SUPERCRITICAL THERMAL POWER STATION</b><br/>"
            "<b><u>CALIBRATION REPORT</u></b></para>", normal)]],
        colWidths=[110, 400]
    )
    header_table.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("BOX",(0,0),(-1,-1),1,colors.black)]))
    story.append(header_table)
    story.append(Spacer(1,8))

    # Info Table
    area, unit_str, location, service = inst.get("Area","BOILER"), inst.get("Unit:","Unit-1"), inst.get("Location","0M"), inst.get("SERVICE DESCRIPTION","")
    left_info = [
        Paragraph(f"<b>Report No:</b> {inst.get('Report No.','')}", normal),
        Paragraph(f"<b>Calibration Date:</b> {calib_dt.strftime('%d-%m-%Y')}", normal),
        Paragraph(f"<b>Calibration Due Date:</b> {due_dt.strftime('%d-%m-%Y')}", normal)
    ]
    right_info = [
        Paragraph(f"<b>Area:</b> {area}", normal),
        Paragraph(f"<b>Unit:</b> {unit_str}", normal),
        Paragraph(f"<b>Location:</b> {location}", normal),
        Paragraph(f"<b>Service:</b> {service}", normal)
    ]
    info_table = Table([[left_info, right_info]], colWidths=[255,255])
    info_table.setStyle(TableStyle([("BOX",(0,0),(-1,-1),1,colors.black)]))
    story.append(info_table)
    story.append(Spacer(1,8))

    # Instrument Details Table
    left_details = [
        Paragraph("<b>Details of Instrument Under Test</b>", normal),
        Paragraph(f"Tag No: {inst_tag}", normal),
        Paragraph(f"Inst. Make: {inst.get('Make','')}", normal),
        Paragraph(f"Model No: {inst.get('Model','')}", normal),
        Paragraph(f"Sr. No: {inst.get('Sr. No.','')}", normal),
        Paragraph(f"Range: {min_val} - {max_val} {unit}", normal),
        Paragraph(f"Type: {inst_type}", normal)
    ]
    right_details = [
        Paragraph("<b>Details of Calibration Master Instrument</b>", normal),
        Paragraph(f"Make/Inst.Type: {master.get('Make/Inst.Type','')}", normal),
        Paragraph(f"Make: {master.get('Make','')}", normal),
        Paragraph(f"Model No: {master.get('Model','')}", normal),
        Paragraph(f"Serial No: {master.get('Serial No.','')}", normal),
        Paragraph(f"Certificate No: {master.get('Certificate No.','')}", normal),
        Paragraph(f"Valid Upto: {master.get('Certificate Valid Upto','')}", normal)
    ]
    details_table = Table([[left_details, right_details]], colWidths=[255,255])
    details_table.setStyle(TableStyle([("BOX",(0,0),(-1,-1),1,colors.black)]))
    story.append(details_table)
    story.append(Spacer(1,8))

    # ============================
    # Calibration Table Logic (TX, SWITCH, GAUGE)
    # ============================

   calib_table_data, col_widths = [],[]

    if inst_type.startswith("TX"):
        headers = [
            Paragraph("SL<br/>NO", wrap_style),
            Paragraph(f"DESIRED<br/>VALUE (UP) ({unit})", wrap_style),
            Paragraph("ACTUAL<br/>VALUE (UP)", wrap_style),
            Paragraph("DESIRED<br/>mA (UP)", wrap_style),
            Paragraph("ACTUAL<br/>mA (UP)", wrap_style),
            Paragraph(f"DESIRED<br/>VALUE (DOWN) ({unit})", wrap_style),
            Paragraph("ACTUAL<br/>VALUE (DOWN)", wrap_style),
            Paragraph("DESIRED<br/>mA (DOWN)", wrap_style),
            Paragraph("ACTUAL<br/>mA (DOWN)", wrap_style),
            Paragraph("REMARKS", wrap_style),
        ]
        calib_table_data.append(headers)
        up_fields = [
            ("As Found (0%) Up", "As Found mA (0%) Up"),
            ("As Found (25%) Up", "As Found mA (25%) Up"),
            ("As Found (50%) Up", "As Found mA (50%) Up"),
            ("As Found (75%) Up", "As Found mA (75%) Up"),
            ("As Found (100%) Up", "As Found mA (100%) Up"),
        ]
        dn_fields = [
            ("As Found (0%) Down", "As Found mA (0%) Down"),
            ("As Found (25%) Down", "As Found mA (25%) Down"),
            ("As Found (50%) Down", "As Found mA (50%) Down"),
            ("As Found (75%) Down", "As Found mA (75%) Down"),
            ("As Found (100%) Down", "As Found mA (100%) Down"),
        ][::-1]

        for i in range(5):
            up_val = to_float_or_none(row.get(up_fields[i][0]))
            up_mA  = to_float_or_none(row.get(up_fields[i][1]))
            dn_val = to_float_or_none(row.get(dn_fields[i][0]))
            dn_mA  = to_float_or_none(row.get(dn_fields[i][1]))
            calib_table_data.append([
                i+1, fmt(desired_values_up[i]), fmt(up_val),
                fmt(desired_mA[i]), fmt(up_mA),
                fmt(desired_values_dn[i]), fmt(dn_val),
                fmt(desired_mA[::-1][i]), fmt(dn_mA), ""
            ])
        col_widths = [30, 60, 60, 55, 55, 60, 60, 55, 55, 55]

    elif inst_type.startswith("SWITCH"):
        headers = ["SL NO", "Switch SET-1", "Switch RESET-1", "Switch SET-2", "Switch RESET-2", "Switch SET-3", "Switch RESET-3"]
        calib_table_data.append(headers)
        calib_table_data.append([
            1,
            row.get("Switch SET-1",""), row.get("Switch RESET-1",""),
            row.get("Switch SET-2",""), row.get("Switch RESET-2",""),
            row.get("Switch SET-3",""), row.get("Switch RESET-3","")
        ])
        col_widths = [35, 75, 75, 75, 75, 75, 75]

    else:  # Gauge
        headers = [
            Paragraph("SL<br/>NO", wrap_style),
            Paragraph(f"DESIRED<br/>VALUE (UP) ({unit})", wrap_style),
            Paragraph("ACTUAL<br/>VALUE (UP)", wrap_style),
            Paragraph("%<br/>ERROR (UP)", wrap_style),
            Paragraph(f"DESIRED<br/>VALUE (DOWN) ({unit})", wrap_style),
            Paragraph("ACTUAL<br/>VALUE (DOWN)", wrap_style),
            Paragraph("%<br/>ERROR (DOWN)", wrap_style),
            Paragraph("REMARKS", wrap_style),
        ]
        calib_table_data.append(headers)
        up_down_fields = [
            ("As Found (0%) Up","As Found (0%) Down"),
            ("As Found (25%) Up","As Found (25%) Down"),
            ("As Found (50%) Up","As Found (50%) Down"),
            ("As Found (75%) Up","As Found (75%) Down"),
            ("As Found (100%) Up","As Found (100%) Down")
        ]
        down_actual_values = [to_float_or_none(row.get(field[1])) for field in up_down_fields][::-1]

        for i in range(5):
            up_val = to_float_or_none(row.get(up_down_fields[i][0]))
            dn_val = down_actual_values[i]
            err_up = ((up_val - desired_values_up[i])/desired_values_up[i]*100) if (up_val and desired_values_up[i] != 0) else None
            err_dn = ((dn_val - desired_values_dn[i])/desired_values_dn[i]*100) if (dn_val and desired_values_dn[i] != 0) else None
            calib_table_data.append([
                i+1, fmt(desired_values_up[i]), fmt(up_val), fmt(err_up),
                fmt(desired_values_dn[i]), fmt(dn_val), fmt(err_dn), ""
            ])
        col_widths = [35, 75, 70, 60, 75, 70, 60, 55]

    calib_table = Table(calib_table_data, hAlign="CENTER", colWidths=col_widths)
    calib_table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.6,colors.black),
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(calib_table); story.append(Spacer(1,12))

    # Remarks + signatures
    remarks_text = Paragraph(f"<b>Remarks:</b> {row.get('Remarks','')}", normal)
    sig_table = Table(
        [[remarks_text],
         [Paragraph(f"<b>Calibrated By:</b> {row.get('Engineer Name','')}", normal),
          Paragraph("<b>Checked By:</b> __________________", normal)]],
        colWidths=[255,255]
    )
    sig_table.setStyle(TableStyle([("BOX",(0,0),(-1,-1),1,colors.black),
                                   ("INNERGRID",(0,0),(-1,-1),0.5,colors.black),
                                   ("ALIGN",(0,1),(0,1),"LEFT"),
                                   ("ALIGN",(1,1),(1,1),"RIGHT")]))
    story.append(sig_table)

    doc.build(story, onFirstPage=draw_dual_border, onLaterPages=draw_dual_border)
    buffer.seek(0)
    return buffer


# ===============================
# STREAMLIT APP
# ===============================
st.set_page_config(page_title="Calibration Report Automation", page_icon="📑", layout="wide")
st.title("📑 Calibration Report Automation")

# --- Session state for activity log ---
if "activity_log" not in st.session_state:
    st.session_state.activity_log = []

# --- Prepare timestamp bounds / clean ---
if timestamp_col:
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce")
    # Boundaries ignoring NaT
    valid_dt = df[timestamp_col].dropna()
    if valid_dt.empty:
        min_date = max_date = datetime.now().date()
    else:
        min_date = valid_dt.min().date()
        max_date = valid_dt.max().date()
else:
    min_date = max_date = datetime.now().date()

with st.sidebar:
    st.subheader("🔎 Filters")
    st.caption(f"Available data window: **{min_date.strftime('%d-%m-%Y')}** to **{max_date.strftime('%d-%m-%Y')}**")

    # Separate Start/End pickers
    start_date = st.date_input("Start date", value=min_date, min_value=min_date, max_value=max_date)
    end_date = st.date_input("End date", value=max_date, min_value=min_date, max_value=max_date)

    # Guard: dates outside bounds (defensive; date_input already clamps)
    out_of_range = (start_date < min_date) or (end_date > max_date)
    if out_of_range:
        st.toast("Selected dates were outside available range and have been clamped.", icon="⚠️")

    # Guard: start after end
    if start_date > end_date:
        st.error("Start date cannot be after End date. Fix the dates to continue.")
        st.stop()

# Apply date mask
if timestamp_col:
    mask = (df[timestamp_col].dt.date >= start_date) & (df[timestamp_col].dt.date <= end_date)
    df_filtered = df[mask].copy()
else:
    df_filtered = df.copy()

# --- Dashboard (top cards) ---
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total records in window", len(df_filtered))
with col2:
    unique_tags = df_filtered["Instrument Tag"].astype(str).str.strip().str.upper().replace({"": None}).dropna().nunique()
    st.metric("Unique instrument tags", unique_tags)
with col3:
    # Determine type mix by joining to instrument_list on TAG (case-insensitive)
    _tmp = df_filtered.merge(
        instrument_list.assign(_TAG_UP=instrument_list["TAG"].astype(str).str.upper()),
        left_on=df_filtered["Instrument Tag"].astype(str).str.upper(),
        right_on="_TAG_UP",
        how="left"
    )
    type_counts = (
        _tmp["INST TYPE"].astype(str).str.upper().str.strip()
        .replace({"": "UNKNOWN", "NAN": "UNKNOWN"})
        .value_counts(dropna=False)
    )
    tx = int(type_counts.get("TX", 0))
    sw = int(type_counts.get("SWITCH", 0))
    gauge = int(type_counts.sum() - tx - sw)
    st.metric("Instrument mix (TX / SW / Others)", f"{tx} / {sw} / {gauge}")
with col4:
    st.metric("Selected window", f"{start_date.strftime('%d-%m-%Y')} → {end_date.strftime('%d-%m-%Y')}")

st.divider()

# --- Single report selector ---
st.subheader("🎯 Generate a Single Report")
if df_filtered.empty:
    st.info("No records found for the selected date window.")
else:
    # Build a friendly selector that uniquely identifies a row
    # Label: [index] Tag | Master SN | Timestamp
    df_sel = df_filtered.copy()
    df_sel["_idx"] = df_sel.index.astype(str)
    label_series = df_sel.apply(
        lambda r: f"[{r['_idx']}] {r.get('Instrument Tag','').strip()} | {r.get('Master Serial No','').strip()} | {r.get(timestamp_col).strftime('%d-%m-%Y %H:%M') if timestamp_col and pd.notna(r.get(timestamp_col)) else ''}",
        axis=1
    )
    options = list(zip(label_series.tolist(), df_sel.index.tolist()))
    label_to_index = {lbl: idx for lbl, idx in options}

    selected_label = st.selectbox("Pick a specific record", options=[lbl for lbl, _ in options])

    c1, c2 = st.columns([1,2])
    with c1:
        gen_single = st.button("📄 Generate Single PDF", type="primary")
    if gen_single and selected_label:
        ridx = label_to_index[selected_label]
        row = df.loc[ridx]

        # Lookup instrument + master rows (case-insensitive, same as your original)
        inst_tag = str(row.get("Instrument Tag", "")).strip()
        master_serial = str(row.get("Master Serial No", "")).strip()

        inst_rows = instrument_list[instrument_list["TAG"].astype(str).str.upper() == inst_tag.upper()]
        master_rows = master_instrument_list[master_instrument_list["Serial No."].astype(str).str.upper() == master_serial.upper()]

        if inst_rows.empty or master_rows.empty:
            st.error("Instrument/MASTER lookup failed for this record. Check TAG or Serial No. in source sheets.")
        else:
            inst, master = inst_rows.iloc[0], master_rows.iloc[0]
            pdf_buffer = generate_pdf(row, inst, master, logo_path=LOGO_PATH)
            single_filename = f"{inst_tag}_{datetime.now().strftime('%d%m%y')}.pdf"

            st.download_button(
                "⬇️ Download This Report (PDF)",
                data=pdf_buffer.getvalue(),
                file_name=single_filename,
                mime="application/pdf",
                key=f"dl_single_{ridx}"
            )
            st.session_state.activity_log.append({
                "time": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                "action": "Single PDF Generated",
                "tag": inst_tag,
                "master_serial": master_serial,
                "filename": single_filename
            })
            st.success("Single report ready to download.")

st.divider()

# --- Bulk ZIP generation ---
st.subheader("📦 Generate All Reports (ZIP)")
colz1, colz2 = st.columns([1,2])
with colz1:
    gen_all = st.button("🧾 Generate All Reports (ZIP)")

if gen_all:
    if df_filtered.empty:
        st.error("No records in the selected date range.")
    else:
        zip_buffer = io.BytesIO()
        generated = 0
        skipped = 0
        skipped_items = []

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for idx, row in df_filtered.iterrows():
                inst_tag = str(row.get("Instrument Tag", "")).strip()
                master_serial = str(row.get("Master Serial No", "")).strip()

                # Case-insensitive lookups (as in your original)
                inst_rows = instrument_list[instrument_list["TAG"].astype(str).str.upper() == inst_tag.upper()]
                master_rows = master_instrument_list[master_instrument_list["Serial No."].astype(str).str.upper() == master_serial.upper()]

                if inst_rows.empty or master_rows.empty:
                    skipped += 1
                    skipped_items.append((inst_tag, master_serial))
                    continue

                inst, master = inst_rows.iloc[0], master_rows.iloc[0]
                pdf_buffer = generate_pdf(row, inst, master, logo_path=LOGO_PATH)
                pdf_filename = f"{inst_tag}_{datetime.now().strftime('%d%m%y')}.pdf"
                zipf.writestr(pdf_filename, pdf_buffer.read())
                generated += 1

        zip_buffer.seek(0)

        # Log and UI
        st.download_button(
            "⬇️ Download All Reports (ZIP)",
            data=zip_buffer,
            file_name="CalibrationReports.zip",
            mime="application/zip",
            key="dl_all_zip"
        )
        st.session_state.activity_log.append({
            "time": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            "action": "Bulk ZIP Generated",
            "generated": generated,
            "skipped": skipped
        })

        st.success(f"ZIP ready. Generated: {generated} • Skipped: {skipped}")
        if skipped and skipped_items:
            with st.expander("See skipped items (missing Instrument or Master mapping)"):
                st.write(pd.DataFrame(skipped_items, columns=["Instrument Tag", "Master Serial No"]))

st.divider()

# --- Activity / Log panel ---
st.subheader("🧾 Activity Log")
if not st.session_state.activity_log:
    st.caption("No activity yet. Generate a single or bulk report to populate this log.")
else:
    log_df = pd.DataFrame(st.session_state.activity_log)
    st.dataframe(log_df, use_container_width=True)

# --- Optional: show a lightweight data snapshot for the window (no heavy tables) ---
with st.expander("Data snapshot (current window)"):
    preview_cols = ["Instrument Tag", "Master Serial No", "Engineer Name", "Remarks"]
    # Only show the columns that actually exist
    preview_cols = [c for c in preview_cols if c in df_filtered.columns]
    small = df_filtered.sort_values(by=timestamp_col) if timestamp_col else df_filtered
    st.dataframe(small[preview_cols].head(50), use_container_width=True)

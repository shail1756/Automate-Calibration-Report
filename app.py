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
    # CONFIGURATION (update SHEET_ID and LOGO_PATH after deployment)
    # ===============================
SERVICE_ACCOUNT_FILE = r"C:\\Users\\skm\\Desktop\\CALIBRATION REPORT AUTOMATION\\data\\service_account.json"
SHEET_ID = "1jgqN9pNWVKsH2gDVCsSmkMpYiyE0_y-vqaxrqt4B8xg"  # <-- replace with your Google Sheet ID
LOGO_PATH = r"NPL_LOGO.png"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    # ===============================
    # GSpread Auth (Streamlit Secrets friendly)
    # ===============================
        if "service_account" in st.secrets:
        svc = st.secrets["service_account"]
        if isinstance(svc, str):
            service_account_info = json.loads(svc)
        else:
            service_account_info = svc
        # gspread helper from dict (no local file required)
        try:
            gc = gspread.service_account_from_dict(service_account_info, scopes=SCOPES)
        except Exception as e:
            # fallback: convert to Credentials and authorize
            creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
            gc = gspread.authorize(creds)
    else:
        # local development fallback
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
        normal = styles["Normal"]; normal.fontSize = 9
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
        story.append(header_table); story.append(Spacer(1,8))

        # Info
        area, unit_str, location, service = inst.get("Area","BOILER"), inst.get("Unit:","Unit-1"), inst.get("Location","0M"), inst.get("SERVICE DESCRIPTION","")
        left_info = [Paragraph(f"<b>Report No:</b> {inst.get('Report No.','')}", normal),
                     Paragraph(f"<b>Calibration Date:</b> {calib_dt.strftime('%d-%m-%Y')}", normal),
                     Paragraph(f"<b>Calibration Due Date:</b> {due_dt.strftime('%d-%m-%Y')}", normal)]
        right_info = [Paragraph(f"<b>Area:</b> {area}", normal),
                      Paragraph(f"<b>Unit:</b> {unit_str}", normal),
                      Paragraph(f"<b>Location:</b> {location}", normal),
                      Paragraph(f"<b>Service:</b> {service}", normal)]
        info_table = Table([[left_info, right_info]], colWidths=[255,255])
        info_table.setStyle(TableStyle([("BOX",(0,0),(-1,-1),1,colors.black)]))
        story.append(info_table); story.append(Spacer(1,8))

        # Details
        left_details = [Paragraph("<b>Details of Instrument Under Test</b>", normal),
                        Paragraph(f"Tag No: {inst_tag}", normal),
                        Paragraph(f"Inst. Make: {inst.get('Make','')}", normal),
                        Paragraph(f"Model No: {inst.get('Model','')}", normal),
                        Paragraph(f"Sr. No: {inst.get('Sr. No.','')}", normal),
                        Paragraph(f"Range: {min_val} - {max_val} {unit}", normal),
                        Paragraph(f"Type: {inst_type}", normal)]
        right_details = [Paragraph("<b>Details of Calibration Master Instrument</b>", normal),
                         Paragraph(f"Make/Inst.Type: {master.get('Make/Inst.Type','')}", normal),
                         Paragraph(f"Make: {master.get('Make','')}", normal),
                         Paragraph(f"Model No: {master.get('Model','')}", normal),
                         Paragraph(f"Serial No: {master.get('Serial No.','')}", normal),
                         Paragraph(f"Certificate No: {master.get('Certificate No.','')}", normal),
                         Paragraph(f"Valid Upto: {master.get('Certificate Valid Upto','')}", normal)]
        details_table = Table([[left_details, right_details]], colWidths=[255,255])
        details_table.setStyle(TableStyle([("BOX",(0,0),(-1,-1),1,colors.black)]))
        story.append(details_table); story.append(Spacer(1,8))

        # Calibration Table
        calib_table_data, col_widths = [], []

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
    # STREAMLIT APP (UI)
    # ===============================
    st.set_page_config(page_title="Calibration Report Automation", page_icon="📑", layout="wide")
    st.title("📑 Calibration Report Automation")

    # Date range filter
    if timestamp_col:
        df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce")
        min_date, max_date = df[timestamp_col].min(), df[timestamp_col].max()
        start_date, end_date = st.date_input("Select date range", [min_date.date(), max_date.date()])
        mask = (df[timestamp_col].dt.date >= start_date) & (df[timestamp_col].dt.date <= end_date)
        df = df[mask]

    if st.button("Generate Reports"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for idx, row in df.iterrows():
                inst_tag = row.get("Instrument Tag", "")
                master_serial = row.get("Master Serial No", "")
                inst_rows = instrument_list[instrument_list["TAG"].str.upper() == str(inst_tag).upper()]
                master_rows = master_instrument_list[master_instrument_list["Serial No."].astype(str).str.upper() == str(master_serial).upper()]
                if inst_rows.empty or master_rows.empty:
                    continue
                inst, master = inst_rows.iloc[0], master_rows.iloc[0]
                pdf_buffer = generate_pdf(row, inst, master, logo_path=LOGO_PATH)
                pdf_filename = f"{inst_tag}_{datetime.now().strftime('%d%m%y')}.pdf"
                zipf.writestr(pdf_filename, pdf_buffer.read())
        zip_buffer.seek(0)
        st.download_button("⬇️ Download All Reports (ZIP)", data=zip_buffer, file_name="CalibrationReports.zip", mime="application/zip")

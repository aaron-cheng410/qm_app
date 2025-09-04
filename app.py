import os
import io
import json
import smtplib
import tempfile
from datetime import date, datetime
from email.message import EmailMessage

import gspread
import pillow_heif
import streamlit as st
from PIL import Image
from oauth2client.service_account import ServiceAccountCredentials
from openai import OpenAI
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import pandas as pd

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {visibility: hidden;}
    .viewerBadge_link__1S137 {display: none !important;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
    .viewerBadge_link__qRIco {display: none;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# =========================
# Config / Secrets
# =========================
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", ""))
SMTP_USER = st.secrets.get("SMTP_USER", os.getenv("SMTP_USER", "you@example.com"))
SMTP_PASS = st.secrets.get("SMTP_PASS", os.getenv("SMTP_PASS", "app_password_here"))
SMTP_FROM = st.secrets.get("SMTP_FROM", os.getenv("SMTP_FROM", "dev@yourdomain.com"))


if not OPENAI_API_KEY:
    st.warning("OPENAI_API_KEY not set. Image descriptions will fail until configured.")

# OpenAI Setup
client = OpenAI(api_key=OPENAI_API_KEY)

# =========================
# Email
# =========================
def send_notification_email(to_emails, subject="QM Submission Notification", body=None, cc=None, bcc=None):
    # normalize
    def _as_list(v):
        if not v: return []
        return v if isinstance(v, (list, tuple, set)) else [v]

    to_list  = _as_list(to_emails)
    cc_list  = _as_list(cc)
    bcc_list = _as_list(bcc)

    msg = EmailMessage()
    msg.set_content(body or "Sheet updated")
    msg["Subject"] = subject
    msg["From"] = SMTP_FROM
    if to_list: msg["To"] = ", ".join(to_list)
    if cc_list: msg["Cc"] = ", ".join(cc_list)
    if bcc_list: msg["Bcc"] = ", ".join(bcc_list)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)



# =========================
# HEIC to JPEG conversion
# =========================
def convert_heic_to_jpeg(uploaded_file):
    uploaded_file.seek(0)
    heif_file = pillow_heif.read_heif(uploaded_file.read())
    image = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data)
    jpeg_bytes = io.BytesIO()
    image.save(jpeg_bytes, format="JPEG")
    jpeg_bytes.seek(0)
    # Ensure a name for downstream consumers
    orig_name = getattr(uploaded_file, "name", "upload.heic")
    jpeg_bytes.name = orig_name.replace(".heic", ".jpeg").replace(".heif", ".jpeg")
    return jpeg_bytes

# =========================
# Google Drive Upload
# =========================
def upload_to_drive(uploaded_file, filename, folder_id):
    gauth = GoogleAuth()
    creds_dict = st.secrets["gcp_service_account"]
    scope = ["https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gauth.credentials = creds
    drive = GoogleDrive(gauth)

    # Make sure file stream at beginning
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_path = tmp_file.name

    file_metadata = {'title': filename}
    if folder_id:
        file_metadata['parents'] = [{'id': folder_id}]

    gfile = drive.CreateFile(file_metadata)
    gfile.SetContentFile(tmp_path)
    gfile.Upload()

    return f"https://drive.google.com/file/d/{gfile['id']}/view"

# =========================
# Client → Subclient config
# Each subclient maps to its own Spreadsheet (not a tab)
# =========================
subclient_config = {
    "Mazi": {
        "Mazi 1": {
            "sheet": "QM - Mazi 1",   # Spreadsheet name
            "folder": "1XUUYPOuP5IaeONMSyIsjMnqHOTkqyAKx",
            "email": ['Jayroan11@outlook.com', 'info@quickmovers.us', 'mazi20000@hotmail.com', 'Jayroan11@gmail.com'],
        },
        "Mazi 2": {
            "sheet": "QM - Mazi 2",
            "folder": "1HEYexy-TYv9Lq-onPpkf7CCH_0DWfGSX",
            "email": ['Jayroan11@outlook.com', 'info@quickmovers.us', 'mazi20000@hotmail.com', 'Jayroan11@gmail.com'],
        },
    },
    "Jay": {
        "Jay 1": {
            "sheet": "QM - Jay 1",
            "folder": "1gTkRU5h8nUb8lVhk_eyRE-94oE7gM1_J",
            "email": ['Jayroan11@outlook.com', 'info@quickmovers.us', 'mazi20000@hotmail.com', 'Jayroan11@gmail.com'],
        },
        "Jay 2": {
            "sheet": "QM - Jay 2",
            "folder": "1eDAGMx56162-YgBapkc9frIj9LIiM-rM",
            "email": ['Jayroan11@outlook.com', 'info@quickmovers.us', 'mazi20000@hotmail.com', 'Jayroan11@gmail.com'],
        },
    },
}

# =========================
# Session state init
# =========================
if "num_rows" not in st.session_state:
    st.session_state.num_rows = 1
if "dates" not in st.session_state:
    st.session_state.dates = [date.today()]

# Refresh date values if they got stringified
for i in range(len(st.session_state.dates)):
    if isinstance(st.session_state.dates[i], str):
        try:
            st.session_state.dates[i] = datetime.strptime(st.session_state.dates[i], "%m/%d/%Y").date()
        except Exception:
            st.session_state.dates[i] = date.today()

# =========================
# App UI
# =========================
st.title("QM Submission")

if st.button("Reset Form"):
    st.session_state.clear()
    st.rerun()

# 1) Client select
client_names = list(subclient_config.keys())
selected_client = st.selectbox("Select Client", [""] + client_names, key="client_select")
if not selected_client:
    st.stop()

# 2) Subclient select (dependent), and only then show add/remove + form
subclient_names = list(subclient_config[selected_client].keys())
selected_subclient = st.selectbox("Select Subclient", [""] + subclient_names, key="subclient_select")
if not selected_subclient:
    st.stop()

# Adjust date array length to num_rows
def _sync_dates_len():
    if len(st.session_state.dates) < st.session_state.num_rows:
        st.session_state.dates.extend([date.today()] * (st.session_state.num_rows - len(st.session_state.dates)))
    elif len(st.session_state.dates) > st.session_state.num_rows:
        st.session_state.dates = st.session_state.dates[:st.session_state.num_rows]

# 3) Add/Remove Date buttons (AFTER subclient chosen)
col_add, col_del = st.columns(2)
with col_add:
    if st.button("➕ Add Date"):
        st.session_state.num_rows += 1
        st.session_state.dates.append(date.today())
with col_del:
    if st.button("🗑️ Remove Last Date") and st.session_state.num_rows > 1:
        st.session_state.num_rows -= 1
        if st.session_state.dates:
            st.session_state.dates.pop()
_sync_dates_len()

# 4) Form
with st.form("multi_entry_form"):
    entries = []
    for i in range(st.session_state.num_rows):
        st.markdown(f"### Entry {i + 1}")

        st.session_state.dates[i] = st.date_input(
            "Date",
            value=st.session_state.dates[i],
            key=f"date_{i}"
        )

        col1, col2 = st.columns(2)
        with col1:
            uploaded = st.file_uploader(
                "Upload Image",
                type=["jpg", "jpeg", "png", "heif", "heic"],
                key=f"file_{i}"
            )
        with col2:
            camera_photo = st.camera_input("Take a Photo", key=f"camera_{i}")

        uploaded_file = camera_photo if camera_photo else uploaded

        carrier = st.text_input("Carrier", key=f"carrier_{i}")
        unique_id = st.text_input("Unique ID Tag", key=f"uniqueid_{i}")
        quantity = st.number_input("Quantity", min_value=1, step=1, value=1, key=f"quantity_{i}")
        from_form = st.text_input("From", key=f"from_{i}")
        location = st.selectbox("Location Stored", ["", "Front", "Back", "Left"], key=f"location_{i}")
        condition = st.text_input("Condition", key=f"condition_{i}")
        inspect = st.selectbox("Inspect", ["", "Yes", "No"], key=f"inspect_{i}")
        wo_input = st.text_input("WO#", key=f"wo_{i}")
        note_text = st.text_area("Note", key=f"note_{i}")

        # Optional: image description from OpenAI Vision
        description = ""
        if uploaded_file and OPENAI_API_KEY:
            try:
                # Convert HEIC/HEIF if needed for description extraction
                to_describe = uploaded_file
                if uploaded_file.type in ["image/heic", "image/heif"]:
                    to_describe = convert_heic_to_jpeg(uploaded_file)

                # Ensure stream at beginning
                try:
                    to_describe.seek(0)
                except Exception:
                    pass

                # The OpenAI SDK expects a file-like object with a name
                if not getattr(to_describe, "name", None):
                    to_describe.name = "image.jpg"

                file_id = client.files.create(file=to_describe, purpose="vision").id
                response = client.responses.create(
                    model="gpt-4.1-mini",
                    input=[{
                        "role": "user",
                        "content": [
                            {"type": "input_text", "text": "What object is in the image and what color is it? Respond only with the object and color, e.g. 'white couch'."},
                            {"type": "input_image", "file_id": file_id},
                        ],
                    }],
                )
                description = response.output[0].content[0].text.strip()
            except Exception as e:
                st.warning(f"Image description failed: {e}")

        entry = {
            "Date": st.session_state.dates[i].strftime("%m/%d/%Y"),
            "Client": selected_client,
            "Subclient": selected_subclient,     # fixed per form, not per row
            "Unique ID Tag": unique_id,
            "Content": description,
            "Quantity": quantity,
            "From": from_form,
            "Carrier": carrier,
            "Location Stored": location,
            "Condition": condition,
            "Inspect": inspect,
            "WO#": int(wo_input) if wo_input.strip().isdigit() else None,
            "Out": None,
            "Note": note_text,
        }
        entries.append(entry)

    review_clicked = st.form_submit_button("Review Summary")

# 5) Review + Submit
if review_clicked:
    st.session_state.entries_preview = entries

if "entries_preview" in st.session_state:
    df_preview = pd.DataFrame(st.session_state.entries_preview)
    if not df_preview.empty:
        st.subheader("📅 Pending Submissions")
        st.dataframe(df_preview)

        notify = st.checkbox("Send notification email on submission", value=False)

        if st.button("✅ Submit to Google Sheet"):
            # Pull mapping for this (client, subclient)
            cfg = subclient_config[selected_client][selected_subclient]
            sheet_name = cfg["sheet"]        # spreadsheet per subclient
            folder_id = cfg["folder"]
            to_email = cfg.get("email")

            def upload_to_google_sheet(df, sheet_name):
                from gspread.utils import rowcol_to_a1
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds_dict = st.secrets["gcp_service_account"]
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                gc = gspread.authorize(creds)
                sheet = gc.open(sheet_name)
                sheet_url = sheet.url  # <-- this gives the Google Sheet URL

                worksheet = sheet.worksheet("TEST")
                existing = worksheet.get_all_values()
                start_row = len(existing) + 1 if existing else 2

                cols = [
                    "Date", "Client", "Subclient", "Unique ID Tag", "Content", "Quantity",
                    "From", "Carrier", "Location Stored", "Condition", "Inspect", "WO#",
                    "Out", "Note", "Drive Link"
                ]
                for col in cols:
                    if col not in df.columns:
                        df[col] = ""

                df = df[cols]
                data = df.values.tolist()
                cell_range = f"A{start_row}:{rowcol_to_a1(start_row + len(data) - 1, len(cols))}"
                worksheet.update(cell_range, data, value_input_option="USER_ENTERED")

                return sheet_url  # <-- return it


            # Add Drive links (support file_uploader OR camera_input)
            for i, row in enumerate(st.session_state.entries_preview):
                uploaded_file = st.session_state.get(f"file_{i}") or st.session_state.get(f"camera_{i}")
                if uploaded_file:
                    file_to_upload = uploaded_file
                    # Convert HEIC/HEIF for Drive
                    if uploaded_file.type in ["image/heic", "image/heif"]:
                        file_to_upload = convert_heic_to_jpeg(uploaded_file)

                    # Filename fallback for camera photos (may have no name)
                    filename = getattr(file_to_upload, "name", f"image_{i}.jpg")
                    try:
                        drive_link = upload_to_drive(file_to_upload, filename, folder_id)
                        row["Drive Link"] = f'=HYPERLINK("{drive_link}", "Photo")'
                    except Exception as e:
                        row["Drive Link"] = ""
                        st.error(f"Drive upload failed for row {i+1}: {e}")
                else:
                    row["Drive Link"] = ""

            df_result = pd.DataFrame(st.session_state.entries_preview)
            try:
                sheet_url = upload_to_google_sheet(df_result, sheet_name)  # <-- capture URL
                st.success("✅ Submitted to Google Sheet")
            except Exception as e:
                st.error(f"Sheet update failed: {e}")
                st.stop()

            if notify and to_email:
                body = f"""Hi,

            We’re happy to let you know that your items have been processed and are in great condition.

            Use this link below to track your items and see images of each one:
            {sheet_url}

            If you have any questions, don’t hesitate to reach out—we’re here to help.

            Thanks again for choosing Quick Movers!

            Best,
            The Quick Movers Team"""

                try:
                    send_notification_email(
                        to_emails=to_email,
                        subject="Quick Movers – Your items have been processed",
                        body=body
                    )
                    st.success(f"Notification email sent to {to_email}")
                except Exception as e:
                    st.error(f"Failed to send email: {e}")


            # Reset
            del st.session_state["entries_preview"]
            st.session_state.num_rows = 1
            st.session_state.dates = [date.today()]

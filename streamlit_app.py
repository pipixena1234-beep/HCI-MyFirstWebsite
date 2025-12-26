import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import zipfile
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
import time
from datetime import datetime
import requests  # <--- THIS WAS MISSING
import base64    # <--- ALSO NEEDED FOR THE GITHUB API

st.set_page_config(page_title="Student Progress Reports", layout="wide")
st.title("📊 Student Progress Report System (Flattened, Term-aware)")

# =====================================
# Report Generation Automation - Date Selection
# =====================================

def push_schedule_to_github(new_datetime_str):
    token = st.secrets["GITHUB_TOKEN"]
    # Update this with your actual GitHub username and repo name
    repo = "pipixena1234-beep/HCI-MyFirstWebsite" 
    path = "schedule.json"
    
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 1. Get the current file to get its SHA (required for updates)
    get_res = requests.get(url, headers=headers)
    sha = get_res.json().get("sha") if get_res.status_code == 200 else None

    # 2. Prepare the new file content
    content_dict = {"target_datetime": new_datetime_str}
    content_json = json.dumps(content_dict, indent=4)
    encoded_content = base64.b64encode(content_json.encode()).decode()

    # 3. Send the update request
    payload = {
        "message": f"Update schedule to {new_datetime_str} [skip ci]",
        "content": encoded_content
    }
    if sha:
        payload["sha"] = sha

    put_res = requests.put(url, json=payload, headers=headers)
    return put_res.status_code

# --- In your Sidebar UI ---
st.sidebar.subheader("⏰ Automation Schedule")
date_pick = st.sidebar.date_input("Target Date", datetime.now())
time_pick = st.sidebar.time_input("Target Time", datetime.now())

target_str = f"{date_pick.strftime('%Y-%m-%d')} {time_pick.strftime('%H:%M')}"

if st.sidebar.button("Update GitHub Schedule"):
    with st.spinner("Pushing to GitHub..."):
        status = push_schedule_to_github(target_str)
        if status in [200, 201]:
            st.sidebar.success(f"✅ GitHub updated to {target_str}!")
        else:
            st.sidebar.error(f"❌ Failed to update. Error code: {status}")


# =====================================
# Parse stacked tables → ONE clean table
# =====================================
def extract_and_flatten(df_raw):
    rows = []
    i = 0
    while i < len(df_raw):
        cell = str(df_raw.iloc[i, 0])
        if cell.startswith("Term:"):
            term = cell.replace("Term:", "").strip()
            header_row = i + 2
            headers = [str(h).strip() for h in df_raw.iloc[header_row].tolist()]
            j = header_row + 1
            while j < len(df_raw) and pd.notna(df_raw.iloc[j, 0]):
                row = dict(zip(headers, df_raw.iloc[j].tolist()))
                row["Term"] = term
                rows.append(row)
                j += 1
            i = j
        else:
            i += 1
    return pd.DataFrame(rows)

# =========================
# Upload Excel
# =========================
uploaded_file = st.file_uploader("Upload Excel (.xlsx) with stacked term tables", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    selected_sheet = st.selectbox("Select Sheet (Subject)", xls.sheet_names)
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
    df = extract_and_flatten(df_raw)

    if df.empty:
        st.error("❌ No valid term tables detected.")
        st.stop()

    # =========================
    # Clean columns
    # =========================
    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df[skills] = df[skills].apply(pd.to_numeric)

    # =========================
    # Term selection
    # =========================
    st.subheader("📅 Term Selection")
    select_terms = st.checkbox("Select terms to generate reports for")
    all_terms = sorted(df["Term"].unique())

    if select_terms:
        options_with_select_all = ["Select All"] + all_terms
        selected_terms = st.multiselect("Available terms:", options_with_select_all, default=all_terms)
        if "Select All" in selected_terms:
            selected_terms = all_terms
    else:
        selected_terms = all_terms

    df = df[df["Term"].isin(selected_terms)]

    # =========================
    # Grading
    # =========================
    df["Average"] = df[skills].mean(axis=1)

    def grade(avg):
        if avg >= 80: return "A"
        elif avg >= 70: return "B"
        elif avg >= 60: return "C"
        elif avg >= 50: return "D"
        else: return "F"

    df["Grade"] = df["Average"].apply(grade)
    df["Remarks"] = df["Average"].apply(
        lambda x: "Excellent work!" if x >= 80 else
                  "Good effort, keep improving!" if x >= 70 else
                  "Needs improvement"
    )

    # =========================
    # Dashboard
    # =========================
    st.header(f"📘 Dashboard – {selected_sheet}")
    st.dataframe(df)
    for term in selected_terms:
        st.subheader(f"Term: {term}")
        st.bar_chart(df[df["Term"] == term][skills].mean())

    # =========================
    # Generate ZIP
    # =========================
    if st.button("📦 Generate PDFs (ZIP)"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for _, row in df.iterrows():
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, f"Progress Report ({row['Term']})", ln=True)
                pdf.set_font("Arial", "", 12)
                pdf.cell(0, 8, f"Student: {row['Student Name'].strip()}", ln=True)
                for s in skills:
                    pdf.cell(0, 8, f"{s}: {row[s]}", ln=True)
                pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)
                pdf.cell(0, 8, f"Grade: {row['Grade']}", ln=True)
                pdf.cell(0, 8, f"Remarks: {row['Remarks']}", ln=True)
                pdf_bytes = BytesIO()
                pdf_bytes.write(pdf.output(dest="S").encode("latin-1"))
                pdf_bytes.seek(0)
                zip_file.writestr(f"{row['Term']}/{row['Student Name'].strip()}_report.pdf", pdf_bytes.read())
        zip_buffer.seek(0)
        st.download_button("⬇️ Download ZIP", data=zip_buffer, file_name="student_reports.zip")


    st.subheader("📤 Upload to Google Drive")
    folder_id_input = st.text_input(
        "Enter Google Drive Folder ID",
        value="0ALncbMfl-gjdUk9PVA"
    )
    # =========================
    # Upload to Google Drive
    # =========================
    if st.button("Upload to Google Drive"):
        try:
            sa_info = json.loads(st.secrets["google_service_account"]["google_service_account"])
            credentials = service_account.Credentials.from_service_account_info(
                sa_info, scopes=['https://www.googleapis.com/auth/drive']
            )
            drive_service = build('drive', 'v3', credentials=credentials)
    
            # --- FIX: Calculate EXACT total steps ---
            # Only count rows that belong to the selected terms
            df_to_process = df[df["Term"].isin(selected_terms)]
            total_steps = len(df_to_process)
            
            if total_steps == 0:
                st.warning("No data found for the selected terms.")
            else:
                prog_bar = st.progress(0)
                status_text = st.empty() # Placeholder for "Processing Alice Tan..."
                current_step = 0
    
                for term in selected_terms:
                    term_clean = term.strip()
                    parent_id = folder_id_input.strip()
    
                    # 1. FIND OR CREATE TERM FOLDER
                    query_folder = (
                        f"name='{term_clean}' "
                        f"and mimeType='application/vnd.google-apps.folder' "
                        f"and '{parent_id}' in parents and trashed=false"
                    )
                    folder_search = drive_service.files().list(
                        q=query_folder, fields="files(id)", 
                        includeItemsFromAllDrives=True, supportsAllDrives=True
                    ).execute()
                    
                    existing_folders = folder_search.get('files', [])
                    if existing_folders:
                        term_folder_id = existing_folders[0]['id']
                    else:
                        folder_metadata = {'name': term_clean, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
                        term_folder = drive_service.files().create(body=folder_metadata, fields='id', supportsAllDrives=True).execute()
                        term_folder_id = term_folder['id']
    
                    # 2. UPLOAD / OVERWRITE PDFs
                    df_term = df[df["Term"] == term]
                    
                    for _, row in df_term.iterrows():
                        student_name = row['Student Name'].strip()
                        file_name = f"{selected_sheet}_{student_name}_report.pdf"
                        
                        # Update status text
                        status_text.text(f"Uploading: {file_name}")
    
                        # --- Generate PDF ---
                        pdf = FPDF()
                        pdf.add_page()
                        pdf.set_font("Arial", "B", 16)
                        pdf.cell(0, 10, f"Progress Report ({selected_sheet}_{term_clean})", ln=True)
                        pdf.set_font("Arial", "", 12)
                        pdf.cell(0, 8, f"Student: {student_name}", ln=True)
                        for s in skills:
                            pdf.cell(0, 8, f"{s}: {row[s]}", ln=True)
                        pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)
                        pdf.cell(0, 8, f"Grade: {row['Grade']}", ln=True)
                        pdf.cell(0, 8, f"Remarks: {row['Remarks']}", ln=True)
    
                        pdf_bytes = BytesIO()
                        pdf_bytes.write(pdf.output(dest="S").encode("latin-1"))
                        pdf_bytes.seek(0)
                        media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf', resumable=True)
    
                        # --- Check if FILE already exists for Overwrite ---
                        query_file = f"name='{file_name}' and '{term_folder_id}' in parents and trashed=false"
                        file_search = drive_service.files().list(q=query_file, fields="files(id)", includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
                        existing_files = file_search.get('files', [])
    
                        if existing_files:
                            drive_service.files().update(fileId=existing_files[0]['id'], media_body=media, supportsAllDrives=True).execute()
                        else:
                            file_metadata = {'name': file_name, 'parents': [term_folder_id]}
                            drive_service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
    
                        # --- FIX: Increment progress based on ACTUAL processed rows ---
                        current_step += 1
                        prog_bar.progress(current_step / total_steps)
    
                status_text.empty() # Clear status when done
                st.success(f"✅ Successfully processed {total_steps} reports!")
    
        except Exception as e:
            st.error(f"Google Drive operation failed: {e}")
        

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

st.set_page_config(page_title="Student Progress Reports", layout="wide")
st.title("📊 Student Progress Report System (Flattened, Term-aware)")

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

    st.subheader("⚠️ Delete All Reports (Testing Only)")
    delete_folder_id = st.text_input(
        "Enter Parent Google Drive Folder ID to clear",
        value="0ALncbMfl-gjdUk9PVA"
    )
    
    if st.button("🗑️ Delete All Reports"):
        try:
            sa_info = json.loads(st.secrets["google_service_account"]["google_service_account"])
            credentials = service_account.Credentials.from_service_account_info(
                sa_info, scopes=['https://www.googleapis.com/auth/drive']
            )
            drive_service = build('drive', 'v3', credentials=credentials)
    
            # 1. FIXED QUERY: Look for items where the parent is the ID provided
            # We also set pageSize=1000 to catch more files in one go (default is 100)
            query = f"'{delete_folder_id}' in parents and trashed = false"
    
            files = drive_service.files().list(
                q=query,
                pageSize=1000, 
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
                fields="files(id, name)"
            ).execute()
    
            items = files.get('files', [])
    
            if not items:
                st.warning("Folder is already empty (or invalid ID provided).")
            else:
                deleted_count = 0
                my_bar = st.progress(0)
                
                for i, f in enumerate(items):
                    try:
                        drive_service.files().update(
                            fileId=f['id'],
                            body={'trashed': True},
                            supportsAllDrives=True
                        ).execute()
                        deleted_count += 1
                    except Exception as e:
                        st.error(f"Error deleting {f['name']}: {e}")
                    
                    # Update progress
                    my_bar.progress((i + 1) / len(items))
    
                st.success(f"✅ Successfully trashed {deleted_count} files!")
    
        except Exception as e:
            st.error(f"Failed to delete reports: {e}")

        
     st.subheader("📤 Upload to Google Drive")
    folder_id_input = st.text_input(
        "Enter Google Drive Folder ID",
        value="0ALncbMfl-gjdUk9PVA"
    )
    
    if st.button("Upload to Google Drive"):
        try:
            sa_info = json.loads(st.secrets["google_service_account"]["google_service_account"])
            credentials = service_account.Credentials.from_service_account_info(
                sa_info, scopes=['https://www.googleapis.com/auth/drive']
            )
            drive_service = build('drive', 'v3', credentials=credentials)
    
            # Create a progress bar
            prog_bar = st.progress(0)
            total_steps = len(selected_terms) * len(df) # Approximate steps
            current_step = 0
    
            for term in selected_terms:
                term_clean = term.strip()
                
                # ==========================================
                # 1. FIXED DELETION LOGIC
                # ==========================================
                # Find existing folders with this Term Name inside the parent folder
                query_existing_folders = (
                    f"name='{term_clean}' "
                    f"and mimeType='application/vnd.google-apps.folder' "
                    f"and '{folder_id_input.strip()}' in parents "
                    f"and trashed=false"
                )
                
                # IMPORTANT: Added includeItemsFromAllDrives=True
                existing_folders = drive_service.files().list(
                    q=query_existing_folders,
                    fields="files(id, name)",
                    includeItemsFromAllDrives=True, 
                    supportsAllDrives=True
                ).execute()
    
                # Delete (Trash) the old folders
                for folder in existing_folders.get('files', []):
                    try:
                        drive_service.files().update(
                            fileId=folder['id'],
                            body={'trashed': True},
                            supportsAllDrives=True
                        ).execute()
                    except Exception as e:
                        st.warning(f"Could not delete old folder '{folder['name']}': {e}")
                    time.sleep(0.2) # Short pause to avoid API limits
    
                # ==========================================
                # 2. CREATE FRESH FOLDER
                # ==========================================
                folder_metadata = {
                    'name': term_clean,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'parents': [folder_id_input.strip()]
                }
                term_folder = drive_service.files().create(
                    body=folder_metadata,
                    fields='id',
                    supportsAllDrives=True
                ).execute()
                term_folder_id = term_folder['id']
    
                # ==========================================
                # 3. UPLOAD PDFs (Optimized)
                # ==========================================
                df_term = df[df["Term"] == term]
                
                for _, row in df_term.iterrows():
                    student_name = row['Student Name'].strip()
                    
                    # --- Generate PDF (Your existing logic) ---
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 16)
                    pdf.cell(0, 10, f"Progress Report ({term_clean})", ln=True)
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
    
                    file_name = f"{student_name}_report.pdf"
    
                    # --- UPLOAD ---
                    # Since we just created a NEW empty folder, we don't need to check 
                    # if the file exists. We can just create it directly.
                    file_metadata = {
                        'name': file_name, 
                        'parents': [term_folder_id]
                    }
                    
                    try:
                        drive_service.files().create(
                            body=file_metadata,
                            media_body=media,
                            fields='id',
                            supportsAllDrives=True
                        ).execute()
                    except Exception as upload_err:
                        st.error(f"Failed to upload {file_name}: {upload_err}")
    
                    # Update Progress
                    current_step += 1
                    if current_step < total_steps:
                        prog_bar.progress(current_step / total_steps)
    
            prog_bar.progress(1.0)
            st.success("✅ PDFs uploaded to Google Drive successfully!")
    
        except Exception as e:
            st.error(f"Google Drive upload failed: {e}")


    

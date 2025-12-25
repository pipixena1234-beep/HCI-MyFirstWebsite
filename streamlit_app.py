import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import zipfile
import json
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

st.title("📊 Student Progress Report System")

# --- 1. CSV Upload ---
uploaded_file = st.file_uploader("Upload CSV", type=["csv"])
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df["Average"] = df[skills].mean(axis=1)

    # Skill-based grading
    def grade(avg):
        if avg >= 80: return "A"
        elif avg >= 70: return "B"
        elif avg >= 60: return "C"
        elif avg >= 50: return "D"
        else: return "F"
    df["Grade"] = df["Average"].apply(grade)

    # AI-style remarks
    def remarks(avg):
        if avg >= 80: return "Excellent work!"
        elif avg >= 70: return "Good effort, keep improving!"
        else: return "Needs improvement, focus on practice."
    df["Remarks"] = df["Average"].apply(remarks)

    st.header("📄 Class Dashboard")
    st.dataframe(df)

    # --- Select students ---
    student_names = df["Student Name"].tolist()
    selected_students = st.multiselect("Select students to generate PDF", student_names)

    # --- Generate PDFs in memory ---
    pdf_files = {}
    for idx, row in df.iterrows():
        if row["Student Name"] not in selected_students:
            continue

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, f"Progress Report: {row['Student Name']}", ln=True)
        pdf.set_font("Arial", "", 12)
        for skill in skills:
            pdf.cell(0, 8, f"{skill}: {row[skill]}", ln=True)
        pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)
        pdf.cell(0, 8, f"Grade: {row['Grade']}", ln=True)
        pdf.cell(0, 8, f"Remarks: {row['Remarks']}", ln=True)

        pdf_bytes = BytesIO()
        pdf_output = pdf.output(dest='S').encode('latin-1')
        pdf_bytes.write(pdf_output)
        pdf_bytes.seek(0)
        pdf_files[row["Student Name"]] = pdf_bytes

    # --- Multi-select ZIP download ---
    if pdf_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for name, pdf_bytes in pdf_files.items():
                zf.writestr(f"{name}_report.pdf", pdf_bytes.read())
                pdf_bytes.seek(0)
        zip_buffer.seek(0)
        st.download_button("Download All Selected PDFs as ZIP", data=zip_buffer, file_name="student_reports.zip")

    # --- Google Drive Upload ---
    st.header("📤 Upload to Google Drive")
    drive_folder_id = st.text_input("Enter Google Drive Folder ID", "1mxhb5P7qob_lhfXdMeC2HWwKP9lDahmU")
    if st.button("Upload PDFs to Drive") and pdf_files:
        try:
            # Load service account JSON from Streamlit Secrets
            sa_info = json.loads(st.secrets["google_service_account"])
            with open("service_account.json", "w") as f:
                json.dump(sa_info, f)

            gauth = GoogleAuth()
            gauth.credentials = ServiceAccountCredentials.from_json_keyfile_name(
                "service_account.json",
                scopes=["https://www.googleapis.com/auth/drive"]
            )
            drive = GoogleDrive(gauth)

            for name, pdf_bytes in pdf_files.items():
                file_drive = drive.CreateFile({
                    "title": f"{name}_report.pdf",
                    "parents": [{"id": drive_folder_id}]
                })
                pdf_bytes.seek(0)
                file_drive.SetContentString(pdf_bytes.read().decode("latin-1"))
                file_drive.Upload()
            st.success("✅ All selected PDFs uploaded to Google Drive!")
        except Exception as e:
            st.error(f"Google Drive upload failed: {e}")

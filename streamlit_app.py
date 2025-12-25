import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import zipfile
import json

# Google API imports
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseUpload

st.set_page_config(page_title="Student Progress Reports", layout="wide")
st.title("📊 Student Progress Report System (Cloud-ready)")

# --- CSV Upload ---
uploaded_file = st.file_uploader(
    "Upload CSV (Columns: Student Name, Logic, UI, Animation, Teamwork)",
    type=["csv"]
)

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.header("📄 Class Dashboard")
    st.dataframe(df)

    # --- Skill-based grading ---
    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df["Average"] = df[skills].mean(axis=1)

    def grade(avg):
        if avg >= 80: return "A"
        elif avg >= 70: return "B"
        elif avg >= 60: return "C"
        elif avg >= 50: return "D"
        else: return "F"

    df["Grade"] = df["Average"].apply(grade)

    # --- AI / Remarks ---
    def remarks(avg):
        if avg >= 80:
            return "Excellent work!"
        elif avg >= 70:
            return "Good effort, keep improving!"
        else:
            return "Needs improvement, focus on practice."

    df["Remarks"] = df["Average"].apply(remarks)

    st.subheader("Average Skills")
    st.bar_chart(df[skills].mean())

    # --- Multi-select PDF / ZIP download ---
    student_options = df['Student Name'].tolist()
    selected_students = st.multiselect(
        "Select students to download PDF",
        options=student_options,
        default=student_options
    )

    if selected_students:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for student in selected_students:
                row = df[df['Student Name'] == student].iloc[0]

                # Generate PDF in-memory
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
                pdf_bytes.write(pdf.output(dest='S').encode('latin-1'))
                pdf_bytes.seek(0)

                zip_file.writestr(f"{row['Student Name']}_report.pdf", pdf_bytes.read())

        zip_buffer.seek(0)
        st.download_button(
            "📦 Download PDFs for Selected Students (ZIP)",
            data=zip_buffer,
            file_name="student_reports.zip"
        )

        # --- Google Drive Upload via official API ---
        st.subheader("📤 Upload selected PDFs to Google Drive")
        folder_id_input = st.text_input(
            "Enter Google Drive Folder ID",
            value="1mxhb5P7qob_lhfXdMeC2HWwKP9lDahmU"
        )

        if st.button("Upload to Google Drive"):
            try:
                # Authenticate using service account from secrets
                sa_info = json.loads(st.secrets["general"]["google_service_account"])
                credentials = service_account.Credentials.from_service_account_info(
                    sa_info, scopes=["https://www.googleapis.com/auth/drive"]
                )
                service = build('drive', 'v3', credentials=credentials)

                for student in selected_students:
                    row = df[df['Student Name'] == student].iloc[0]

                    # Generate PDF again for upload
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
                    pdf_bytes.write(pdf.output(dest='S').encode('latin-1'))
                    pdf_bytes.seek(0)

                    media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf')
                    file_metadata = {
                        'name': f"{row['Student Name']}_report.pdf",
                        'parents': [folder_id_input]
                    }
                    service.files().create(body=file_metadata, media_body=media).execute()

                st.success("✅ PDFs uploaded successfully to Google Drive!")

            except Exception as e:
                st.error(f"Google Drive upload failed: {e}")

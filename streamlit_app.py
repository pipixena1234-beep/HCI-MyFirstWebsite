import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import zipfile
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account

st.set_page_config(page_title="Student Progress Reports", layout="wide")
st.title("📊 Student Progress Report System (Cloud-ready)")

# --- 1. CSV Upload ---
uploaded_file = st.file_uploader(
    "Upload CSV (Columns: Student Name, Logic, UI, Animation, Teamwork)",
    type=["csv"]
)

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.header("📄 Class Dashboard")
    st.dataframe(df)

    # --- 2. Skill-based grading ---
    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df["Average"] = df[skills].mean(axis=1)
    df["Grade"] = df["Average"].apply(
        lambda x: "A" if x >= 80 else "B" if x >= 70 else "C" if x >= 60 else "D" if x >= 50 else "F"
    )
    df["Remarks"] = df["Average"].apply(
        lambda x: "Excellent work!" if x >= 80 else "Good effort!" if x >= 70 else "Needs improvement."
    )

    st.subheader("Average Skills")
    st.bar_chart(df[skills].mean())

    # --- 3. Multi-select PDF / ZIP download ---
    selected_students = st.multiselect(
        "Select students to generate PDFs",
        df["Student Name"].tolist(),
        default=df["Student Name"].tolist()
    )

    if selected_students:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for student in selected_students:
                row = df[df["Student Name"] == student].iloc[0]

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
            "📦 Download ZIP of Selected PDFs",
            zip_buffer,
            file_name="student_reports.zip"
        )

        # --- 4. Google Drive Upload ---
        folder_id = st.text_input(
            "Enter Google Drive Folder ID",
            value="YOUR_FOLDER_ID_HERE"
        )
        if st.button("Upload PDFs to Google Drive"):
            try:
                # --- Load service account credentials locally ---
                SERVICE_ACCOUNT_FILE = "service_account.json"
                SCOPES = ['https://www.googleapis.com/auth/drive']
                credentials = service_account.Credentials.from_service_account_file(
                    SERVICE_ACCOUNT_FILE, scopes=SCOPES
                )

                drive_service = build('drive', 'v3', credentials=credentials)

                for student in selected_students:
                    row = df[df["Student Name"] == student].iloc[0]

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
                    drive_service.files().create(
                        body={'name': f"{row['Student Name']}_report.pdf", 'parents':[folder_id]},
                        media_body=media
                    ).execute()

                st.success("✅ Selected PDFs uploaded to Google Drive successfully!")

            except Exception as e:
                st.error(f"Google Drive upload failed: {e}")

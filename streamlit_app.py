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
st.title("📊 Student Progress Report System (Term-wise, Cloud-ready)")

# --- 1. Upload Excel file ---
uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx) with multiple sheets",
    type=["xlsx"]
)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # --- 2. Select sheet (subject) ---
    selected_sheet = st.selectbox("Select Sheet (Subject) to generate reports for:", sheet_names)
    
    # Read sheet and build final table with 'Term' column
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
    final_rows = []
    term = None
    for idx, row in df_raw.iterrows():
        if pd.notna(row[0]) and "Term" in str(row[0]):
            term = str(row[0]).split(":")[-1].strip()
            header_idx = idx + 2  # skip separator row
            continue
        if term and idx >= header_idx:
            if pd.notna(row[0]):
                final_rows.append(list(row[:5]) + [term])
    
    df = pd.DataFrame(final_rows, columns=["Student Name", "Logic", "UI", "Animation", "Teamwork", "Term"])
    df[["Logic", "UI", "Animation", "Teamwork"]] = df[["Logic", "UI", "Animation", "Teamwork"]].astype(float)

    st.header(f"📄 Class Dashboard - {selected_sheet}")
    st.dataframe(df)

    # --- 3. Skill-based grading ---
    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df["Average"] = df[skills].mean(axis=1)

    def grade(avg):
        if avg >= 80: return "A"
        elif avg >= 70: return "B"
        elif avg >= 60: return "C"
        elif avg >= 50: return "D"
        else: return "F"

    df["Grade"] = df["Average"].apply(grade)

    def remarks(avg):
        if avg >= 80: return "Excellent work!"
        elif avg >= 70: return "Good effort, keep improving!"
        else: return "Needs improvement, focus on practice."

    df["Remarks"] = df["Average"].apply(remarks)

    # --- 4. Select terms ---
    terms = df['Term'].unique().tolist()
    select_term_checkbox = st.checkbox("Select terms to generate reports for")
    if select_term_checkbox:
        selected_terms = st.multiselect("Available terms:", terms, default=terms)
    else:
        selected_terms = terms

    # --- 5. Show dashboard per term ---
    for term in selected_terms:
        st.subheader(f"Term: {term}")
        df_term = df[df['Term'] == term]
        st.dataframe(df_term)
        st.bar_chart(df_term[skills].mean())

    # --- 6. Generate PDFs and ZIP ---
    if st.button("Generate PDFs for selected terms"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for term in selected_terms:
                df_term = df[df['Term'] == term]
                for _, row in df_term.iterrows():
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
                    zip_file.writestr(f"{term}/{row['Student Name']}_report.pdf", pdf_bytes.read())
        zip_buffer.seek(0)
        st.download_button(
            "📦 Download ZIP of PDFs",
            data=zip_buffer,
            file_name="student_reports_by_term.zip"
        )

    # --- 7. Upload to Google Drive per term ---
    st.subheader("📤 Upload to Google Drive")
    folder_id_input = st.text_input("Enter parent Google Drive Folder ID for term folders:")

    if st.button("Upload to Google Drive"):
        try:
            sa_info = json.loads(st.secrets["google_service_account"]["google_service_account"])
            credentials = service_account.Credentials.from_service_account_info(
                sa_info, scopes=['https://www.googleapis.com/auth/drive']
            )
            drive_service = build('drive', 'v3', credentials=credentials)

            for term in selected_terms:
                # Check or create folder for term
                query = f"name='{term}' and mimeType='application/vnd.google-apps.folder' and '{folder_id_input}' in parents"
                response = drive_service.files().list(q=query, fields="files(id, name)").execute()
                if response['files']:
                    term_folder_id = response['files'][0]['id']
                else:
                    folder_metadata = {'name': term, 'mimeType': 'application/vnd.google-apps.folder', 'parents':[folder_id_input]}
                    term_folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
                    term_folder_id = term_folder['id']

                # Upload PDFs
                df_term = df[df['Term'] == term]
                for _, row in df_term.iterrows():
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

                    media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf', resumable=True)
                    file_metadata = {'name': f"{row['Student Name']}_report.pdf", 'parents':[term_folder_id]}
                    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

            st.success("✅ PDFs uploaded to Google Drive successfully!")

        except Exception as e:
            st.error(f"Google Drive upload failed: {e}")

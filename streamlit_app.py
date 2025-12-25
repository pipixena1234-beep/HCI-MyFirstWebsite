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
            headers = df_raw.iloc[header_row].tolist()
            headers = [str(h).strip() for h in headers]

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
uploaded_file = st.file_uploader(
    "Upload Excel (.xlsx) with stacked term tables",
    type=["xlsx"]
)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    selected_sheet = st.selectbox(
        "Select Sheet (Subject)",
        xls.sheet_names
    )

    df_raw = pd.read_excel(
        uploaded_file,
        sheet_name=selected_sheet,
        header=None
    )

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
        selected_terms = st.multiselect(
            "Available terms:",
            all_terms,
            default=all_terms
        )
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
                pdf.cell(0, 8, f"Student: {row['Student Name']}", ln=True)
                for s in skills:
                    pdf.cell(0, 8, f"{s}: {row[s]}", ln=True)
                pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)
                pdf.cell(0, 8, f"Grade: {row['Grade']}", ln=True)
                pdf.cell(0, 8, f"Remarks: {row['Remarks']}", ln=True)
                pdf_bytes = BytesIO()
                pdf_bytes.write(pdf.output(dest="S").encode("latin-1"))
                pdf_bytes.seek(0)
                zip_file.writestr(
                    f"{row['Term']}/{row['Student Name']}_report.pdf",
                    pdf_bytes.read()
                )
        zip_buffer.seek(0)
        st.download_button(
            "⬇️ Download ZIP",
            data=zip_buffer,
            file_name="student_reports.zip"
        )

    # =========================
    # Upload to Google Drive
    # =========================
    st.subheader("📤 Upload to Google Drive")
    folder_id_input =  st.text_input(
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

            for term in selected_terms:
                # Check or create folder for term
                query = f"name='{term}' and mimeType='application/vnd.google-apps.folder' and '{folder_id_input}' in parents"
                response = drive_service.files().list(q=query, fields="files(id, name)").execute()
                if response['files']:
                    term_folder_id = response['files'][0]['id']
                else:
                    folder_metadata = {
                        'name': term,
                        'mimeType': 'application/vnd.google-apps.folder',
                        'parents':[folder_id_input]
                    }
                    term_folder = drive_service.files().create(
                    body=folder_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                    ).execute()
                    term_folder_id = term_folder['id']

                # Upload PDFs
                df_term = df[df["Term"] == term]
                for _, row in df_term.iterrows():
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 16)
                    pdf.cell(0, 10, f"Progress Report ({row['Term']})", ln=True)
                    pdf.set_font("Arial", "", 12)
                    pdf.cell(0, 8, f"Student: {row['Student Name']}", ln=True)
                    for s in skills:
                        pdf.cell(0, 8, f"{s}: {row[s]}", ln=True)
                    pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)
                    pdf.cell(0, 8, f"Grade: {row['Grade']}", ln=True)
                    pdf.cell(0, 8, f"Remarks: {row['Remarks']}", ln=True)
                    pdf_bytes = BytesIO()
                    pdf_bytes.write(pdf.output(dest="S").encode("latin-1"))
                    pdf_bytes.seek(0)

                    media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf', resumable=True)
                    file_metadata = {'name': f"{row['Student Name']}_report.pdf", 'parents':[term_folder_id]}
                    drive_service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()

            st.success("✅ PDFs uploaded to Google Drive successfully!")

        except Exception as e:
            st.error(f"Google Drive upload failed: {e}")

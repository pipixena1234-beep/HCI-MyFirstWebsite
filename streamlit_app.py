import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from io import BytesIO
import zipfile

# Optional: Google Drive
try:
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive
    GDRIVE_AVAILABLE = True
except:
    GDRIVE_AVAILABLE = False

st.title("📊 Student Progress Report System")
st.caption("App version: 1.0")

# --- 1. CSV Upload ---
uploaded_file = st.file_uploader("Upload CSV (Student Name, Logic, UI, Animation, Teamwork)", type=["csv"])
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

    # --- 2. Skill-based grading ---
    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df["Average"] = df[skills].mean(axis=1)

    def grade(avg):
        if avg >= 80: return "A"
        elif avg >= 70: return "B"
        elif avg >= 60: return "C"
        elif avg >= 50: return "D"
        else: return "F"

    df["Grade"] = df["Average"].apply(grade)

    # --- 3. AI / Rule-based remarks ---
    def remarks(avg):
        if avg >= 80:
            return "Excellent work!"
        elif avg >= 70:
            return "Good effort, keep improving!"
        else:
            return "Needs improvement, focus on practice."

    df["Remarks"] = df["Average"].apply(remarks)

    # --- 4. Class Dashboard ---
    st.header("📄 Class Dashboard")
    st.dataframe(df)
    st.subheader("Average Skills")
    st.bar_chart(df[skills].mean())

    # --- 5. Multi-select PDF / ZIP download ---
    student_options = df['Student Name'].tolist()
    selected_students = st.multiselect(
        "Select students to download PDF",
        options=student_options,
        default=student_options  # default: all selected
    )

    if selected_students:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for student in selected_students:
                row = df[df['Student Name'] == student].iloc[0]

                # Generate PDF
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

                zip_file.writestr(f"{row['Student Name']}_report.pdf", pdf_bytes.read())

        zip_buffer.seek(0)
        st.download_button(
            "📦 Download PDFs for Selected Students (ZIP)",
            data=zip_buffer,
            file_name="student_reports.zip"
        )

        # --- 6. Optional Google Drive upload ---
        if GDRIVE_AVAILABLE:
            st.subheader("📤 Upload to Google Drive (Local / Service Account)")
            if st.button("Upload selected PDFs to Google Drive"):
                gauth = GoogleAuth()
                gauth.LocalWebserverAuth()  # local auth flow
                drive = GoogleDrive(gauth)
                for student in selected_students:
                    row = df[df['Student Name'] == student].iloc[0]
                    # Generate PDF again
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

                    file_drive = drive.CreateFile({'title': f"{row['Student Name']}_report.pdf"})
                    file_drive.SetContentString(pdf_bytes.read().decode('latin-1'))
                    file_drive.Upload()
                st.success("✅ Selected PDFs uploaded to Google Drive successfully!")

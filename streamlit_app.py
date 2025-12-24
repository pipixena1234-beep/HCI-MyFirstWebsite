import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from io import BytesIO
import zipfile

st.title("📊 Student Progress Report System")

# --- 1. CSV Upload ---
uploaded_file = st.file_uploader("Upload CSV", type=["csv"])
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df["Average"] = df[skills].mean(axis=1)

    def grade(avg):
        if avg >= 80: return "A"
        elif avg >= 70: return "B"
        elif avg >= 60: return "C"
        elif avg >= 50: return "D"
        else: return "F"

    df["Grade"] = df["Average"].apply(grade)

    # --- 2. AI / Rule-Based Remarks ---
    def remarks(avg):
        if avg >= 80:
            return "Excellent work!"
        elif avg >= 70:
            return "Good effort, keep improving!"
        else:
            return "Needs improvement, focus on practice."
    df["Remarks"] = df["Average"].apply(remarks)

    st.header("📄 Class Dashboard")
    st.dataframe(df)

    # --- 3. Skill Charts ---
    st.bar_chart(df[skills].mean())

    # --- 4. Multiselect for students ---
    student_options = df['Student Name'].tolist()
    selected_students = st.multiselect(
        "Select students to download PDF",
        options=student_options,
        default=student_options  # default: all selected
    )
    
    if selected_students:
        # Create a BytesIO object to store ZIP in memory
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for student in selected_students:
                row = df[df['Student Name'] == student].iloc[0]  # get student row
    
                # --- 5. Generate PDF ---
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
    
                # --- 6. Save PDF to memory ---
                pdf_bytes = BytesIO()
                pdf_output = pdf.output(dest='S').encode('latin-1')
                pdf_bytes.write(pdf_output)
                pdf_bytes.seek(0)
    
                # --- 7. Add PDF to ZIP ---
                zip_file.writestr(f"{row['Student Name']}_report.pdf", pdf_bytes.read())
    
        zip_buffer.seek(0)
        st.download_button(
            "📦 Download PDFs for Selected Students (ZIP)",
            data=zip_buffer,
            file_name="student_reports.zip"
        )


import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from io import BytesIO

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

    # --- 4. PDF Download per student ---
    for idx, row in df.iterrows():
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
    
        st.download_button(
            f"Download PDF for {row['Student Name']}",
            data=pdf_bytes,
            file_name=f"{row['Student Name']}_report.pdf"
        )


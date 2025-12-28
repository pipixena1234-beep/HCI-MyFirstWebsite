import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import zipfile
import json
import base64
import requests
import altair as alt
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
import openpyxl
from openpyxl.styles import Font

# 1. MOVE THIS TO THE TOP (Right after imports)
month_order = ["Jan", "Feb", "March", "Apr", "May", "June", "July", "August", "Sept", "Oct", "Nov", "Dec"]

# =====================================
# 1. Page Configuration
# =====================================
st.set_page_config(page_title="Student Progress System", layout="wide")
st.title("📊 Student Progress Report System")

# =====================================
# 2. GitHub Automation Sidebar
# =====================================
def push_schedule_to_github(new_datetime_str):
    repo_name = "pipixena1234-beep/HCI-MyFirstWebsite" 
    file_path = "schedule.json"
    try:
        token = st.secrets["GITHUB_TOKEN"]
        url = f"https://api.github.com/repos/{repo_name}/contents/{file_path}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        get_res = requests.get(url, headers=headers)
        sha = get_res.json().get("sha") if get_res.status_code == 200 else None
        content_dict = {"target_datetime": new_datetime_str}
        encoded_content = base64.b64encode(json.dumps(content_dict, indent=4).encode()).decode()
        payload = {"message": f"Update schedule to {new_datetime_str}", "content": encoded_content}
        if sha: payload["sha"] = sha
        return requests.put(url, json=payload, headers=headers).status_code
    except Exception as e: return str(e)

st.sidebar.subheader("⏰ Automation Schedule")
date_pick = st.sidebar.date_input("Target Date", datetime.now())
time_pick = st.sidebar.time_input("Target Time", datetime.now())
target_str = f"{date_pick.strftime('%Y-%m-%d')} {time_pick.strftime('%H:%M')}"

if st.sidebar.button("Update GitHub Schedule"):
    with st.spinner("Pushing..."):
        status = push_schedule_to_github(target_str)
        if status in [200, 201]: st.sidebar.success("✅ GitHub Updated!")
        else: st.sidebar.error(f"❌ Error: {status}")

# =====================================
# 3. Data Parsing Engine
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
                row_data = dict(zip(headers, df_raw.iloc[j].tolist()))
                row_data["Term"] = term
                rows.append(row_data)
                j += 1
            i = j
        else: i += 1
    return pd.DataFrame(rows)

# =====================================
# 4. File Upload & State Management
# =====================================
uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    selected_sheet = st.selectbox("Select Subject", xls.sheet_names)
    state_key = f"df_{selected_sheet}"

    # Initialize Session State if new sheet or file
    if state_key not in st.session_state:
        df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
        st.session_state[state_key] = extract_and_flatten(df_raw)

    # Use the Master Data from State
    master_df = st.session_state[state_key]

    # --- Term Selection Filter ---
    st.subheader("📅 Term Selection")
    all_terms = sorted(master_df["Term"].unique())
    select_terms_active = st.checkbox("Filter by specific terms")

    if select_terms_active:
        options = ["Select All"] + all_terms
        selected_terms = st.multiselect("Terms to process:", options, default=all_terms)
        if "Select All" in selected_terms: selected_terms = all_terms
    else:
        selected_terms = all_terms

    # Filtered Data for calculations
    df = master_df[master_df["Term"].isin(selected_terms)].copy()
    skills = ["Logic", "UI", "Animation", "Teamwork"]
    for s in skills: df[s] = pd.to_numeric(df[s], errors='coerce')

    # =====================================
    # 5. Data Editor (The Fix & Validation)
    # =====================================
    st.header(f"✏️ Data Review & Editor – {selected_sheet}")
    
    # Audit: Flagging Nulls
    audit_df = df.copy()
    audit_df.insert(0, "Status", audit_df.apply(lambda r: "🚨 MISSING" if r[skills].isnull().any() else "✅ OK", axis=1))
    
    show_nulls = st.checkbox("🔍 View only rows with 🚨")
    display_df = audit_df[audit_df["Status"] == "🚨 MISSING"] if show_nulls else audit_df

    edited_df = st.data_editor(display_df, num_rows="dynamic", use_container_width=True, key=f"editor_{state_key}")

    col_save, col_dl = st.columns(2)
    with col_save:
        if st.button("🔄 Apply Edits to Dashboard", use_container_width=True):
            cleaned_edits = edited_df.drop(columns=["Status"])
            st.session_state[state_key].update(cleaned_edits)
            st.success("Changes Saved to Session!")
            st.rerun()

    # 2. UPDATED FUNCTION
    def save_to_stacked_format(df_to_save, sheet_name, skill_cols):
        """Reconstructs the original Excel layout with Terms and dotted lines."""
        output = BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        
        current_row = 1
        
        # Sort terms using the global month_order
        terms = sorted(
            df_to_save["Term"].unique(), 
            key=lambda x: month_order.index(x) if x in month_order else 99
        )
        
        for term in terms:
            # Write Term Header
            ws.cell(row=current_row, column=1, value=f"Term: {term}").font = Font(bold=True)
            current_row += 1
            
            # Write Dotted Line
            ws.cell(row=current_row, column=1, value="--------------------------")
            current_row += 1
            
            # Write Table Headers
            headers = ["Student Name"] + skill_cols
            for col_num, header in enumerate(headers, 1):
                ws.cell(row=current_row, column=col_num, value=header).font = Font(bold=True)
            current_row += 1
            
            # Write Student Data
            term_data = df_to_save[df_to_save["Term"] == term]
            for _, row in term_data.iterrows():
                ws.cell(row=current_row, column=1, value=row["Student Name"])
                for col_num, skill in enumerate(skill_cols, 2):
                    # We use .get() or index to ensure we match the right column
                    ws.cell(row=current_row, column=col_num, value=row[skill])
                current_row += 1
            
            # Add space between tables
            current_row += 2
            
        wb.save(output)
        return output.getvalue()
    
    # 3. UPDATED CALL (Inside your UI)
    with col_dl:
        # Pass the variables explicitly to avoid NameErrors
        stacked_excel_data = save_to_stacked_format(
            st.session_state[state_key], 
            selected_sheet, 
            skills
        )
        
        st.download_button(
            label="📥 Download Corrected Excel (Original Format)",
            data=stacked_excel_data,
            file_name=f"Corrected_{selected_sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # =====================================
    # 6. Calculations & Metrics
    # =====================================
    df["Average"] = df[skills].mean(axis=1)
    df["Grade"] = df["Average"].apply(lambda x: "A" if x>=80 else "B" if x>=70 else "C" if x>=60 else "D" if x>=50 else "F")
    
    st.divider()
    st.subheader("📌 Performance Highlights")
    m1, m2, m3 = st.columns(3)

    if not df.empty:
        # Metric 1: Top Student
        top_row = df.loc[df['Average'].idxmax()]
        m1.metric("🏆 Top Student", top_row['Student Name'], f"{top_row['Average']:.1f} Avg")

        # Metric 2: Growth Calculations
        df_melted = df.melt(id_vars=['Term'], value_vars=skills, var_name='Skill', value_name='Score')
        df_grouped = df_melted.groupby(['Term', 'Skill'])['Score'].mean().reset_index()
        
        if len(selected_terms) > 0:
            base_t = selected_terms[0]
            df_base = df_grouped[df_grouped['Term'] == base_t][['Skill', 'Score']].rename(columns={'Score':'Base'})
            df_final = pd.merge(df_grouped, df_base, on='Skill')
            df_final['Growth'] = ((df_final['Score'] - df_final['Base']) / df_final['Base']) * 100
            
            latest_g = df_final[df_final['Term'] == selected_terms[-1]]
            if not latest_g.empty:
                best_s = latest_g.loc[latest_g['Growth'].idxmax()]
                m2.metric("📈 Most Improved", best_s['Skill'], f"{best_s['Growth']:.1f}% Growth")

        # Metric 3: Success Rate
        rate = (len(df[df['Grade'].isin(['A', 'B'])]) / len(df)) * 100
        m3.metric("🎯 Class Success Rate", f"{rate:.0f}%", "Grades A & B")

        # =====================================
        # 7. Growth Chart (Dual Axis)
        # =====================================
        st.header("📊 Performance & Growth Trends")
        
        base_chart = alt.Chart(df_final).encode(x=alt.X('Term:N', sort=month_order))
        
        bars = base_chart.mark_bar(opacity=0.5).encode(
            xOffset='Skill:N',
            y=alt.Y('Score:Q', scale=alt.Scale(domain=[0, 100])),
            color='Skill:N'
        )
        
        lines = base_chart.mark_line(size=3, point=True).encode(
            y=alt.Y('Growth:Q', title="Growth %", axis=alt.Axis(format='+')),
            color='Skill:N',
            tooltip=['Term', 'Skill', 'Score', 'Growth']
        )
        
        st.altair_chart(alt.layer(bars, lines).resolve_scale(y='independent').properties(height=500), use_container_width=True)

    # =====================================
    # 8. Report Export & Google Drive
    # =====================================
    st.divider()
    col_zip, col_drive = st.columns(2)
    
    with col_zip:
        if st.button("📦 Generate Student PDF ZIP"):
            z_buf = BytesIO()
            with zipfile.ZipFile(z_buf, "w") as zf:
                for _, r in df.iterrows():
                    pdf = FPDF()
                    pdf.add_page(); pdf.set_font("Arial", "B", 16)
                    pdf.cell(0, 10, f"Report: {r['Student Name']}", ln=True)
                    pdf.set_font("Arial", "", 12)
                    pdf.cell(0, 8, f"Term: {r['Term']} | Avg: {r['Average']:.1f} | Grade: {r['Grade']}", ln=True)
                    zf.writestr(f"{r['Term']}/{r['Student Name']}_report.pdf", pdf.output(dest="S").encode("latin-1"))
            st.download_button("⬇️ Download ZIP", z_buf.getvalue(), "student_reports.zip")

    with col_drive:
        f_id = st.text_input("G-Drive Folder ID", "0ALncbMfl-gjdUk9PVA")
        if st.button("🚀 Upload Current View to Drive"):
            st.info("Initiating Google Drive Upload...")
            # (Your service account logic here)

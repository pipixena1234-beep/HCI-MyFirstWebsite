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
from datetime import datetime
import requests  # <--- THIS WAS MISSING
import base64    # <--- ALSO NEEDED FOR THE GITHUB API
import altair as alt

st.set_page_config(page_title="Student Progress Reports", layout="wide")
st.title("📊 Student Progress Report System")

# =====================================
# Report Generation Automation - Date Selection
# =====================================

def push_schedule_to_github(new_datetime_str):
    # Update this with your actual GitHub username and repo name
    repo_name = "pipixena1234-beep/HCI-MyFirstWebsite" 
    file_path = "schedule.json"
    
    try:
        token = st.secrets["GITHUB_TOKEN"]
    except KeyError:
        st.error("GITHUB_TOKEN not found in Streamlit Secrets!")
        return 500

    url = f"https://api.github.com/repos/{repo_name}/contents/{file_path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 1. Get the current file (to get the SHA)
    get_res = requests.get(url, headers=headers)
    sha = None
    
    if get_res.status_code == 200:
        sha = get_res.json().get("sha")
    elif get_res.status_code == 404:
        # This is okay! It means the file doesn't exist yet.
        st.info("File not found on GitHub. Creating it for the first time...")
        sha = None 
    else:
        st.error(f"GitHub API Error: {get_res.status_code}")
        return get_res.status_code

    # 2. Prepare content
    content_dict = {"target_datetime": new_datetime_str}
    json_string = json.dumps(content_dict, indent=4)
    encoded_content = base64.b64encode(json_string.encode()).decode()

    # 3. Push to GitHub
    payload = {
        "message": f"Update schedule to {new_datetime_str} [skip ci]",
        "content": encoded_content
    }
    if sha:
        payload["sha"] = sha

    put_res = requests.put(url, json=payload, headers=headers)
    return put_res.status_code

# --- In your Sidebar UI ---
st.sidebar.subheader("⏰ Automation Schedule")
date_pick = st.sidebar.date_input("Target Date", datetime.now())
time_pick = st.sidebar.time_input("Target Time", datetime.now())

target_str = f"{date_pick.strftime('%Y-%m-%d')} {time_pick.strftime('%H:%M')}"

if st.sidebar.button("Update GitHub Schedule"):
    with st.spinner("Pushing to GitHub..."):
        status = push_schedule_to_github(target_str)
        if status in [200, 201]:
            st.sidebar.success(f"✅ GitHub updated to {target_str}!")
        else:
            st.sidebar.error(f"❌ Failed to update. Error code: {status}")


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
    
    # =========================
    # Data Quality & Instant Edit
    # =========================
    st.header(f"✏️ Data Review & Editor – {selected_sheet}")
    
    # 1. Count total nulls
    total_nulls = df.isnull().sum().sum()
    
    # 2. Create the Check Button/Toggle
    show_nulls_only = st.checkbox(f"🔍 Show only rows with missing data ({total_nulls} found)")
    
    # 3. Filter the dataframe based on the button
    if show_nulls_only:
        # This filters for rows where ANY value in your skill columns is null
        display_df = df[df[skills].isnull().any(axis=1)]
        if display_df.empty:
            st.success("✨ No missing values found in the skill columns!")
            display_df = df # Fallback to show all if none found
    else:
        display_df = df
    
    # 4. The Instant Editor (using the filtered display_df)
    st.info("💡 You can edit cells directly below. Changes will reflect in your charts and reports.")
    edited_df = st.data_editor(
        display_df, 
        num_rows="dynamic",
        use_container_width=True,
        key="editor"
    )
    
    # 5. Sync edits back to the main dataframe
    # If you edited the filtered view, we need to update the main 'df'
    if show_nulls_only:
        df.update(edited_df)
    else:
        df = edited_df

    # =========================
    # High-Level Metrics
    # =========================
    st.subheader("📌 Key Highlights")
    col_m1, col_m2, col_m3 = st.columns(3)

    # 1. Top Performing Student
    if not df.empty:
        # 1. Prepare Data
        df_melted = df.melt(id_vars=['Term'], value_vars=skills, var_name='Skill', value_name='TermScore')
        df_term_grouped = df_melted.groupby(['Term', 'Skill'])['TermScore'].mean().reset_index()
    
        # 2. Calculate Growth Percentage
        # Compares each term against the very first term available
        terms_sorted = sorted(df['Term'].unique())
        first_term = terms_sorted[0]
        df_first_values = df_term_grouped[df_term_grouped['Term'] == first_term][['Skill', 'TermScore']]
        df_first_values.rename(columns={'TermScore': 'BaselineScore'}, inplace=True)
        
        df_final = pd.merge(df_term_grouped, df_first_values, on='Skill')
        df_final['GrowthPct'] = ((df_final['TermScore'] - df_final['BaselineScore']) / df_final['BaselineScore']) * 100
        # Find the row with the highest average
        top_row = df.loc[df['Average'].idxmax()]
        top_student = top_row['Student Name']
        top_score = top_row['Average']
        
        col_m1.metric(
            label="🏆 Top Performing Student", 
            value=top_student, 
            delta=f"{top_score:.1f} Avg"
        )

        # 2. Most Improved Skill (Highest Growth)
        # Looking at the df_final we created for the chart
        if 'df_final' in locals() and not df_final.empty:
            best_growth_row = df_final.loc[df_final['GrowthPct'].idxmax()]
            col_m2.metric(
                label="📈 Most Improved Skill", 
                value=best_growth_row['Skill'], 
                delta=f"{best_growth_row['GrowthPct']:.1f}% Growth"
            )
        else:
            col_m2.metric("📈 Most Improved Skill", "Calculating...", delta=None)

        # 3. Overall Class Success Rate (Percentage of A & B grades)
        success_count = len(df[df['Grade'].isin(['A', 'B'])])
        success_rate = (success_count / len(df)) * 100
        col_m3.metric(
            label="🎯 Class Success Rate", 
            value=f"{success_rate:.0f}%", 
            delta="Grades A & B"
        )
    
    st.divider() # Adds a clean line between metrics and the data table

    st.header(f"📊 Integrated Performance & Growth Trend – {selected_sheet}")

    if not df.empty:
        # 3. Create the Base Chart
        # 1. Define the chronological order for sorting
        month_order = [
            "Jan", "Feb", "March", "Apr", "May", "June", 
            "July", "August", "Sept", "Oct", "Nov", "Dec"
        ]
        
        # 2. Update the Base Chart to use this order
        base = alt.Chart(df_final).encode(
            x=alt.X('Term:N', 
                    title='Academic Term', 
                    sort=month_order) # This forces the Jan -> Dec order
        )
            
        # 4. MULTIPLE BARS (Average Scores) - Primary Y-Axis (Left)
        bars = base.mark_bar(opacity=0.6).encode(
            xOffset='Skill:N',
            y=alt.Y('TermScore:Q', title='Average Score', scale=alt.Scale(domain=[0, 100])),
            color=alt.Color('Skill:N', legend=alt.Legend(title="Skills Performance", orient='top')),
            tooltip=['Term', 'Skill', 'TermScore']
        )
    
        # 5. MULTIPLE LINES + POINTS (Growth Trend) - Secondary Y-Axis (Right)
        # The right axis will auto-scale to fit your actual growth data
        lines = base.mark_line(size=3).encode(
            y=alt.Y('GrowthPct:Q', title='Growth %', axis=alt.Axis(titleColor='#ff4b4b', format='+')),
            color=alt.Color('Skill:N', legend=None), 
            tooltip=['Term', 'Skill', alt.Tooltip('GrowthPct:Q', format='.1f', title='Growth %')]
        )
    
        points = base.mark_point(size=50).encode(
            y='GrowthPct:Q',
            color='Skill:N'
        )
    
        # 6. Resolve Dual Axis and Combine
        combined_chart = alt.layer(bars, lines + points).resolve_scale(
            y='independent'
        ).properties(
            width='container',
            height=500,
            title="Skill Scores (Bars) vs. Growth Percentage (Lines)"
        ).interactive()
    
        st.altair_chart(combined_chart, use_container_width=True)
    
    else:
        st.warning("No data found to generate the combined chart.")
    
    # Create two columns for the buttons
    st.subheader("📤 Upload to Google Drive")
    folder_id_input = st.text_input(
        "Enter Google Drive Folder ID",
        value="0ALncbMfl-gjdUk9PVA"
    )
    col1, col2 = st.columns(2)

    # =========================
    # Generate ZIP (Column 1)
    # =========================
    with col1:
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
            
    # =========================
    # Upload to Google Drive (Column 2)
    # =========================
    with col2:
        if st.button("Upload to Google Drive"):
            try:
                sa_info = json.loads(st.secrets["google_service_account"]["google_service_account"])
                credentials = service_account.Credentials.from_service_account_info(
                    sa_info, scopes=['https://www.googleapis.com/auth/drive']
                )
                drive_service = build('drive', 'v3', credentials=credentials)
        
                # --- FIX: Calculate EXACT total steps ---
                # Only count rows that belong to the selected terms
                df_to_process = df[df["Term"].isin(selected_terms)]
                total_steps = len(df_to_process)
                
                if total_steps == 0:
                    st.warning("No data found for the selected terms.")
                else:
                    prog_bar = st.progress(0)
                    status_text = st.empty() # Placeholder for "Processing Alice Tan..."
                    current_step = 0
        
                    for term in selected_terms:
                        term_clean = term.strip()
                        parent_id = folder_id_input.strip()
        
                        # 1. FIND OR CREATE TERM FOLDER
                        query_folder = (
                            f"name='{term_clean}' "
                            f"and mimeType='application/vnd.google-apps.folder' "
                            f"and '{parent_id}' in parents and trashed=false"
                        )
                        folder_search = drive_service.files().list(
                            q=query_folder, fields="files(id)", 
                            includeItemsFromAllDrives=True, supportsAllDrives=True
                        ).execute()
                        
                        existing_folders = folder_search.get('files', [])
                        if existing_folders:
                            term_folder_id = existing_folders[0]['id']
                        else:
                            folder_metadata = {'name': term_clean, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
                            term_folder = drive_service.files().create(body=folder_metadata, fields='id', supportsAllDrives=True).execute()
                            term_folder_id = term_folder['id']
        
                        # 2. UPLOAD / OVERWRITE PDFs
                        df_term = df[df["Term"] == term]
                        
                        for _, row in df_term.iterrows():
                            student_name = row['Student Name'].strip()
                            file_name = f"{selected_sheet}_{student_name}_report.pdf"
                            
                            # Update status text
                            status_text.text(f"Uploading: {file_name}")
        
                            # --- Generate PDF ---
                            pdf = FPDF()
                            pdf.add_page()
                            pdf.set_font("Arial", "B", 16)
                            pdf.cell(0, 10, f"Progress Report ({selected_sheet}_{term_clean})", ln=True)
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
        
                            # --- Check if FILE already exists for Overwrite ---
                            query_file = f"name='{file_name}' and '{term_folder_id}' in parents and trashed=false"
                            file_search = drive_service.files().list(q=query_file, fields="files(id)", includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
                            existing_files = file_search.get('files', [])
        
                            if existing_files:
                                drive_service.files().update(fileId=existing_files[0]['id'], media_body=media, supportsAllDrives=True).execute()
                            else:
                                file_metadata = {'name': file_name, 'parents': [term_folder_id]}
                                drive_service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
        
                            # --- FIX: Increment progress based on ACTUAL processed rows ---
                            current_step += 1
                            prog_bar.progress(current_step / total_steps)
        
                    status_text.empty() # Clear status when done
                    st.success(f"✅ Successfully processed {total_steps} reports!")
        
            except Exception as e:
                st.error(f"Google Drive operation failed: {e}")
        

    
    

import os
import json
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
from io import BytesIO
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account

def run_automation():
    # --- 1. DATETIME CHECK ---
    if os.path.exists("schedule.json"):
        with open("schedule.json", "r") as f:
            config = json.load(f)
        
        target_str = config.get("target_datetime")
        # Convert string from JSON to a python datetime object
        target_dt = datetime.strptime(target_str, "%Y-%m-%d %H:%M")
        
        # Get Current Time (Adjusted for Malaysia UTC+8)
        # GitHub runners use UTC, so we add 8 hours to get your local time
        now_local = datetime.utcnow() + timedelta(hours=8)
        
        if now_local < target_dt:
            print(f"Current local time is {now_local.strftime('%Y-%m-%d %H:%M')}.")
            print(f"Target is {target_str}. Too early! Skipping.")
            return
        else:
            print(f"Time reached! Starting upload process...")
    else:
        print("No schedule.json found. Skipping.")
        return
        
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

def run_automation():
    # --- 1. DATE CHECK ---
    if os.path.exists("schedule.json"):
        with open("schedule.json", "r") as f:
            config = json.load(f)
        
        target_date = config.get("target_date")
        today = datetime.now().strftime("%Y-%m-%d")
        
        if today != target_date:
            print(f"Today ({today}) is not the scheduled date ({target_date}). Skipping.")
            return
    else:
        print("No schedule.json found. Skipping automation.")
        return

    # --- 2. SETUP & CREDENTIALS ---
    try:
        # Pulling from GitHub Secrets environment variable
        sa_info = json.loads(os.environ["GDRIVE_SERVICE_ACCOUNT"])
        folder_id = os.environ["GDRIVE_FOLDER_ID"]
        excel_path = "data.xlsx" # Ensure your excel file is in the repo or fetched
        
        credentials = service_account.Credentials.from_service_account_info(
            sa_info, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=credentials)

        # --- 3. DATA PROCESSING ---
        xls = pd.ExcelFile(excel_path)
        sheet_name = xls.sheet_names[0] # Adjust if you want specific sheets
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        df = extract_and_flatten(df_raw)

        skills = ["Logic", "UI", "Animation", "Teamwork"]
        df[skills] = df[skills].apply(pd.to_numeric)
        df["Average"] = df[skills].mean(axis=1)
        df["Grade"] = df["Average"].apply(lambda x: "A" if x>=80 else "B" if x>=70 else "C" if x>=60 else "D" if x>=50 else "F")
        df["Remarks"] = df["Average"].apply(lambda x: "Excellent!" if x>=80 else "Good effort!" if x>=70 else "Needs improvement")

        selected_terms = df["Term"].unique()

        # --- 4. UPLOAD LOGIC ---
        for term in selected_terms:
            term_clean = term.strip()
            
            # Find/Create Folder
            query = f"name='{term_clean}' and mimeType='application/vnd.google-apps.folder' and '{folder_id}' in parents and trashed=false"
            res = drive_service.files().list(q=query, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            files = res.get('files', [])
            
            if files:
                term_folder_id = files[0]['id']
            else:
                meta = {'name': term_clean, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [folder_id]}
                term_folder_id = drive_service.files().create(body=meta, fields='id', supportsAllDrives=True).execute()['id']

            df_term = df[df["Term"] == term]
            for _, row in df_term.iterrows():
                student_name = row['Student Name'].strip()
                file_name = f"{sheet_name}_{student_name}_report.pdf"
                
                # PDF Generation
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, f"Progress Report ({sheet_name}_{term_clean})", ln=True)
                pdf.set_font("Arial", "", 12)
                pdf.cell(0, 8, f"Student: {student_name}", ln=True)
                for s in skills:
                    pdf.cell(0, 8, f"{s}: {row[s]}", ln=True)
                pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)
                
                pdf_bytes = BytesIO()
                pdf_bytes.write(pdf.output(dest="S").encode("latin-1"))
                pdf_bytes.seek(0)
                media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf', resumable=True)

                # Overwrite Logic
                q_file = f"name='{file_name}' and '{term_folder_id}' in parents and trashed=false"
                f_res = drive_service.files().list(q=q_file, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
                exist_f = f_res.get('files', [])

                if exist_f:
                    drive_service.files().update(fileId=exist_f[0]['id'], media_body=media, supportsAllDrives=True).execute()
                    print(f"Updated: {file_name}")
                else:
                    drive_service.files().create(body={'name': file_name, 'parents': [term_folder_id]}, media_body=media, supportsAllDrives=True).execute()
                    print(f"Created: {file_name}")

        print("✅ Automation successful!")

    except Exception as e:
        print(f"❌ Automation failed: {e}")

if __name__ == "__main__":
    run_automation()

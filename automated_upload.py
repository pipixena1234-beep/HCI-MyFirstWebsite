import os
import json
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
from io import BytesIO
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account

# --- Helper: Flatten Logic (Same as your Streamlit version) ---
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

def main():
    # 1. Check Schedule
    if not os.path.exists("schedule.json"):
        print("❌ No schedule.json found. Exiting.")
        return

    with open("schedule.json", "r") as f:
        config = json.load(f)
    
    target_str = config.get("target_datetime") # Format: "2025-12-30 14:30"
    target_dt = datetime.strptime(target_str, "%Y-%m-%d %H:%M")
    
    # Malaysia Time (UTC+8) adjustment for GitHub Runners
    now_local = datetime.utcnow() + timedelta(hours=8)
    
    print(f"⏰ Current Local Time: {now_local.strftime('%Y-%m-%d %H:%M')}")
    print(f"🎯 Target Time: {target_str}")

    if now_local < target_dt:
        print("⏳ Time not reached yet. Skipping upload.")
        return

    print("🚀 Time reached! Starting Google Drive process...")

    # 2. Setup Google Drive
    sa_info = json.loads(os.environ["GDRIVE_SERVICE_ACCOUNT"])
    folder_id = os.environ["GDRIVE_FOLDER_ID"]
    data_file_id = os.environ["DATA_EXCEL_FILE_ID"]
    
    creds = service_account.Credentials.from_service_account_info(
        sa_info, scopes=['https://www.googleapis.com/auth/drive']
    )
    drive_service = build('drive', 'v3', credentials=creds)

    # --- NEW: Download data.xlsx from Drive instead of reading from Repo ---
    print("📥 Downloading data.xlsx from Google Drive...")
    request = drive_service.files().get_media(fileId=data_file_id)
    
    fh = BytesIO() # No 'io.' prefix needed anymore
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    
    fh.seek(0)
    
    # Process with pandas
    try:
        xls = pd.ExcelFile(fh, engine='openpyxl')
        sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        print(f"✅ Data loaded from sheet: {sheet_name}")
    except Exception as e:
        print(f"❌ Failed to read Excel: {e}")
        return
        
    df = extract_and_flatten(df_raw)

    skills = ["Logic", "UI", "Animation", "Teamwork"]
    df[skills] = df[skills].apply(pd.to_numeric)
    df["Average"] = df[skills].mean(axis=1)
    df["Grade"] = df["Average"].apply(lambda x: "A" if x>=80 else "B" if x>=70 else "C" if x>=60 else "D" if x>=50 else "F")
    df["Remarks"] = "Report Generated Automatically"

    # 4. Upload Logic
    for term in df["Term"].unique():
        term_clean = term.strip()
        
        # Folder Check/Create
        query = f"name='{term_clean}' and mimeType='application/vnd.google-apps.folder' and '{folder_id}' in parents and trashed=false"
        res = drive_service.files().list(q=query, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        folders = res.get('files', [])
        term_folder_id = folders[0]['id'] if folders else drive_service.files().create(
            body={'name': term_clean, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [folder_id]},
            fields='id', supportsAllDrives=True).execute()['id']

        df_term = df[df["Term"] == term]
        for _, row in df_term.iterrows():
            student_name = row['Student Name'].strip()
            file_name = f"{sheet_name}_{student_name}_report.pdf"
            
            # PDF Creation
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, f"Progress Report: {student_name}", ln=True)
            pdf.set_font("Arial", "", 12)
            for s in skills:
                pdf.cell(0, 8, f"{s}: {row[s]}", ln=True)
            pdf.cell(0, 8, f"Grade: {row['Grade']}", ln=True)

            pdf_bytes = BytesIO()
            pdf_bytes.write(pdf.output(dest="S").encode("latin-1"))
            pdf_bytes.seek(0)
            media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf', resumable=True)

            # Overwrite Check
            q_file = f"name='{file_name}' and '{term_folder_id}' in parents and trashed=false"
            f_res = drive_service.files().list(q=q_file, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            exist_f = f_res.get('files', [])

            if exist_f:
                drive_service.files().update(fileId=exist_f[0]['id'], media_body=media, supportsAllDrives=True).execute()
                print(f"✅ Updated: {file_name}")
            else:
                drive_service.files().create(body={'name': file_name, 'parents': [term_folder_id]}, media_body=media, supportsAllDrives=True).execute()
                print(f"🆕 Created: {file_name}")

if __name__ == "__main__":
    main()

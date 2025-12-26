import os
import json
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
from io import BytesIO
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account

# --- Helper: Flatten Logic ---
def extract_and_flatten(df_raw):
    rows = []
    i = 0
    while i < len(df_raw):
        cell = str(df_raw.iloc[i, 0])
        if cell.startswith("Term:"):
            term = cell.replace("Term:", "").strip()
            header_row = i + 2
            if header_row >= len(df_raw): break
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
        print("❌ No schedule.json found.")
        return

    with open("schedule.json", "r") as f:
        config = json.load(f)
    
    target_dt = datetime.strptime(config.get("target_datetime"), "%Y-%m-%d %H:%M")
    now_local = datetime.utcnow() + timedelta(hours=8)
    
    if now_local < target_dt:
        print(f"⏳ Waiting for {target_dt}. Current: {now_local}")
        return

    # 2. Setup Google Drive
    sa_info = json.loads(os.environ["GDRIVE_SERVICE_ACCOUNT"])
    root_folder_id = os.environ["GDRIVE_FOLDER_ID"]
    data_file_id = os.environ["DATA_EXCEL_FILE_ID"]
    
    creds = service_account.Credentials.from_service_account_info(
        sa_info, scopes=['https://www.googleapis.com/auth/drive']
    )
    drive_service = build('drive', 'v3', credentials=creds)

    # 3. Download data.xlsx
    request = drive_service.files().get_media(fileId=data_file_id)
    fh = BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    
    fh.seek(0)
    xls = pd.ExcelFile(fh, engine='openpyxl')
    skills = ["Logic", "UI", "Animation", "Teamwork"]

    # 4. Process Every Sheet
    for sheet_name in xls.sheet_names:
        print(f"📖 Subject: {sheet_name}")
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        df = extract_and_flatten(df_raw)

        if df.empty: continue

        df[skills] = df[skills].apply(pd.to_numeric, errors='coerce').fillna(0)
        df["Average"] = df[skills].mean(axis=1)

        # Optimization: Group by Term so we only check/create folders once per term
        for term_clean, group in df.groupby("Term"):
            term_clean = str(term_clean).strip()
            
            # --- FOLDER LOGIC (Runs once per Term) ---
            q_folder = f"name='{term_clean}' and mimeType='application/vnd.google-apps.folder' and '{root_folder_id}' in parents and trashed=false"
            res_folder = drive_service.files().list(q=q_folder, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            folders = res_folder.get('files', [])
            
            if folders:
                term_folder_id = folders[0]['id']
            else:
                print(f"📁 Creating new folder: {term_clean}")
                term_folder_id = drive_service.files().create(
                    body={'name': term_clean, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [root_folder_id]},
                    fields='id', supportsAllDrives=True).execute()['id']

            # --- STUDENT LOGIC ---
            for _, row in group.iterrows():
                student_name = str(row['Student Name']).strip()
                file_name = f"{sheet_name}_{student_name}_report.pdf"
                
                # PDF Creation
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, f"{sheet_name} - {student_name}", ln=True, align='C')
                pdf.ln(10)
                pdf.set_font("Arial", "", 12)
                for s in skills:
                    pdf.cell(0, 8, f"{s}: {int(row[s])}", ln=True)
                pdf.cell(0, 8, f"Average: {row['Average']:.2f}", ln=True)

                pdf_bytes = BytesIO(pdf.output(dest="S").encode("latin-1"))
                media = MediaIoBaseUpload(pdf_bytes, mimetype='application/pdf')

                # --- OVERWRITE CHECK ---
                q_file = f"name='{file_name}' and '{term_folder_id}' in parents and trashed=false"
                res_file = drive_service.files().list(q=q_file, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
                files = res_file.get('files', [])

                try:
                    if files:
                        drive_service.files().update(fileId=files[0]['id'], media_body=media, supportsAllDrives=True).execute()
                        print(f"✅ Overwritten: {file_name}")
                    else:
                        drive_service.files().create(body={'name': file_name, 'parents': [term_folder_id]}, media_body=media, supportsAllDrives=True).execute()
                        print(f"🆕 Created: {file_name}")
                except Exception as e:
                    print(f"⚠️ Small delay due to API limit, retrying {file_name}...")
                    import time
                    time.sleep(2) # Give the API a breather if it struggles
                    
if __name__ == "__main__":
    main()

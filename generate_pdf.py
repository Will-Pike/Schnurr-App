import os
import shutil
import tempfile
import requests
import gspread
import pdfkit
from PyPDF2 import PdfMerger
from jinja2 import Environment, FileSystemLoader
from oauth2client.service_account import ServiceAccountCredentials

# Uncomment this line for Linux or MacOS local environment:
WKHTMLTOPDF_PATH = "/usr/bin/wkhtmltopdf"
# For Windows local environment, use the following line instead:
# WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
SERVICE_FILE = "./service-account.json"
SPREADSHEET_ID = "16xuo0Uuyku5qD5Ul6VDO86I3rVSFzUedgVXMKfUv5CE"
PDFKIT_CONFIG = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

def generate_report_for_project(project):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    rows = sheet.get_all_records()

    env = Environment(loader=FileSystemLoader("templates"))
    template = env.get_template("report.html")

    pdf_files = []

    price_dict = load_price_dictionary()

    with tempfile.TemporaryDirectory() as temp_dir:
        for idx, row in enumerate(rows):
            if row.get("Project", "") != project:
                continue

            issue_type = row.get("Issue:", "")
            # Use price dictionary if available, else fallback to sheet value or "N/A"
            cost = price_dict.get(issue_type, row.get("Estimated Cost", "N/A"))

            record = {
                "project": row.get("Project", ""),  # <-- Add this line
                "obs_number": row.get("OBS ID#", "N/A"),
                "date": row.get("Timestamp", ""),
                "floor": row.get("Floor:", ""),
                "room": row.get("Room:", ""),
                "user": row.get("User:", ""),
                "description": issue_type,
                "responsible": row.get("Who is responsible?", ""),  # <-- Add this line
                "cost": cost,
                "photo_path": ""
            }

            photo_url = row.get("Upload photo:", "")
            temp_img_path = None
            if "drive.google.com" in photo_url and "id=" in photo_url:
                file_id = photo_url.split("id=")[-1]
                direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
                try:
                    response = requests.get(direct_url, stream=True, timeout=10)
                    response.raise_for_status()
                    temp_img_path = os.path.join(temp_dir, f"photo_{idx+1}.jpg")
                    with open(temp_img_path, 'wb') as f:
                        shutil.copyfileobj(response.raw, f)
                    record["photo_path"] = f"file:///{temp_img_path.replace(os.sep, '/')}"
                    # Only print if download succeeded
                    with open(temp_img_path, 'rb') as f:
                        print(f.read(10))
                except Exception as e:
                    print(f"⚠️ Image download failed for OBS {record['obs_number']}: {e}")
            
            html_content = template.render(records=[record])
            
            # Debug HTML
            with open(f"debug_{idx+1}.html", "w", encoding="utf-8") as f:
                f.write(html_content)

            output_file = os.path.join(temp_dir, f"report_{idx+1}.pdf")
            pdfkit.from_string(
                html_content, 
                output_file, 
                configuration=PDFKIT_CONFIG,
                options={'enable-local-file-access': None}
            )
            pdf_files.append(output_file)

        if not pdf_files:
            raise FileNotFoundError(f"No records found for project: {project}")

        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
        combined_pdf = f"combined_report_{project.replace(' ', '_')}.pdf"
        merger.write(combined_pdf)
        merger.close()

        # Save PDF to a known location, e.g.:
        output_path = f"/tmp/report_{project}.pdf"  # or a temp dir on Windows
        shutil.move(combined_pdf, output_path)  # Move the file to the desired location

        return output_path

def load_price_dictionary():
    PRICE_SHEET_ID = "1DBpjjmtaiUeGV_eeCwrihEOBrDk8aRKdDUQERFQBLRA"
    PRICE_SHEET_NAME = "Sheet1"
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(PRICE_SHEET_ID).worksheet(PRICE_SHEET_NAME)
    rows = sheet.get_all_records()
    # Build a dictionary: {issue_type: cost}
    price_dict = {row["Issue Type"]: row["Cost Estimate"] for row in rows}
    return price_dict

def get_report_record_count(project):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    rows = sheet.get_all_records()
    return sum(1 for row in rows if row.get("Project", "") == project)

def get_last_row_data():
    """Get the last row of data from the spreadsheet"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        rows = sheet.get_all_records()
        
        if rows:
            last_row = rows[-1]  # Get the last row
            return {
                "project": last_row.get("Project", ""),
                "user": last_row.get("User:", ""),
                "floor": last_row.get("Floor:", ""),
                "room": last_row.get("Room:", "")
            }
        return None
    except Exception as e:
        print(f"Error getting last row data: {e}")
        return None







import os
import shutil
import tempfile
import requests
import gspread
import pdfkit
from PyPDF2 import PdfMerger
from jinja2 import Environment, FileSystemLoader
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials
import io
from googleapiclient.http import MediaIoBaseUpload
import pickle
import os
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import io
from googleapiclient.http import MediaIoBaseUpload
import time

# Uncomment this line for Linux or MacOS local environment:
WKHTMLTOPDF_PATH = "/usr/bin/wkhtmltopdf"
# For Windows local environment, use the following line instead:
# WKHTMLTOPDF_PATH = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
SERVICE_FILE = "./service-account.json"
SPREADSHEET_ID = "16xuo0Uuyku5qD5Ul6VDO86I3rVSFzUedgVXMKfUv5CE"
PDFKIT_CONFIG = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

# Google Drive folder ID where photos should be uploaded
DRIVE_FOLDER_ID = "1J0vCCtKs2nBvL0cZFuye_FtGKjk7ZkEFZwgqtaHtRx_ygPItIuz5eiegm_FyWvQl866QR-bC"

# OAuth settings
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CLIENT_SECRETS_FILE = "client_secret.json"  # You'll need to create this
TOKEN_FILE = "token.pickle"

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

def get_highest_obs_for_project(project):
    """Get the highest OBS number for a specific project from the spreadsheet"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        rows = sheet.get_all_records()
        
        highest_obs = 0
        for row in rows:
            if row.get("Project", "") == project:
                obs_id = row.get("OBS ID#", "")
                # Extract number from OBS ID (assuming format like "1-A-001" or just "1")
                try:
                    # If it's a format like "1-A-001", extract the last part
                    if isinstance(obs_id, str) and '-' in obs_id:
                        obs_number = int(obs_id.split('-')[-1])
                    else:
                        obs_number = int(obs_id)
                    highest_obs = max(highest_obs, obs_number)
                except (ValueError, TypeError):
                    continue
        
        return highest_obs
    except Exception as e:
        print(f"Error getting highest OBS for project {project}: {e}")
        return 0

def get_obs_list_for_project(project):
    """Get list of all OBS entries for a specific project"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        rows = sheet.get_all_records()
        
        obs_list = []
        for idx, row in enumerate(rows):
            if row.get("Project", "") == project:
                obs_list.append({
                    "row_index": idx + 2,  # +2 because sheets are 1-indexed and we skip header
                    "obs_id": row.get("OBS ID#", ""),
                    "date": row.get("Timestamp", ""),
                    "floor": row.get("Floor:", ""),
                    "room": row.get("Room:", ""),
                    "issue": row.get("Issue:", "")
                })
        
        # Sort by OBS ID descending (newest first)
        obs_list.sort(key=lambda x: str(x.get("obs_id", "")), reverse=True)
        return obs_list
    except Exception as e:
        print(f"Error getting OBS list: {e}")
        return []

def get_obs_details(project, obs_id):
    """Get detailed information for a specific OBS entry"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        rows = sheet.get_all_records()
        
        for idx, row in enumerate(rows):
            if row.get("Project", "") == project and str(row.get("OBS ID#", "")) == str(obs_id):
                # Get photo URLs using the correct column name
                photo_url = row.get("Upload photo:", "")
                
                # Debug: Print what we found
                print(f"Found photo URL for OBS {obs_id}: '{photo_url}'")
                
                # Clean up photo URLs - handle various formats
                if photo_url:
                    # Remove any extra whitespace, newlines, and normalize separators
                    photo_url = photo_url.replace('\n', ',').replace('\r', ',').replace(';', ',')
                    photo_url = photo_url.strip()
                
                return {
                    "row_index": idx + 2,
                    "project": row.get("Project", ""),
                    "obs_id": row.get("OBS ID#", ""),
                    "timestamp": row.get("Timestamp", ""),
                    "user": row.get("User:", ""),
                    "floor": row.get("Floor:", ""),
                    "room": row.get("Room:", ""),
                    "issue": row.get("Issue:", ""),
                    "responsible": row.get("Who is responsible?", ""),
                    "photo_url": photo_url
                }
        return None
    except Exception as e:
        print(f"Error getting OBS details: {e}")
        return None

def update_obs_in_spreadsheet(project, obs_id, updated_data):
    """Update an OBS entry in the spreadsheet"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        
        # Get header row to map field names to column numbers
        headers = sheet.row_values(1)
        
        # Create a mapping of field names to column indices
        column_map = {}
        for i, header in enumerate(headers):
            column_map[header] = i + 1  # +1 because sheets are 1-indexed
        
        # Find the row to update
        rows = sheet.get_all_records()
        for idx, row in enumerate(rows):
            if row.get("Project", "") == project and str(row.get("OBS ID#", "")) == str(obs_id):
                row_index = idx + 2  # +2 because sheets are 1-indexed and we skip header
                
                # Update the fields using the column mapping
                if 'user' in updated_data and 'User:' in column_map:
                    sheet.update_cell(row_index, column_map['User:'], updated_data['user'])
                if 'floor' in updated_data and 'Floor:' in column_map:
                    sheet.update_cell(row_index, column_map['Floor:'], updated_data['floor'])
                if 'room' in updated_data and 'Room:' in column_map:
                    sheet.update_cell(row_index, column_map['Room:'], updated_data['room'])
                if 'issue' in updated_data and 'Issue:' in column_map:
                    sheet.update_cell(row_index, column_map['Issue:'], updated_data['issue'])
                if 'responsible' in updated_data and 'Who is responsible?' in column_map:
                    sheet.update_cell(row_index, column_map['Who is responsible?'], updated_data['responsible'])
                # Handle photo URL updates if needed
                if 'photo_urls' in updated_data and 'Upload photo:' in column_map:
                    sheet.update_cell(row_index, column_map['Upload photo:'], updated_data['photo_urls'])
                
                return True
        return False
    except Exception as e:
        print(f"Error updating OBS: {e}")
        return False

def debug_spreadsheet_data(project, obs_id):
    """Debug function to see raw spreadsheet data"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        
        # Get headers
        headers = sheet.row_values(1)
        
        # Get all rows
        rows = sheet.get_all_records()
        
        # Find the specific row
        for idx, row in enumerate(rows):
            if row.get("Project", "") == project and str(row.get("OBS ID#", "")) == str(obs_id):
                return {
                    "headers": headers,
                    "row_data": row,
                    "photo_column_names": [h for h in headers if 'photo' in h.lower() or 'upload' in h.lower()],
                    "photo_data": row.get("Upload photo:", "")
                }
        
        return {"error": "Row not found", "headers": headers}
    except Exception as e:
        return {"error": str(e)}

def get_oauth_drive_service():
    """Get an authenticated Drive service using OAuth"""
    creds = None
    
    # Load existing token
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, return None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            return None  # Need to run OAuth flow
    
    # Save the credentials for the next run
    with open(TOKEN_FILE, 'wb') as token:
        pickle.dump(creds, token)
    
    return build('drive', 'v3', credentials=creds)

def upload_photo_to_drive(file_data, filename, project, obs_id):
    """Upload a photo using OAuth (user's Drive quota)"""
    try:
        service = get_oauth_drive_service()
        if not service:
            return None
        
        # Create a unique filename
        timestamp = int(time.time())
        unique_filename = f"{project}_{obs_id}_{timestamp}_{filename}"
        
        # Detect file type
        mimetype = 'image/jpeg'
        if filename.lower().endswith('.png'):
            mimetype = 'image/png'
        elif filename.lower().endswith('.gif'):
            mimetype = 'image/gif'
        elif filename.lower().endswith('.webp'):
            mimetype = 'image/webp'
        
        # File metadata - use your original folder ID
        file_metadata = {
            'name': unique_filename,
            'parents': ['1J0vCCtKs2nBvL0cZFuye_FtGKjk7ZkEFZwgqtaHtRx_ygPItIuz5eiegm_FyWvQl866QR-bC']
        }
        
        # Create media upload
        media = MediaIoBaseUpload(
            io.BytesIO(file_data),
            mimetype=mimetype,
            resumable=True
        )
        
        # Upload the file
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        file_id = file.get('id')
        
        # Make the file publicly viewable
        permission = {
            'type': 'anyone',
            'role': 'reader'
        }
        service.permissions().create(
            fileId=file_id,
            body=permission
        ).execute()
        
        # Return the shareable URL
        shareable_url = f"https://drive.google.com/open?id={file_id}"
        
        print(f"Uploaded photo: {unique_filename} -> {shareable_url}")
        return shareable_url
        
    except Exception as e:
        print(f"Error uploading photo to Drive: {e}")
        return None

def add_photo_urls_to_obs(project, obs_id, new_photo_urls):
    """Add new photo URLs to an existing OBS entry"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        
        # Get header row to find the photo column
        headers = sheet.row_values(1)
        photo_column = None
        for i, header in enumerate(headers):
            if header == "Upload photo:":
                photo_column = i + 1  # +1 because sheets are 1-indexed
                break
        
        if not photo_column:
            print("Photo column 'Upload photo:' not found in spreadsheet")
            return False
        
        # Find the row to update
        rows = sheet.get_all_records()
        for idx, row in enumerate(rows):
            if row.get("Project", "") == project and str(row.get("OBS ID#", "")) == str(obs_id):
                row_index = idx + 2  # +2 because sheets are 1-indexed and we skip header
                
                # Get existing photo URLs
                existing_urls = row.get("Upload photo:", "")
                
                # Combine existing and new URLs
                if existing_urls.strip():
                    combined_urls = existing_urls + ", " + ", ".join(new_photo_urls)
                else:
                    combined_urls = ", ".join(new_photo_urls)
                
                # Update the cell
                sheet.update_cell(row_index, photo_column, combined_urls)
                
                print(f"Updated photo URLs for OBS {obs_id}: {combined_urls}")
                return True
        
        print(f"OBS {obs_id} not found for project {project}")
        return False
        
    except Exception as e:
        print(f"Error adding photo URLs to spreadsheet: {e}")
        return False

def remove_photo_url_from_obs(project, obs_id, photo_url_to_remove):
    """Remove a specific photo URL from an OBS entry"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        
        # Get header row to find the photo column
        headers = sheet.row_values(1)
        photo_column = None
        for i, header in enumerate(headers):
            if header == "Upload photo:":
                photo_column = i + 1  # +1 because sheets are 1-indexed
                break
        
        if not photo_column:
            print("Photo column 'Upload photo:' not found in spreadsheet")
            return False
        
        # Find the row to update
        rows = sheet.get_all_records()
        for idx, row in enumerate(rows):
            if row.get("Project", "") == project and str(row.get("OBS ID#", "")) == str(obs_id):
                row_index = idx + 2  # +2 because sheets are 1-indexed and we skip header
                
                # Get existing photo URLs
                existing_urls = row.get("Upload photo:", "")
                
                # Parse and filter out the URL to remove
                if existing_urls:
                    url_list = [url.strip() for url in existing_urls.split(',') if url.strip()]
                    filtered_urls = [url for url in url_list if url != photo_url_to_remove.strip()]
                    
                    # Update the cell with remaining URLs
                    updated_urls = ', '.join(filtered_urls) if filtered_urls else ''
                    sheet.update_cell(row_index, photo_column, updated_urls)
                    
                    print(f"Removed photo URL from OBS {obs_id}: {photo_url_to_remove}")
                    return True
                else:
                    print(f"No photos found for OBS {obs_id}")
                    return False
        
        print(f"OBS {obs_id} not found for project {project}")
        return False
        
    except Exception as e:
        print(f"Error removing photo URL from spreadsheet: {e}")
        return False

def delete_photo_from_drive(photo_url):
    """Optionally delete photo from Google Drive (use with caution)"""
    try:
        # Extract file ID from URL
        import re
        file_id_match = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', photo_url)
        if not file_id_match:
            print(f"Could not extract file ID from URL: {photo_url}")
            return False
        
        file_id = file_id_match.group(1)
        
        # Use OAuth to delete (requires drive scope)
        service = get_oauth_drive_service()
        if not service:
            print("Could not get OAuth Drive service")
            return False
        
        service.files().delete(fileId=file_id).execute()
        print(f"Deleted file from Drive: {file_id}")
        return True
        
    except Exception as e:
        print(f"Error deleting file from Drive: {e}")
        return False

def add_photos_to_pdf(pdf, photo_urls, page_width, margin):
    """Add multiple photos to PDF with proper layout"""
    if not photo_urls:
        return
    
    # Parse photo URLs
    urls = [url.strip() for url in photo_urls.split(',') if url.strip()]
    if not urls:
        return
    
    from reportlab.lib.units import inch
    from reportlab.platypus import Image
    import requests
    from io import BytesIO
    import tempfile
    import os
    
    # Calculate layout
    photos_per_row = 2 if len(urls) > 1 else 1
    photo_width = (page_width - 2 * margin - (photos_per_row - 1) * 0.25 * inch) / photos_per_row
    photo_height = photo_width * 0.75  # 4:3 aspect ratio
    
    current_x = margin
    current_y = pdf._y - photo_height - 0.25 * inch
    
    for i, url in enumerate(urls):
        try:
            # Convert Google Drive URL to downloadable format
            if 'drive.google.com' in url:
                # Extract file ID and create direct download URL
                import re
                file_id_match = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
                if file_id_match:
                    file_id = file_id_match.group(1)
                    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
                else:
                    print(f"Could not extract file ID from URL: {url}")
                    continue
            else:
                download_url = url
            
            # Download the image
            response = requests.get(download_url, timeout=30)
            if response.status_code == 200:
                # Create temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp_file:
                    temp_file.write(response.content)
                    temp_path = temp_file.name
                
                try:
                    # Add image to PDF
                    img = Image(temp_path, width=photo_width, height=photo_height)
                    img.drawOn(pdf._doc, current_x, current_y)
                    
                    # Add photo caption
                    pdf.setFont("Helvetica", 8)
                    pdf.drawString(current_x, current_y - 15, f"Photo {i + 1}")
                    
                    # Update position for next photo
                    if (i + 1) % photos_per_row == 0:
                        # Move to next row
                        current_x = margin
                        current_y -= photo_height + 0.5 * inch
                    else:
                        # Move to next column
                        current_x += photo_width + 0.25 * inch
                    
                finally:
                    # Clean up temporary file
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
            else:
                print(f"Failed to download image from {download_url}: {response.status_code}")
                
        except Exception as e:
            print(f"Error processing photo {url}: {e}")
            continue
    
    # Update PDF y position
    pdf._y = current_y - 0.25 * inch







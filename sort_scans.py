import os
import tempfile
import fitz
from PIL import Image
import pytesseract
import ollama

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

from pathvalidate import sanitize_filename

SCOPES = ['https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
PDF_MIME_TYPE = 'application/pdf'
LLM_MODEL_NAME = 'llama3:8b'
LLM_PROMPT_TEMPLATE = """The following text is a text of a letter or invoice obtained via OCR, likely in German or English.
Extract the sender and the short topic (3-5 words) from it, to use as a file name.
Avoid including any introductory or concluding remarks like 'The topic is:' or similar.

TEXT:
{text}
END OF TEXT

Remember, your task is to extract the sender and the short topic (3-5 words) from it, to use as a file name.
Respond ONLY in the format "SENDER - TOPIC", without any markup, comments or surrounding text.
"""
MAX_FILENAME_LENGTH = 150
MAX_OCR_TEXT_FOR_LLM = 4000


def authenticate_drive():
    """Google Drive API initialization wrapper"""
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)


def find_folder_id(service, folder_name):
    """Finds the ID of a folder given its name."""
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        print(f"Error: Folder '{folder_name}' not found.")
        return None
    return items[0]['id']


def perform_ocr_on_pdf(pdf_path):
    """OCR text from a PDF"""
    text = ''
    image_temp_files = []
    try:
        doc = fitz.open(pdf_path)
        print(f'Processing {doc.page_count} pages for OCR...')

        pages = []
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as img_temp:
                img_path = img_temp.name
                pix.save(img_path)
                image_temp_files.append(img_path)

            try:
                img = Image.open(img_path)
                page_text = pytesseract.image_to_string(img)
                pages.append(page_text)
                print(f'Page {page_num + 1}/{doc.page_count} processed.')
            except Exception as e:
                print(f'Warning: OCR failed for page {page_num + 1}: {e}')

        text = '\n'.join(pages)
        print('OCR complete.')
        doc.close()
    except Exception as e:
        print(f'An error occurred during PDF processing or OCR: {e}')
        return None
    finally:
        for img_path in image_temp_files:
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception as e:
                    print(f'Error cleaning up temporary image file {img_path}: {e}')

    return text.strip()


def get_topic_from_llm(text):
    """Sends text to the local LLM (Ollama) to extract a topic."""
    if not text:
        return None

    if len(text) > MAX_OCR_TEXT_FOR_LLM:
        print(f'Warning: Truncating text from {len(text)} to {MAX_OCR_TEXT_FOR_LLM} characters for LLM.')
        text = text[:MAX_OCR_TEXT_FOR_LLM]

    prompt = LLM_PROMPT_TEMPLATE.format(text=text)
    try:
        print(f'Sending text to LLM "{LLM_MODEL_NAME}"...')
        response = ollama.generate(model=LLM_MODEL_NAME, prompt=prompt, stream=False)
        if response and 'response' in response:
            topic = response['response'].strip()
            print(f'LLM Response: {topic}')
            return topic
        else:
            print('LLM returned an unexpected response format.')
            return None
    except Exception as e:
        print(f'An error occurred during LLM interaction: {e}')
        return None


def convert_topic_to_filename(topic):
    """Sanitizes a string to be safe for filenames and shortens it."""
    if not topic:
        return None
    sanitized = sanitize_filename(topic)
    sanitized = sanitized.replace(' ', '_')
    sanitized = sanitized.strip(' _-')
    sanitized = sanitized[:MAX_FILENAME_LENGTH]
    if not sanitized:
        return 'Untitled'
    return sanitized


def main():
    drive_service = authenticate_drive()
    if not drive_service:
        print("Authentication failed. Exiting.")
        return

    folder_name = input("Enter the name of the Google Drive folder containing PDFs: ")

    folder_id = find_folder_id(drive_service, folder_name)
    if not folder_id:
        print('Folder not found.')
        return

    print(f'Found folder "{folder_name}" with ID: {folder_id}')
    print(f'Listing PDF files ({PDF_MIME_TYPE}) in the folder...')

    query = f'"{folder_id}" in parents and mimeType="{PDF_MIME_TYPE}" and trashed=false'
    results = drive_service.files().list(
        q=query,
        fields="files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No PDF files found in the specified folder.')
        return

    print(f'Found {len(items)} PDF files. Processing...')

    for item in items:
        file_id = item['id']
        original_name = item['name']
        print(f'Processing file: "{original_name}" (ID: {file_id})')

        temp_pdf_path = None
        try:
            request = drive_service.files().get_media(fileId=file_id)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
                downloader = MediaIoBaseDownload(temp_file, request)
                done = False
                print('Downloading...')
                while done is False:
                    status, done = downloader.next_chunk()
                temp_pdf_path = temp_file.name
            print('Download complete.')

            ocr_text = perform_ocr_on_pdf(temp_pdf_path)
            if not ocr_text:
                print(f'Warning: Could not extract sufficient text.')
                continue

            topic = get_topic_from_llm(ocr_text)
            if not topic:
                print(f'Warning: Could not extract topic.')
                continue

            new_name_base = convert_topic_to_filename(topic)
            if not new_name_base:
                print(f'Warning: Sanitized filename is empty. Skipping rename.')
                continue

            file_extension = os.path.splitext(original_name)[1]
            new_name = f"{new_name_base}{file_extension}"
            print(f'New name: "{new_name}"')

            if new_name == original_name:
                continue
            print(f"Renaming '{original_name}' to '{new_name}' on Drive...")
            try:
                drive_service.files().update(
                    fileId=file_id,
                    body={'name': new_name},
                    fields='id, name'
                ).execute()
                print('Rename successful on Drive.')
            except Exception as e:
                print(f'Error renaming file on Drive: {e}')

        except Exception as e:
            print(f'An unexpected error occurred during processing file "{original_name}": {e}')
        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                    print(f'Cleaned up temporary PDF file: {temp_pdf_path}')
                except Exception as e:
                    print(f'Error cleaning up temporary PDF file {temp_pdf_path}: {e}')


if __name__ == '__main__':
    main()

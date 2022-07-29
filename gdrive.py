import io

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload



def copyDocument(creds, document_id, title):
    """Makes a copy of a document using its id
    The new document is in the same Google Drive Folder as the original
    Returns the id of the new document
    """

    try:
        drive_service = build('drive', 'v3', credentials=creds)
        body = {
            'name': title
        }
        drive_response = drive_service.files().copy(fileId=document_id, body=body).execute()
        document_copy_id = drive_response.get('id')
        return document_copy_id
    
    except HttpError as error:
        print('ERROR in gdrive.py > copyDocuments(): cannot connect to Google Drive')
        print(error)
        return



def makePDF(creds, document_id, title, folder_id):
    """Creates a PDF version of the provided Google Doc
    Uploads it to Google Drive
    Returns bytes of the PDF
    """

    try:
        PDF_MIME_TYPE = 'application/pdf'
        drive_service = build('drive', 'v3', credentials=creds)
        byte_string = drive_service.files().export(fileId=document_id, mimeType=PDF_MIME_TYPE).execute()
        
        media_object = MediaIoBaseUpload(io.BytesIO(byte_string), mimetype=PDF_MIME_TYPE)
        request = drive_service.files().create(media_body=media_object, body={
            'parents': [folder_id],
            'name': title + '.pdf'
        }).execute()

    except HttpError as error:
        print('ERROR in gdrive.py > makePDF(): cannot connect or upload to Google Drive')
        return

    return byte_string
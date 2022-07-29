from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



def getDocument(creds, document_id):
    """Retrieve the document object from its id
    Note: not used anymore
    """

    try:
        docs_service = build('docs', 'v1', credentials=creds)
        # Retrieve the documents' contents from the Docs service
        document = docs_service.documents().get(documentId=document_id).execute()
        return document

    except HttpError as error:
        print(f'ERROR in gdrive.py > getDocuments(): cannot connect or fetch the Google Doc')
        print(error)
        return




def fillDocument(creds, document_id, data):
    """Creates requests (json) to replace every {{}} keyword in the template
    Then sends the requests to the Docs service
    Note: data is a dictionnary
    """

    requests = []
    for key in data.keys():
        requests.append(
            {
                "replaceAllText": {
                    "replaceText": data[key],
                    "containsText": {
                        "text": key,
                        "matchCase": "true"
                    }
                }
            }
        )

    try:
        docs_service = build('docs', 'v1', credentials=creds)
        result = docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        return result # What's returned doesn't really matter
    
    except HttpError as error:
        print('ERROR in gdocs.py > fillDocument(): Cannot connect or update Google Doc')
        print(error)
        return
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
                'replaceAllText': {
                    'replaceText': data[key],
                    'containsText': {
                        'text': key,
                        'matchCase': 'true'
                    }
                }
            }
        )

    try:
        docs_service = build('docs', 'v1', credentials=creds)
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

        # Minor tweak for the square meter notation to display properly
        requests = [correctSuperscript(creds, document_id)]
        print(requests)
        docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
    
    except HttpError as error:
        print('ERROR in gdocs.py > fillDocument(): Cannot connect or update Google Doc')
        print(error)
    
    return



def correctSuperscript(creds, document_id):
    """Corrects the square meters notation in detail1_info
    """

    doc = getDocument(creds, document_id)
    json = doc['body']['content'][7]['table']['tableRows'][1]['tableCells'][0]['content'][0]['paragraph']['elements'][0]
    line = json['textRun']['content']

    # Get the index position of the square meters notation
    index = json['startIndex']
    for word in line.split(' '):
        index += len(word)
        if word[-1] == '2':
            break
        index += 1

    request = {
                    'updateTextStyle': {
                        'textStyle': {
                            'baselineOffset': 'SUPERSCRIPT'
                        },
                        'fields': 'baselineOffset',
                        'range': {
                            "startIndex": index - 1,
                            "endIndex": index,
                        }
                    }
                }
    return request
import gsheets, gauth, gdrive, gdocs, mail

# Step 0. Opens folder with relevant URLs, IDs, ...
misc = mail.getInfos('credentials/miscellaneous.txt')

# Step 1. Retrieves data from Google Sheets
print('Fetching data from Google Sheets...', end='\r')
data_raw = gsheets.getData(misc['sheet_url'], misc['worksheet_name'], 'credentials/credentials1.json')
data_completed = gsheets.calcMissingFields(data_raw)
data_formatted = gsheets.formatData(data_completed)
print('Fetching data from Google Sheets... Done')

# Step 2. Creates and fills each invoice
print('Connecting to Google Workspace\'s API...', end='\r')
credentials = gauth.getCredentials('credentials/token.json', 'credentials/credentials3.json')
print('Connecting to Google Workspace\'s API... Done')

for line in data_formatted:
    title = line['business_id'] + '-S'
    print(f'Working on {title}... Filling invoice', end='\r')

    doc_id = gdrive.copyDocument(credentials, misc['template_doc_id'], title)
    gdocs.fillDocument(credentials, doc_id, line)

    # Step 3. Makes a PDF copy in the Google Drive folder
    print('\033[K', end='\r') # Clears the line
    print(f'Working on {title}... Creating PDF', end='\r')
    attachment = {'title': title + '.pdf'}
    attachment['byte_string'] = gdrive.makePDF(credentials, doc_id, title, misc['folder_id'])

    # Step 4. Sends email to correspondant and puts a copy in the sender's 'Sent' inbox
    print('\033[K', end='\r') # Clears the line
    print(f'Working on {title}... Sending email', end='\r')
    message = mail.createMessage(misc['email_subject'], line, attachment)
    mail.send('credentials/mail_smtp.txt', message)
    mail.archiveCopy('credentials/mail_imaps.txt', title)
    print('\033[K', end='\r') # Clears the line
    print(f'Working on {title}... Done')
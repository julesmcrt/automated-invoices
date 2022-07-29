import smtplib, ssl, imaplib

from email.mime.application import MIMEApplication
from email.utils import formatdate
from email.message import EmailMessage



def createMessage(subject, details, attachment, content_path='templates/default/'):
    """Returns a MIMEMultipart object to be send
    Details is the mail's metadata (comes from the Google Sheet extraction)
    attachment is a dict with 2 keys: title (string) and byte_string (string)
    content_path is used for changing the email you want to send, for some edge cases
    Warning: the file specified in content_path must have 'mail.txt' and 'mail.html'
    """

    message = EmailMessage()
    # message['From'] is added when sending the email
    # message['Bcc] is also added when sending the email to have a back-up inside our mailbox
    message['To'] = details['send_to']
    message['Date'] = formatdate(localtime=True)
    message['Subject'] = subject

    # retrieves the mail content (core body of the email)
    # creates both a text (plain) version and a html version for html compatible mailboxes
    try:
        with open(content_path + 'mail.txt', 'r') as text_raw:
            text = text_raw.read()
            text = text.replace('BUSINESS_ID', details['business_id'])
            message.set_content(text)
    except FileNotFoundError as error:
        print(f'ERROR in mail.py > createMessage(): file mail.txt doesn\'t exist in given content_path: {content_path}')
        print(f'Mail for business: {details["business_id"]} was NOT sent')
        return
    
    try:
        with open(content_path + 'mail.html', 'r') as html_raw:
            html = html_raw.read()
            html = html.replace('BUSINESS_ID', details['business_id'])
            message.add_alternative(html, subtype='html')
    except FileNotFoundError as error:
        print(f'ERROR in mail.py > createMessage(): file mail.html doesn\'t exist in given content_path: {content_path}')
        print(f'Mail for business: {details["business_id"]} was still sent as a plain (.txt) version only')
        return

    # adds attachment to mail
    # Note: needs to be after the add_alternative() function
    ctype = 'application/pdf'
    maintype, subtype = ctype.split('/', 1)
    message.add_attachment(attachment['byte_string'], maintype=maintype, subtype=subtype, filename=attachment['title'])

    return message



def getInfos(credentials):
    """Formats the .txt file provided by the credentials (string) input (it's a filepath)
    Note: only used in private
    """

    with open(credentials, 'r') as file:
        lines = file.readlines()
        infos = {}
        for line in lines:
            key, value = line.split(' = ')
            if value[-1] == '\n':
                infos[key] = value[:-1]
            else:
                infos[key] = value
    
    return infos



def send(credentials, message):
    """Sends email using the credentials text file
    message is an EmailMessage object
    """

    infos = getInfos(credentials)

    message['From'] = infos['connexion_email']
    message['Bcc'] = infos['connexion_email']

    context = ssl.create_default_context()
    with smtplib.SMTP(infos['smtp_server'], infos['port']) as server:
        server.starttls(context=context)
        server.login(infos['connexion_email'], infos['password'])
        server.send_message(message)

    return



def archiveCopy(credentials, bill_title):
    """Saves the email to the right folder in the mailbox
    Note: the email must be in the 'INBOX'
    therefore the mail must have been sent to itself (the email sender's address)
    """

    infos = getInfos(credentials)

    with imaplib.IMAP4_SSL(infos['smtp_server'], infos['port']) as server:
        server.login(infos['connexion_email'], infos['password'])
        server.select('INBOX', readonly=False)
        
        # imap search syntax: spaces between arguments are AND
        search_request = f'(UNSEEN FROM \'{infos["connexion_email"]}\') TEXT \'{bill_title}\''
        status, data = server.search(None, search_request)
        mail_ids = data[0].split()
        if len(mail_ids) != 1:
            print('ERROR in mail.py > archiveCopy(): too many (or too few) emails found')
            return
        mail_id = mail_ids[0]
        
        server.store(mail_id, '+FLAGS', '\\Seen') # flags mail as seen
        server.copy(mail_id, '"TRESORERIE.2022.Factures de solde"')
        server.store(mail_id, '+FLAGS', '\\Deleted') # flags item as 'to be deleted later'
        server.expunge() # deletes flagged items
    
    return
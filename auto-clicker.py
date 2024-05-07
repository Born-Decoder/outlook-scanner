import win32com.client
from datetime import datetime, timedelta
import re
import webbrowser
import warnings
warnings.filterwarnings("ignore")

delta = int(input('DELTA MINS: ').strip() or 5)
sender = input('SENDER: ').strip() or "testemail@test.com"
subject = input('SUBJECT: ').strip() or 'test subject'
link_keyword = input('URL KEYWORD: ').strip() or 'https'
match_type = input('TYPE B for Best Match, E for Exact Match: ').strip() or 'E'
## customize to any browser as per needed
# edge_path = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
# webbrowser.register('edge', None, webbrowser.BackgroundBrowser(edge_path))
print(f'\n\n\nSUBJECT: {subject}\nSENDER: {sender}\nDELTA: {delta}\nLINK KEYWORD: {link_keyword}\nMATCH TYPE: {match_type}\n')

n = 30
dots = 1
lines = n - 1
while True:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    received_dt = datetime.now() - timedelta(minutes=delta)
    received_dt = received_dt.strftime('%d/%m/%Y %H:%M %p')
    unfiltered_emails = inbox.Items.Restrict(f"[ReceivedTime] >= '{received_dt}'")
    unfiltered_emails = unfiltered_emails.Restrict(f"[Subject] = '{subject}'")
    emails = []
    for email in unfiltered_emails:
        if sender in email.senderemailaddress:
            if match_type == 'B':
                emails.append(email)
            if sender == email.senderemailaddress:
                emails.append(email)
    
    print(f'Total emails with matching from {received_dt} are: {len(emails)}{"-" * dots}{"_" * lines}', end='\r')
    dots += 1
    lines -= 1
    if dots > n:
        dots = 1
        lines = n - 1
    if len(emails) == 0:
        continue
    email = emails[len(emails) - 1]
    email_content = email.body
    urls= re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', email_content)
    url_to_go = ''
    for url in urls:
        if link_keyword in url.lower():
            url_to_go = url[:-1]
            break
    # webbrowser.get('edge').open(url_to_go)
    webbrowser.open(url_to_go, new=0)
    exit()
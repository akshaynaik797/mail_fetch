import email
import imaplib

import pdfkit


ins = 'vipul'
inam = (1, 'inamdar hospital', 'mediclaim@inamdarhospital.org', 'Mediclaim@2019', 'imap.gmail.com', 'inbox', 'X')
max = (2, 'Max PPT', 'Tpappg@maxhealthcare.com', 'May@2020', 'outlook.office365.com', 'inbox', 'X')

server = 'outlook.office365.com'
user = 'Tpappg@maxhealthcare.com'
pwd = 'Sept@2020'
ic_email = 'cashless@vipulmedcorp.com'
fromtime = '19-Sep-2020'
totime = '20-Sep-2020'
ic_subject = 'Re: Pre-Authorization for MOHD NASIR : 361500501910003173'
mail = imaplib.IMAP4_SSL(server)
mail.login(user, pwd)
mail.select("inbox", readonly=True)

#
# type, data1 = mail.search(None,
#                               '(FROM ' + ic_email + ' since ' + fromtime + ' before ' + totime + ' (SUBJECT "%s"))' % ic_subject)


q = f'(since "{fromtime}" before "{totime}" (SUBJECT "{ic_subject}"))'
type, data1 = mail.search(None, q)

result, data = mail.fetch(b'40296', "(RFC822)")
# raw_email = data[0][1].decode('utf-8')
try:
    raw_email = data[0][1].decode('utf-8')
except UnicodeDecodeError:
    try:
        raw_email = data[0][1].decode('ISO-8859-1')
    except UnicodeDecodeError:
        try:
            raw_email = data[0][1].decode('ascii')
        except UnicodeDecodeError:
            pass
email_message = email.message_from_string(raw_email)
subject = email_message['Subject']
for mail.part in email_message.walk():
    if mail.part.get_content_type() == "text/html" or mail.part.get_content_type() == "text/plain":
        mail.body = mail.part.get_payload(decode=True)
        mail.file_name = ins + '/email.html'
        mail.output_file = open(mail.file_name, 'w')
        mail.output_file.write("Body: %s" % (mail.body.decode('utf-8')))
        mail.output_file.close()
        pdfkit.from_file(ins + '/email.html', ins+'/'+ ins + '.pdf')
print(type, mail.data, result, data)


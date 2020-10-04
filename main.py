import email
import imaplib
import os
import re
import random
import string
from datetime import datetime, timedelta

import pdfkit

from make_log import log_exceptions
# from config import mydb

folder = "files/"
if not os.path.exists(folder):
    os.mkdir(folder)

def get_from_query():
    try:
        q = """
              SELECT
                  claimNo AS ReferenceNo,
                  preauthNo AS PreAuthID,
                  Patient_Name AS PatientName,
                  Date_Of_Admission AS DateOfAdmision,
                  Date_Of_Discharge AS DateOfDischarge,
                  InsurerID AS InsurerTPAID,
                  status AS Status
              FROM
                  vnusoftw_vnuiclaimdb.claim
              WHERE
                  (status LIKE '%Sent To TPA/ Insurer%'
                      OR status LIKE '%In Progress%')
                      AND STR_TO_DATE(Date_of_Admission, '%d/%m/%Y') >= CURDATE() - 2
              UNION ALL SELECT
                  refno AS ReferenceNo,
                  preauthNo AS PreAuthID,
                  p_sname AS PatientName,
                  admission_date AS DateOfAdmision,
                  dischargedate AS DateOfDischarge,
                  insname AS InsurerTPAID,
                  status AS Status
              FROM
                  vnusoftw_vnuiclaimdb.preauth
              WHERE
                  (status LIKE '%Sent To TPA/ Insurer%'
                      OR status LIKE '%In Progress%')
                      AND STR_TO_DATE(admission_date, '%d/%m/%Y') >= CURDATE() - 2
        """

        # mycursor = mydb.cursor()
        # mycursor.execute(q)
        # result = mycursor.fetchall()
        result = [('MSS-1000293', '21CB08NIK0555', 'SATISH', '04/10/2020', '04/10/2020', '24', 'In Progress'),
                  ('MSS-1000306', '', 'NEHA  PUNDIR ', '02/10/2020', '03/10/2020', '16', 'Sent To TPA/ Insurer'),
                  ('MSS-1000248', '90222021332949', 'LALIT SINGH RAWAT', '03/10/2020', '04/10/2020', '15', 'Sent To TPA/ Insurer'),
                  ('MSS-1000269', 'NI-64-30680', 'SUNITA  BHATIJA ', '02/10/2020 12:46:50', '06/10/2020', '20', 'Sent To TPA/ Insurer'),
                  ('MSS-1000299', None, 'ALKA JAIN', '02/10/2020 10:20:55', '09/10/2020', '3', 'Sent To TPA/ Insurer'),
                  ('MSS-1000305', None, 'SAVITA VERMA ', '12/10/2020', '12/10/2020', '8', 'Sent To TPA/ Insurer'),
                  ('MSS-1000306', None, 'NEHA  PUNDIR ', '02/10/2020', '05/10/2020', '16', 'Sent To TPA/ Insurer'),
                  ('MSS-1000315', None, 'SANDEEP ARORA ', '04/10/2020', '06/10/2020', '3', 'Sent To TPA/ Insurer'),
                  ('MSS-1000320', None, 'LATA  NIWAS ', '02/10/2020 19:23:35', '05/10/2020', '15', 'Sent To TPA/ Insurer'),
                  ('MSS-1000326', 'RC-HS20-11366659', 'ALKA BALI ', '05/10/2020', '06/10/2020', 'I14', 'In Progress'),
                  ('MSS-1000330', 'CIG/2021/161116/0375652', 'SOHAN  SINGH ', '03/10/2020 17:26:35', '06/10/2020', 'I29', 'In Progress')]
        return result
    except:
        log_exceptions()


def download_pdf(hospital, subject):
    try:
        file_list, sender = [], ""
        fromtime = datetime.now().strftime("%d-%b-%Y")
        totime = datetime.now() + timedelta(days=1)
        totime = totime.strftime("%d-%b-%Y")

        server, email_id, password, inbox = "", "", "", ""
        if 'Max' in hospital:
            server, email_id, password, inbox = "outlook.office365.com", "Tpappg@maxhealthcare.com", "Sept@2020", '"Deleted Items"'
        elif 'inamdar' in hospital:
            server, email_id, password, inbox = "imap.gmail.com", "mediclaim@inamdarhospital.org", "Mediclaim@2019", '"[Gmail]/Trash"'
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_id, password)
        mail.select('inbox', readonly=True)
        subject = subject.replace("\r", "").replace("\n", "")
        type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{subject}"))')
        if data == [b'']:
            mail.select(inbox, readonly=True)
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{subject}"))')
        mid_list = data[0].split()
        if mid_list == []:
            return file_list, subject, sender
        result, data = mail.fetch(mid_list[-1], "(RFC822)")
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
        temp = re.compile(r"(?<=\<).*(?=\>)").search(email_message['From'])
        if temp is not None:
            sender = temp.group()
        for mail.part in email_message.walk():
            filename = mail.part.get_filename()
            if filename is not None:
                att_path = os.path.join(folder, filename)
                if not os.path.isfile(att_path):
                    fp = open(att_path, 'wb')
                    fp.write(mail.part.get_payload(decode=True))
                    file_list.append(att_path)
                    fp.close()
        return file_list, subject, sender
    except:
        log_exceptions(subject=subject)
        return file_list, subject, sender

def download_html(hospital, subject):
    try:
        file_list, sender = [], ""
        lowercase = string.ascii_lowercase
        filename = ''.join(random.choice(lowercase) for i in range(6))
        fromtime = datetime.now().strftime("%d-%b-%Y")
        totime = datetime.now() + timedelta(days=1)
        totime = totime.strftime("%d-%b-%Y")

        server, email_id, password, inbox = "", "", "", ""
        if 'Max' in hospital:
            server, email_id, password, inbox = "outlook.office365.com", "Tpappg@maxhealthcare.com", "Sept@2020", '"Deleted Items"'
        elif 'inamdar' in hospital:
            server, email_id, password, inbox = "imap.gmail.com", "mediclaim@inamdarhospital.org", "Mediclaim@2019", '"[Gmail]/Trash"'
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_id, password)
        mail.select('inbox', readonly=True)
        subject = subject.replace("\r", "").replace("\n", "")
        type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{subject}"))')
        if data == [b'']:
            mail.select(inbox, readonly=True)
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{subject}"))')
        mid_list = data[0].split()
        if mid_list == []:
            return file_list, subject, sender
        result, data = mail.fetch(mid_list[-1], "(RFC822)")
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
        temp = re.compile(r"(?<=\<).*(?=\>)").search(email_message['From'])
        if temp is not None:
            sender = temp.group()
        for mail.part in email_message.walk():
            if mail.part.get_content_type() == "text/html" or mail.part.get_content_type() == "text/plain":
                mail.body = mail.part.get_payload(decode=True)
                mail.file_name = folder + 'email.html'
                mail.output_file = open(mail.file_name, 'w')
                mail.output_file.write("Body: %s" % (mail.body.decode('utf-8')))
                mail.output_file.close()
                pdfkit.from_file(folder + 'email.html', folder + filename + '.pdf')
                file_list.append(folder + filename + '.pdf')
                if os.path.exists(folder + 'email.html'):
                    os.remove(folder + 'email.html')
        return file_list, subject, sender
    except:
        log_exceptions(subject=subject)
        return file_list, subject, sender

if __name__ == "__main__":
    a = get_from_query()
    records = []
    for i in a:
        f1 = download_pdf('Max', i[2])
        if not f1[0]:
            f1 = download_html('Max', i[2])
        records.append((i[0], i[2], f1))
    pass
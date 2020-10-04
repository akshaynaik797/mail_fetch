import email
import imaplib
import os
import re
import random
import sqlite3
import string
import subprocess
from datetime import datetime, timedelta

import pdfkit

from make_log import log_exceptions

# from config import mydb

folder, dbname = "files/", "database1.db"
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
                  ('MSS-1000248', '90222021332949', 'LALIT SINGH RAWAT', '03/10/2020', '04/10/2020', '15',
                   'Sent To TPA/ Insurer'),
                  ('MSS-1000269', 'NI-64-30680', 'SUNITA  BHATIJA ', '02/10/2020 12:46:50', '06/10/2020', '20',
                   'Sent To TPA/ Insurer'),
                  ('MSS-1000299', None, 'ALKA JAIN', '02/10/2020 10:20:55', '09/10/2020', '3', 'Sent To TPA/ Insurer'),
                  ('MSS-1000305', None, 'SAVITA VERMA ', '12/10/2020', '12/10/2020', '8', 'Sent To TPA/ Insurer'),
                  ('MSS-1000306', None, 'NEHA  PUNDIR ', '02/10/2020', '05/10/2020', '16', 'Sent To TPA/ Insurer'),
                  ('MSS-1000315', None, 'SANDEEP ARORA ', '04/10/2020', '06/10/2020', '3', 'Sent To TPA/ Insurer'),
                  ('MSS-1000320', None, 'LATA  NIWAS ', '02/10/2020 19:23:35', '05/10/2020', '15',
                   'Sent To TPA/ Insurer'),
                  ('MSS-1000326', 'RC-HS20-11366659', 'ALKA BALI ', '05/10/2020', '06/10/2020', 'I14', 'In Progress'),
                  (
                  'MSS-1000330', 'CIG/2021/161116/0375652', 'SOHAN  SINGH ', '03/10/2020 17:26:35', '06/10/2020', 'I29',
                  'In Progress')]
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


def get_insurer_and_process(subject, email_id):
    result = []
    try:
        q = f"select " \
            f"email_ids.IC, IC_name.IC_name, email_master.table_name, email_master.subject " \
            f"from email_ids " \
            f"inner join email_master on email_ids.IC=email_master.ic_id " \
            f"inner join IC_name on email_ids.IC=IC_name.IC " \
            f"where email_ids='{email_id}' and email_master.subject like '%{subject}%'"
        with sqlite3.connect(dbname) as con:
            cur = con.cursor()
            result = cur.execute(q).fetchall()
        return result
    except:
        log_exceptions(subject)
        return result


def get_run_no():
    run_no, q = 1, "select runno from updation_detail_log_copy order by runno desc limit 1;"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q).fetchone()
        if result is not None:
            return str(result[0]+1)
    return str(run_no)


def run_table_insert(run_no, subject, date, attach_path, email_id, completed):
    q = f"insert into run_table values ('{run_no}','{subject}','{date}','{attach_path}','{email_id}','{completed}')"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        cur.execute(q)
        return True


def run_table_get():
    q, result = "select * from run_table where completed = ''", []
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q).fetchall()
        return result

if __name__ == "__main__":
    # a = get_from_query()
    # b = get_insurer_and_process('Cashless Letter From Raksha Health Insurance TPA Pvt.Ltd. (N9014912431MAGICB,92000034200400000150,Alok Tyagi.)', "communication.abh@adityabirlacapital.com")
    # records = []
    # for i in a:
    #     f1 = download_pdf('Max', i[2])
    #     if not f1[0]:
    #         f1 = download_html('Max', i[2])
    #     records.append((i[0], i[2], f1))
    # pass
    #run no, subject, date, attach_path, email_id, completed
    #provide a function which will return table data with completed = blank
    #get run no from table before line 209 and pass into variable run_no
    run_no = get_run_no()
    run_table_insert(run_no, 'subject', 'date', 'attach_path', 'email_id', '')
    a = run_table_get()
    pass

    # a = ["python",'fhpl_query.py', '/home/akshay/PycharmProjects/trial -live/fhpl/attachments_pdf_query/1380120-0711_32182.pdf', run_no, 'fhpl', 'query', 'Preauthorization Request of Patient Name : Harsh Ghalott   , Patient Uhid No : NIC.21861308 and Status : Pending', '04/10/2020 16:45:07', 'Max PPT', '32182']
    # subprocess.run(a)
    # subprocess.run(["python", ins + "_" + ct + ".py", mail.filePath, str(run_no), ins, ct, subject, l_time, hid,
    #                 str(mail.latest_email_id)[2:-1]])

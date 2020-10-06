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
import requests
from dateutil import parser as date_parser

from cust_time_functs import ifutc_to_indian
from make_log import log_exceptions

# from config import mydb

folder, dbname = "files/", "database1.db"
if not os.path.exists(folder):
    os.mkdir(folder)


def get_from_query():
    result = []
    try:
        result = requests.get("http://3.7.8.68/api/get_from_query")
        if result.status_code == 200:
            return result.json()
    except:
        log_exceptions()
        return result


def download_pdf(hospital, subject):
    try:
        file_list, sender, l_time = [], "", ""
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
        l_time = email_message['Date']
        l_time = date_parser.parse(ifutc_to_indian(l_time)).strftime('%d/%m/%Y %H:%M:%S')
        temp = re.compile(r"(?<=\<).*(?=\>)").search(email_message['From'])
        if temp is not None:
            sender = temp.group()
        for mail.part in email_message.walk():
            filename = mail.part.get_filename()
            if filename is not None:
                #check for blacklist
                #if in blacklist
                #continue
                att_path = os.path.join(folder, filename)
                if not os.path.isfile(att_path):
                    fp = open(att_path, 'wb')
                    fp.write(mail.part.get_payload(decode=True))
                    file_list.append(att_path)
                    fp.close()
        return file_list, subject, sender, l_time
    except:
        log_exceptions(subject=subject)
        return file_list, subject, sender, l_time


def download_html(hospital, subject):
    try:
        file_list, sender, l_time = [], "", ""
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
            return file_list, subject, sender, l_time
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
        l_time = email_message['Date']
        l_time = date_parser.parse(ifutc_to_indian(l_time)).strftime('%d/%m/%Y %H:%M:%S')
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
        return file_list, subject, sender, l_time
    except:
        log_exceptions(subject=subject)
        return file_list, subject, sender, l_time


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
            return str(result[0] + 1)
    return str(run_no)


def run_table_insert(subject, date, attach_path, email_id, completed):
    q = f"insert into run_table (`subject`, `date`, `attachment_path`, `email_id`, `completed`) values ('{subject}','{date}','{attach_path}','{email_id}','{completed}')"
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
    a = get_from_query()
    records = []
    run_no = get_run_no()
    for i in a:
        f1 = download_pdf('Max', i[2])
        if not f1[0]:
            f1 = download_html('Max', i[2])
        records.append((i[0], i[2], f1))
    pass
    # for j in records:
    #     run_table_insert('subject', 'date', 'attach_path', 'email_id', '')
        # subprocess.run(["python", ins + "_" + ct + ".py", filepath, run_no, ins, ct, subject, l_time, hid])

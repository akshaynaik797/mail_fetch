import email
import imaplib
import os
import re
import random
import sqlite3
import string
from datetime import datetime, timedelta

import pdfkit
import requests
from dateutil import parser as date_parser

from cust_time_functs import ifutc_to_indian
from make_log import log_exceptions, custom_log_data


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


def get_mail_id_list(hospital, result):
    try:
        if hospital is None:
            hospital = "Max"
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
        mail_id_list = []
        mail.select('inbox', readonly=True)

        for i in result:
            p_name, temp_list = i[2], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (BODY "{p_name}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        for i in result:
            p_name, temp_list = i[2], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{p_name}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        for i in result:
            pre_id, temp_list = i[1], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (BODY "{pre_id}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        for i in result:
            pre_id, temp_list = i[1], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{pre_id}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)

        mail.select(inbox, readonly=True)
        for i in result:
            p_name, temp_list = i[2], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (BODY "{p_name}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        for i in result:
            p_name, temp_list = i[2], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{p_name}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        for i in result:
            pre_id, temp_list = i[1], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (BODY "{pre_id}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        for i in result:
            pre_id, temp_list = i[1], []
            type, data = mail.search(None, f'(since "{fromtime}" before "{totime}" (SUBJECT "{pre_id}"))')
            temp_list = data[0].split()
            mail_id_list.extend(temp_list)
        mail_id_list = sorted(list(set(mail_id_list)))
        return mail_id_list
    except:
        log_exceptions()
        return []

def download_pdf_and_html(hospital, mail_id_list):
    try:
        if hospital is None:
            hospital = "Max"
        file_name, sender, l_time = "", "", ""
        server, email_id, password, inbox = "", "", "", ""
        if 'Max' in hospital:
            server, email_id, password, inbox = "outlook.office365.com", "Tpappg@maxhealthcare.com", "Sept@2020", '"Deleted Items"'
        elif 'inamdar' in hospital:
            server, email_id, password, inbox = "imap.gmail.com", "mediclaim@inamdarhospital.org", "Mediclaim@2019", '"[Gmail]/Trash"'
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_id, password)
        delete_id_list = []

        mail.select('inbox', readonly=True)
        for mid in mail_id_list:
            try:
                result, data = mail.fetch(mid, "(RFC822)")
                if result != "OK":
                    delete_id_list.append(mid)
                    continue
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
                if check_if_sub_and_ltime_exist(subject, l_time):
                    continue
                # temp = re.compile(r"(?<=\<).*(?=\>)").search(email_message['From'])
                # if temp is not None:
                #     sender = temp.group()
                sender = email_message['From']
                filename = ""
                lowercase = string.ascii_lowercase
                date_tag = datetime.now().strftime("%Y-%m-%d-%H-%M-%S-%f_")
                for mail.part in email_message.walk():
                    filename = mail.part.get_filename()
                    if filename is not None:
                        #check for blacklist
                        #if in blacklist
                        #continue
                        if validate_filename(filename) is False:
                            continue
                        try:
                            att_path = ""
                            att_path = os.path.join(folder, date_tag+filename)
                        except TypeError:
                            zz = 1
                            att_path = os.path.join(folder, filename)
                        if not os.path.isfile(att_path):
                            fp = open(att_path, 'wb')
                            fp.write(mail.part.get_payload(decode=True))
                            file_name = att_path
                            fp.close()
                if file_name == "" and filename == "":
                    #code for html
                    lowercase = string.ascii_lowercase
                    filename = ''.join(random.choice(lowercase) for i in range(6))
                    email_message = email.message_from_string(raw_email)
                    for mail.part in email_message.walk():
                        if mail.part.get_content_type() == "text/html" or mail.part.get_content_type() == "text/plain":
                            mail.body = mail.part.get_payload(decode=True)
                            mail.file_name = folder + 'email.html'
                            mail.output_file = open(mail.file_name, 'w')
                            mail.output_file.write("Body: %s" % (mail.body.decode('utf-8')))
                            mail.output_file.close()
                            pdfkit.from_file(folder + 'email.html', folder + filename + '.pdf')
                            file_name = folder + filename + '.pdf'
                            if os.path.exists(folder + 'email.html'):
                                os.remove(folder + 'email.html')
                #insert file_name, subject, sender, l_time in run table
                run_table_insert(subject, l_time, file_name, sender, "")
            except:
                log_exceptions()


        mail.select(inbox, readonly=True)
        for mid in delete_id_list:
            result, data = mail.fetch(mid, "(RFC822)")
            if result != "OK":
                continue
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
            if check_if_sub_and_ltime_exist(subject, l_time):
                continue
            # temp = re.compile(r"(?<=\<).*(?=\>)").search(email_message['From'])
            # if temp is not None:
            #     sender = temp.group()
            sender = email_message['From']
            date_tag = datetime.now().strftime("%Y-%m-%d-%H-%M-%S-%f_")
            for mail.part in email_message.walk():
                filename = mail.part.get_filename()
                if filename is not None:
                    #check for blacklist
                    #if in blacklist
                    #continue
                    if validate_filename(filename) is False:
                        continue
                    att_path = os.path.join(folder, date_tag+filename)
                    if not os.path.isfile(att_path):
                        fp = open(att_path, 'wb')
                        fp.write(mail.part.get_payload(decode=True))
                        file_name = att_path
                        fp.close()
            if file_name == "":
                #code for html
                lowercase = string.ascii_lowercase
                filename = ''.join(random.choice(lowercase) for i in range(6))
                email_message = email.message_from_string(raw_email)
                for mail.part in email_message.walk():
                    if mail.part.get_content_type() == "text/html" or mail.part.get_content_type() == "text/plain":
                        mail.body = mail.part.get_payload(decode=True)
                        mail.file_name = folder + 'email.html'
                        mail.output_file = open(mail.file_name, 'w')
                        mail.output_file.write("Body: %s" % (mail.body.decode('utf-8')))
                        mail.output_file.close()
                        pdfkit.from_file(folder + 'email.html', folder + filename + '.pdf')
                        file_name = folder + filename + '.pdf'
                        if os.path.exists(folder + 'email.html'):
                            os.remove(folder + 'email.html')
            #insert file_name, subject, sender, l_time in run table
            run_table_insert(subject, l_time, file_name, sender, "")

        return True
    except:
        log_exceptions(mid)
        return False


def download_html(hospital, subject):
    try:
        file_name, sender, l_time = "", "", ""
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
            return file_name, subject, sender, l_time
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
                file_name = folder + filename + '.pdf'
                if os.path.exists(folder + 'email.html'):
                    os.remove(folder + 'email.html')
        return file_name, subject, sender, l_time
    except:
        log_exceptions(subject=subject)
        return file_name, subject, sender, l_time


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


def run_table_insert(subject, date, attach_path, email_id, completed, mail_id):
    if subject is not None:
        subject = subject.replace("'", "")
    q = f"insert into run_table (`subject`, `date`, `attachment_path`, `email_id`, `completed`, `mail_id`) values " \
        f"('{subject}','{date}','{attach_path}','{email_id}','{completed}','{mail_id}')"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        try:
            cur.execute(q)
            return True
        except:
            log_exceptions(query=q)
            return False


def get_details():
    datadict, result, fields = dict(), "", ["row_no", "subject", "date", "attachment", "email_id", "completed"]
    q = "select * from run_table where completed = ''"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q).fetchall()
        for i, j in enumerate(result):
            tempdict = {}
            for key, value in zip(fields, j):
                    tempdict[key] = value
            datadict[i] = tempdict
        return result


def post_details(row_no, flag):
    q = f"update run_table set completed = '{flag}' where row_no='{row_no}'"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q)


def check_if_sub_and_ltime_exist(subject, l_time):
    try:
        if subject is None:
            return False
        subject = subject.replace("'", '')
        with sqlite3.connect(dbname) as con:
            xyz = 10
            cur = con.cursor()
            b = f"select completed, mail_id from updation_detail_log where emailsubject='{subject}' and date='{l_time}'"
            cur.execute(b)
            r = cur.fetchone()
            if r is not None:
                if r[0] == None:
                    b = f"select completed, row_no from run_table where subject='{subject}' and date='{l_time}'"
                    cur.execute(b)
                    r1 = cur.fetchone()
                    if r1 is not None:
                        custom_log_data(filename="run_table", subject=subject, l_time=l_time, msg='found in run table')
                        return True
                    return False
                elif r[0] in ['x', 'X', "D"]:
                    custom_log_data(filename="updation_detail", subject=subject, l_time=l_time, flag=r[0], msg='found in table')
                    return True
            else:
                return False
                #check completed flag
                #if flag = null or blank
                # return false .. these entries will be passed to another function which will check if these entries exists in run_table ,
                # whichever entry is matched in run_table will be put in a different log file(run_table_log_file)
                #if flag = 'x' or 'X' or "D"
                #make log of subject and msg is {subject} is already processed
    except:
        log_exceptions()
        return False


def insert_entry_if_sub_and_ltime_exist(subject, l_time):
    q = f"update run_table set completed = 'Found in updation_detail_log' where subject='{subject}' and date='{l_time}'"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q)


def validate_filename(filename):
    valid = [r"^.*.pdf$", r"^.*.htm$", r"^.*.html$"]
    invalid = ['Gov']
    for i in valid:
        if not re.compile(i).fullmatch(filename):
            return False
        for j in invalid:
            if j in filename:
                return False
        return True


def process_row(row_no, ins, process, hospital):
    run_no = get_run_no()
    q = f"select attachement_path, subject, date from run_table where row_no='{row_no}'"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q)
    #subprocess.run(["python", ins + "_" + process + ".py", filepath, run_no, ins, process, subject, l_time, hosp_id])
    return 1


if __name__ == "__main__":
    a = get_from_query()
    if isinstance(a, dict):
        print(a)
    b = get_mail_id_list('Max', a)
    download_pdf_and_html('Max', b)
    pass
    # records = []
    # run_no = get_run_no()
    # for i in a:
    #     f1 = download_pdf('Max', i[2])
    #     if not f1[0]:
    #         f1 = download_html('Max', i[2])
    #     records.append((i[0], i[2], f1))
    # for j in records:
    #     subject, date, attach_path, email_id, completed = j[2][1], j[2][3], j[2][0], j[2][2], ''
        #insert in run_table if not exist in updation_Detail_log
        # run_table_insert(subject, date, attach_path, email_id, completed)
        # subprocess.run(["python", ins + "_" + ct + ".py", filepath, run_no, ins, ct, subject, l_time, hid])

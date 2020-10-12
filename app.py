import os
import sqlite3
import subprocess

import requests
import threading

from flask import Flask, request, jsonify, send_from_directory, abort
from apscheduler.schedulers.background import BackgroundScheduler

from functions import run_process, process_row, log_api_data
from make_log import log_exceptions
from settings import dbname, folder
from flask_cors import CORS, cross_origin
app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
scheduler = BackgroundScheduler()
state = "running"
sem = threading.Semaphore()


@app.route("/api/postUpdateDetailsLogs", methods=["POST"])
def postUpdateLog():
    preauthid = ''
    amount = ''
    status = ''
    lettertime = ''
    policyno = ''
    memberid = ''
    row_no = ''
    comment = ''
    completed = ''
    tagid = ''
    refno = ''

    if request.method != 'POST':
        return jsonify(
            {
                'status': 'failed',
                'message': 'inavlid request method.Only Post method Allowed'
            }
        )
    if request.form.get('row_no') != None:
        row_no = request.form['row_no']
    if request.form.get('completed') != None:
        completed = request.form['completed']  # completd = D
    if completed == 'D':
        with sqlite3.connect("database1.db") as con:
            cur = con.cursor()
            # query = f'update updation_detail_log set completed= "D" where row_no={row_no};'
            # print(query)
            # log_api_data('query', query)
            # cur.execute(query)
            apimessage = 'Record successfully updated, and API not called'
            return jsonify({
                'status': 'success',
                'message': apimessage})

    # with sqlite3.connect("database1.db") as con:
    #     cur = con.cursor()
    #     q = f'select preauthid,amount,status,process,lettertime,policyno,memberid,hos_id from updation_detail_log where row_no={row_no}'
    #     print(q)
    #     log_api_data('q', q)
    #     cur.execute(q)
    #     r = cur.fetchone()
    #     hosid = r[7]

    if request.form.get('preauthid') != None:
        preauthid = request.form['preauthid']

    if request.form.get('amount') != None:
        amount = request.form['amount']

    if request.form.get('status') != None:
        status = request.form['status']

    if request.form.get('lettertime') != None:
        lettertime = request.form['lettertime']

    if request.form.get('policyno') != None:
        policyno = request.form['policyno']

    if request.form.get('memberid') != None:
        memberid = request.form['memberid']

    if request.form.get('comment') != None:
        comment = request.form['comment']
    if request.form.get('tag_id') != None:
        tagid = request.form['tag_id']
    if request.form.get('refno') != None:
        refno = request.form['refno']
    # if (r is not None
    #         and preauthid == r[0]
    #         and amount == r[1]
    #         and status == r[2]
    #         # and process == r[3]
    #         and lettertime == r[4]
    #         and policyno == r[5]
    #         and memberid == r[6]):
    #     char = 'X'
    # else:
    #     char = 'x'
    char = 'X'
    if row_no == '':
        return jsonify(
            {
                'status': 'failed',
                'message': 'Parameter Field Are Empty'
            }
        )

    try:
        # query = "update updation_detail_log set"
        # flag = 0
        # if request.form.get('preauthid') != None:
        #     query = query + " preauthid='%s'" % preauthid
        #     flag = 1
        #
        # if request.form.get('amount') != None:
        #     if flag == 1:
        #         query = query + ", "
        #     query = query + " amount='%s'" % amount
        #     flag = 1
        #
        # if request.form.get('status') != None:
        #     if flag == 1:
        #         query = query + ", "
        #     query = query + " status='%s'" % status
        #     flag = 1
        #
        # if request.form.get('lettertime') != None:
        #     if flag == 1:
        #         query = query + ", "
        #     query = query + " lettertime='%s'" % lettertime
        #     flag = 1
        #
        # if request.form.get('comment') != None:
        #     if flag == 1:
        #         query = query + ", "
        #     query = query + " comment='%s'" % comment
        #     flag = 1
        #
        # if request.form.get('policyno') != None:
        #     if flag == 1:
        #         query = query + ", "
        #     query = query + " policyno='%s'" % policyno
        #     flag = 1
        #
        # if request.form.get('memberid') != None:
        #     if flag == 1:
        #         query = query + ", "
        #     query = query + " memberid='%s'" % memberid
        #     flag = 1
        #
        # if len(query) > len("update updation_detail_log set"):
        #     # query=query+", completed='X'"
        #     query = query + " where row_no=%s" % row_no
        #     print(query)
        #     log_api_data('query', query)
        #
        #     sem.acquire()
        #     print('Lock Acquired')
        #     con = sqlite3.connect("database1.db")
        #     cur = con.cursor()
        #     cur.execute(query)
        #     con.commit()
        #
        #     cur.close()
        #
        #     sem.release()
        #     print('Lock Released')
            # akshay code to call API............ first, fetch file_path from local db

            with sqlite3.connect("database1.db") as con:
                cur = con.cursor()
                q = f'select attachment_path from run_table where row_no={row_no};'
                print(q)
                log_api_data('q', q)
                cur.execute(q)
                r = cur.fetchone()
                if r:
                    r = r[0]
                else:
                    r = ''
            if r == None:
                apimessage = "Record updated in db, but API failed due to no File"
            else:
                print(row_no)
                log_api_data('row_no', row_no)
                print(r)
                log_api_data('r', r)
                files = {'doc': open(r, 'rb')}
                if refno[0] == 'M':
                    API_ENDPOINT = "https://vnusoftware.com/iclaimmax/api/preauth/"
                else:
                    API_ENDPOINT = "https://vnusoftware.com/iclaimportal/api/preauth"
                data = {
                    'preauthid': preauthid,
                    # 'pname': sys.argv[10],
                    'amount': amount,
                    'status': status,
                    'process': '',
                    'lettertime': lettertime,
                    'policyno': policyno,
                    'memberid': memberid,
                    'write': 'X',
                    'tagid': tagid,
                    'comment': comment,
                    'refno': refno,
                }

                r = requests.post(url=API_ENDPOINT, data=data, files=files)
                print(data)
                log_api_data('data', data)
                pastebin_url = r.text
                print(pastebin_url)
                log_api_data('pastebin_url', pastebin_url)
                if char == 'X':
                    query = f'update run_table set completed= "X" where row_no={row_no};'
                elif char == 'x':
                    query = f'update run_table set completed= "x" where row_no={row_no};'
                with sqlite3.connect("database1.db") as con:
                    cur = con.cursor()
                    cur.execute(query)
                if pastebin_url.find("Data Update Success") == -1:
                    apimessage = "Record updated in db, and API failed"
                    subprocess.run(["python", "sms_api.py", "api error"])
                else:
                    apimessage = 'Record successfully updated, and API successfully called'
                    # update completed flag in table
                #set completed = completed value where refno is found.
                with sqlite3.connect("database1.db") as con:
                    cur = con.cursor()
                    q = f"update run_table set completed = '{completed}' where ref_no = '{refno}'"
                    cur.execute(q)

            # if api call returns success message, then message = 'Record succ updated, and API succ called.
            # if not, then message = 'Record updated in db, but API failed.
            return jsonify({
                'status': 'success',
                'message': apimessage
            })
    except Exception as e:
        log_exceptions()
        sem.release()
        print(e.__str__())
        log_api_data('e.__str__()', e.__str__())
        return jsonify({
            'status': 'failure',
            'message': 'Record does not updated',
            'reason': e.__str__()
        })


@app.route("/api/getupdatedetailsLog", methods=["POST"])
def getupdatelog():
    if request.method != 'POST':
        return jsonify(
            {
                'status': 'failed',
                'message': 'inavlid request method.Only Post method Allowed'
            }
        )
    runno = ''
    if request.json.get('runno') != None:
        runno = request.json['runno']

    if runno == '':
        return jsonify(
            {
                'status': 'failed',
                'message': 'Parameter Field Are Empty'
            }
        )
    try:
        data = None
        con = sqlite3.connect("database1.db")
        cur = con.cursor()
        if runno == '00':
            query = """SELECT runno,insurerid,process,emailsubject,date,file_path,hos_id,preauthid,amount,status,lettertime,policyno,memberid,row_no,comment from updation_detail_log_copy WHERE completed is NULL and error is NULL and hos_id = 'inamdar hospital' """  # if runno = '0'->all

        elif runno != '0':
            query = """SELECT runno,insurerid,process,emailsubject,date,file_path,hos_id,preauthid,amount,status,lettertime,policyno,memberid,row_no,comment from updation_detail_log_copy WHERE completed is NULL and error is NULL and runno=%s""" % runno  # if runno = '0'->all
        else:
            query = """SELECT runno,insurerid,process,emailsubject,`date`,file_path,hos_id,preauthid,amount,`status`, \
          lettertime,policyno,memberid,row_no,comment from updation_detail_log_copy WHERE error IS NULL and completed is NULL"""

        print(query)
        # log_api_data('query', query)
        cur.execute(query)
        data = cur.fetchall()
        if data:
            myList = []
            for row in data:
                localDic = {}
                localDic['runno'] = row[0]
                localDic['insurerid'] = row[1]
                localDic['process'] = row[2]
                localDic['emailsubject'] = row[3]
                localDic['date'] = str(row[4])
                localDic['file_path'] = row[5]
                localDic['hos_id'] = row[6]
                localDic['preauthid'] = row[7]
                localDic['amount'] = row[8]
                localDic['status'] = row[9]
                localDic['lettertime'] = row[10]
                localDic['policyno'] = row[11]
                localDic['memberid'] = row[12]
                localDic['row_no'] = row[13]
                localDic['comment'] = row[14]

                url = request.url_root
                url = url + 'api/downloadfile?filename='
                url = url + str(row[5])
                localDic['file_path'] = url

                if localDic['memberid'] != None or localDic['preauthid'] != None or localDic['policyno'] != None or \
                        localDic['comment'] != None:
                    if localDic['hos_id'] == 'Max PPT':
                        url = 'https://vnusoftware.com/iclaimmax/api/preauth/vnupatientsearch'
                    else:
                        url = 'https://vnusoftware.com/iclaimportal/api/preauth/vnupatientsearch'
                    payload = {
                        'memberid': localDic['memberid'],
                        'preauthid': localDic['preauthid'],
                        'policyno': localDic['policyno'],
                        'comment': localDic['comment']
                    }

                    try:
                        temp = {}

                        for i, j in payload.items():
                            print(i, j)
                            if j == None:
                                temp[i] = ''
                            else:
                                temp[i] = j

                        payload = temp

                        response = requests.post(url, data=payload)
                        result = response.json()
                        print(result)
                        # log_api_data('result', result)
                        if result['status'] == '1':
                            localDic['searchdata'] = result['searchdata']
                        else:
                            localDic["searchdata"] = {
                                "refno": "",
                                "approve_amount": "",
                                "current_status": "",
                                "process": "",
                                "InsurerId": "",
                                "insname": "",
                                "Cumulative_flag": ""
                            }
                    except Exception as e:
                        log_exceptions(payload=payload, url=url, response=response)
                        print(e)
                        # log_api_data('e', e)
                        localDic["searchdata"] = {
                            "refno": "",
                            "approve_amount": "",
                            "current_status": "",
                            "process": "",
                            "InsurerId": "",
                            "insname": ""
                        }
                else:
                    localDic["searchdata"] = {
                        "refno": "",
                        "approve_amount": "",
                        "current_status": "",
                        "process": "",
                        "InsurerId": "",
                        "insname": ""
                    }

                myList.append(localDic)

                # localRespons['url']=url

                # x=row[13]
                # print(x)
                # localDic['apiparameter']=json.loads(row[13])
                # myList.append(localDic)
            # return render_template("detailview.html",updateList=myList)

            return jsonify({
                'status': 'pass',
                'data': (myList)
            })
        else:
            localDic = {}
            localDic['runno'] = ''
            localDic['insurerid'] = ''
            localDic['process'] = ''
            localDic['emailsubject'] = ''
            localDic['date'] = ''
            localDic['file_path'] = ''
            localDic['hos_id'] = ''
            localDic['preauthid'] = ''
            localDic['amount'] = ''
            localDic['status'] = ''
            localDic['lettertime'] = ''
            localDic['policyno'] = ''
            localDic['memberid'] = ''
            localDic['row_no'] = ''
            localDic['comment'] = ''
            return jsonify({'status': 'fail',
                            'data': (localDic)})


    except Exception as e:
        log_exceptions()
        print(e)
        # log_api_data('e', e)
        return jsonify({
            'status': 'failure',
            'reason': e.__str__()
        })


@app.route('/get_details')
def get_details():
    datadict = dict()
    datalist, result, fields = list(), "", ["row_no", "subject", "date", "attachment", "email_id", "completed",
                                            "mail_id", "p_name", "pre_id", "ref_no"]
    q = "select * from run_table where completed = ''"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q).fetchall()
        for i, j in enumerate(result):
            tempdict = {}
            for key, value in zip(fields, j):
                tempdict[key] = value
            datalist.append(tempdict)
            datalist[-1]['attachment'] = request.url_root + 'get_file/' + os.path.split(datalist[-1]['attachment'])[1]
        datadict['data'] = datalist
        return jsonify(datadict)


@app.route('/post_details', methods=["POST"])
def post_details():
    try:
        row_no, flag = "", ""
        if request.method == 'POST':
            a = request.json
            if request.json.get('row_no') != None:
                row_no = request.json['row_no']
            if request.json.get('flag') != None:
                flag = request.json['flag']
            q = f"update run_table set completed = '{flag}' where row_no='{row_no}'"
            with sqlite3.connect(dbname) as con:
                cur = con.cursor()
                cur.execute(q)
            return jsonify(True)
    except Exception as e:
        return jsonify(e)


@app.route('/process_records', methods=["POST"])
def process_records():
    row_no, ins, process, hospital = "", "", "", ""
    if request.method == 'POST':
        try:
            row_no, ins, process, hospital = request.json['row_no'], request.json['ins'], request.json['process'], \
                                             request.json['hospital'],
            # processing logic here
            process_row(row_no, ins, process, hospital)
            return 'True'
        except Exception as e:
            log_exceptions()
            return jsonify(e)


@app.route('/process_subject', methods=["POST"])
def process_subject():
    hospital, subject, ins, process, result = "", "", "", "", []
    if request.method == 'POST':
        try:
            subject, ins, process, hospital = request.json['subject'], request.json['ins'], request.json['process'], \
                                              request.json['hospital']
            q = f"select row_no from run_table where subject='{subject}'"
            with sqlite3.connect(dbname) as con:
                cur = con.cursor()
                result = cur.execute(q).fetchall()
            for i in result:
                # processing logic here
                process_row(i[0], ins, process, hospital)
        except:
            log_exceptions()
    return jsonify(subject)


@app.route('/run_loop', methods=["GET"])
def run_loop():
    interval = request.args.get('interval')
    job = scheduler.add_job(run_process, 'interval', minutes=int(interval), args=[interval])
    scheduler.start()
    return "running"


@app.route('/stop_loop', methods=["GET"])
def stop_loop():
    global state
    state = request.args.get('state')
    return f"set state {state}"


@app.route('/get_state', methods=["GET"])
def get_state():
    global state
    return request.url


@app.route('/get_file/<filename>', methods=["GET"])
def get_file(filename):
    try:
        # return request.url
        return send_from_directory(folder, filename=filename, as_attachment=True)
    except FileNotFoundError:
        abort(404)
    return request.url_root

if __name__ == '__main__':
    app.run()

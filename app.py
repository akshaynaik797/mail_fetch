import sqlite3
from time import sleep

import requests
from flask import Flask, request, jsonify

from functions import get_from_query, get_mail_id_list, download_pdf_and_html, process_row
from make_log import log_exceptions

app = Flask(__name__)

dbname, state = "database1.db", "running"


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
    if request.form.get('runno') != None:
        runno = request.form['runno']

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
        except:
            log_exceptions()
    return jsonify(row_no, ins, process)


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
    global state
    interval = request.args.get('interval')
    while 1:
        if state != 'stop':
            print(f"running every {interval} seconds.")
            try:
                #all steps here
                a = get_from_query()
                if isinstance(a, dict):
                    raise Exception
                print("got api response")
                b = get_mail_id_list('Max', a)
                print("got id list")
                download_pdf_and_html('Max', b)
                print("downloaded files and save data in db")
            except:
                log_exceptions()
                pass
            sleep(int(interval))
        else:
            return f"stopped becuase state={state}"


@app.route('/stop_loop', methods=["GET"])
def stop_loop():
    global state
    state = request.args.get('state')
    return f"set state {state}"


@app.route('/get_state', methods=["GET"])
def get_state():
    global state
    return f"state = {state}"


if __name__ == '__main__':
    app.run()

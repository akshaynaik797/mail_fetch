import sqlite3
from time import sleep

from flask import Flask, request, jsonify

app = Flask(__name__)

dbname, state = "database1.db", "running"


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
        except:
            pass
    return jsonify(row_no, ins, process)


@app.route('/process_subject', methods=["POST"])
def process_subject():
    subject, ins, process, result = "", "", "", []
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
                pass
        except:
            pass
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
                pass
            except:
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

import sqlite3
from flask import Flask, request, jsonify
app = Flask(__name__)


dbname = "database1.db"

@app.route('/get_details')
def get_details():
    result = ""
    q = "select * from run_table where completed = ''"
    with sqlite3.connect(dbname) as con:
        cur = con.cursor()
        result = cur.execute(q).fetchall()
    return jsonify(result)


@app.route('/post_details',methods=["POST"])
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


@app.route('/process_records',methods=["POST"])
def process_records():
    row_no, ins, process = "", "", ""
    if request.method == 'POST':
        try:
            row_no, ins, process = request.json['row_no'], request.json['ins'], request.json['process']
        except:
            pass
    return jsonify(row_no, ins, process)


if __name__  == '__main__':
    app.run()

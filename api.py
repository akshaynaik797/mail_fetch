import sqlite3
import sqlite3
from flask import Flask, request, jsonify
app = Flask(__name__)


dbname = "database1.db"

@app.route('/get_details')
def get_details():
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

if __name__  == '__main__':
    app.run()

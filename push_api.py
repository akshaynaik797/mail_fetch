import json
import sqlite3
from make_log import log_data

import requests


def api_trigger(time_diff):
    url = "https://exp.host/--/api/v2/push/send"
    # payload = "{\n  \"to\": \"ExponentPushToken[t0GrJvLoWrZPJAO2n1jBjc]\",\n  \"title\":\"Varun3\",\n  \"body\": \"Piyush\"\n}"
    headers = {
        'Content-Type': 'application/json'
    }
    with sqlite3.connect("database1.db") as con:
        cur = con.cursor()
        b = f"SELECT * from mob_app"
        cur.execute(b)
        r = cur.fetchall()
    if r is not None:
        for i in r:
            payload = "{\n  \"to\": \"" + i[1] + "\",\n  \"title\":\"Update Status\",\n  \"body\": \"Row took more than " + str(time_diff) + " seconds. Please infrom Varun\"\n}"
            response = requests.request("POST", url, headers=headers, data=payload)
            log_data(token=i[1], response=response.text)
            pass

    # response = requests.request("POST", url, headers=headers, data = payload)
    # print(response.text.encode('utf8'))

def api_update_trigger(ref_no, comment, status):
    if status == 'Acknowledgement':
        status = "In progress"
    body = ref_no+"-"+status+"-"+comment
    url = "https://exp.host/--/api/v2/push/send"
    # payload = "{\n  \"to\": \"ExponentPushToken[t0GrJvLoWrZPJAO2n1jBjc]\",\n  \"title\":\"Varun3\",\n  \"body\": \"Piyush\"\n}"
    headers = {
        'Content-Type': 'application/json'
    }
    with sqlite3.connect("database1.db") as con:
        cur = con.cursor()
        b = f"SELECT * from mob_app"
        cur.execute(b)
        r = cur.fetchall()
    if r is not None:
        for i in r:
            # payload = "{\n  \"to\": \"" + i[1] + "\",\n  \"title\":\"Update Status\",\n  \"body\": \"Row took more than " + str(time_diff) + " seconds. Please infrom Varun\"\n}"
            payload = {
                "to": i[1],
                "title": "Please Update Status",
                "body": body
            }
            response = requests.request("POST", url, headers=headers, data=json.dumps(payload))
            log_data(token=i[1], response=response.text)
            pass

if __name__ == "__main__":
    api_update_trigger('a', 'b', 'c')
    pass
import sqlite3
from make_log import log_exceptions

def check_if_sub_and_ltime_exist(subject, l_time):
    try:
        subject = subject.replace("'", '')
        with sqlite3.connect("database1.db") as con:
            xyz = 10
            cur = con.cursor()
            b = f"select * from updation_detail_log where emailsubject='{subject}' and date='{l_time}'"
            cur.execute(b)
            r = cur.fetchone()
            if r is not None:
                return True
            return False
    except:
        False

if __name__ == "__main__":
    l_time = '15/09/2020 13:37:28'
    subject = "Service Anniversary : 16th Sep'20"
    print(check_if_sub_and_ltime_exist(subject, l_time))
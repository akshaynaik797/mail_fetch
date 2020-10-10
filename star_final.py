import datetime
import re
import subprocess
import sys

import pdftotext

from make_log import log_exceptions
from bad_pdf import text_from_pdf

now = datetime.datetime.now()
'''
wbkName = 'log file.xlsx'
wbk= openpyxl.load_workbook(wbkName)
s1=wbk.worksheets[0]
s2=wbk.worksheets[1]
row_count_1 = s2.max_row
s2.cell(row_count_1+1, column=1).value=sys.argv[2]
s2.cell(row_count_1+1, column=2).value=sys.argv[3]
s2.cell(row_count_1+1, column=3).value=sys.argv[4]
s2.cell(row_count_1+1, column=5).value=now
s2.cell(row_count_1+1, column=7).value=sys.argv[5]
s2.cell(row_count_1+1, column=8).value=sys.argv[6]
'''

subprocess.run(["python", "updation.py", "1", "max1", "1", sys.argv[2]])
subprocess.run(["python", "updation.py", "1", "max", "2", sys.argv[3]])
subprocess.run(["python", "updation.py", "1", "max", "3", sys.argv[4]])
subprocess.run(["python", "updation.py", "1", "max", "5", str(now)])
subprocess.run(["python", "updation.py", "1", "max", "7", sys.argv[5]])
subprocess.run(["python", "updation.py", "1", "max", "8", sys.argv[6]])

flag = 0


with open(sys.argv[1], "rb") as f:
    pdf = pdftotext.PDF(f)

with open('star/output.txt', 'w') as f:
    f.write(" ".join(pdf))
num_lines = sum(1 for line in open('star/output.txt'))
if num_lines < 2:
    flag = 1
    text_from_pdf(sys.argv[1], 'star/output.txt')

with open('star/output.txt', 'r') as myfile:
    f = myfile.read()

try:
    hg = []

    subject = sys.argv[5]
    temp = re.compile(r".*(?=-)").search(subject)
    if temp is not None:
        preid = temp.group().strip()
    else:
        preid = ""
    hg.append(preid)
    status = 'Approved'

    badchars = (',', ':', '-')
    datadict = {}
    regexdict = {'amount': [r"(?<=final authorized amount is revised to Rs.)\d+",
                            r"(?<=Total Authorised Amount : Rs.).*",
                            r"(?<=Total Authorised Amount).*",
                            r"(?<=Initial Approval Amount :- Rs.).*\d+"],
                 'polno': [r"(?<=Policy Number).*\d+", r"(?<=Policy No.).*\d+"],
                 'memid': [r"(?<=ID/TPA/Insurer Id of the).*(?=-)"],
                 'pname': [r"(?<=Patient Name)[ A-Z:]+(?= )", r"(?<=Name of the Insured-Patient).*(?=,)", r"(?<=Patient\? s Member).*"],
                 }

    for i in regexdict:
        for j in regexdict[i]:
            data = re.compile(j).search(f)
            if data is not None:
                temp = data.group().strip()
                for k in badchars:
                    temp = temp.replace(k, "")
                datadict[i] = temp.strip()
                break
            datadict[i] = ""
    hg.append(datadict['amount'])
    hg.append(datadict['polno'])
    hg.append(datadict['memid'])
    hg.append(datadict['pname'])


    subprocess.run(["python", "updation.py", "1", "max", "9", 'Yes'])
    subprocess.run(["python", "updation.py", "1", "max", "10", 'NA'])

    try:
        subprocess.run(
            ["python", "test_api.py", hg[0], hg[1], hg[2], 'Cl', status, sys.argv[6], sys.argv[1], '', hg[3], hg[4]])
        '''wbk= openpyxl.load_workbook(wbkName)
		s2=wbk.worksheets[1]
		s2.cell(row_count_1+1, column=11).value='YES'
		'''
        subprocess.run(["python", "updation.py", "1", "max", "11", 'Yes'])
    except Exception as e:
        # s2.cell(row_count_1+1, column=11).value='NO'
        log_exceptions()
        subprocess.run(["python", "updation.py", "1", "max", "11", 'No'])
except Exception as e:
    # s2.cell(row_count_1+1, column=9).value='No'
    # s2.cell(row_count_1+1, column=11).value='NO'
    log_exceptions()
    subprocess.run(["python", "updation.py", "1", "max", "9", 'Yes'])
    subprocess.run(["python", "updation.py", "1", "max", "11", 'No'])
now = datetime.datetime.now()
# s2.cell(row_count_1+1, column=6).value=now
# wbk.save(wbkName)
subprocess.run(["python", "updation.py", "1", "max", "6", str(now)])

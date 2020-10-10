import datetime
import subprocess
import sys

import pdftotext

from patient_name_fun import pname_fun

now = datetime.datetime.now()

subprocess.run(["python", "updation.py", "1", "max1", "1", sys.argv[2]])
subprocess.run(["python", "updation.py", "1", "max", "2", sys.argv[3]])
subprocess.run(["python", "updation.py", "1", "max", "3", sys.argv[4]])
subprocess.run(["python", "updation.py", "1", "max", "5", str(now)])
subprocess.run(["python", "updation.py", "1", "max", "7", sys.argv[5]])
subprocess.run(["python", "updation.py", "1", "max", "8", sys.argv[6]])

with open(sys.argv[1], "rb") as f:
    pdf = pdftotext.PDF(f)

with open('star/output.txt', 'w') as f:
    f.write(" ".join(pdf))
with open('star/output.txt', 'r') as myfile:
    f = myfile.read()
try:
    hg = []
    w = sys.argv[5]  # f.find('Claim intimation No.')+20
    g = w.find('-')  # f[w:]
    # u=g.find('-')+w
    hg.append(w[:g - 1])

    status = 'Information Awaiting'
    hg.append('')
    hg.append('')
    hg.append('')

    hg = [sub.replace(':', '') for sub in hg]
    hg = [sub.replace('  ', '') for sub in hg]
    hg = [sub.replace('Rs.', '') for sub in hg]
    print(hg, status)
    subprocess.run(["python", "updation.py", "1", "max", "9", 'Yes'])
    subprocess.run(["python", "updation.py", "1", "max", "10", 'NA'])

    try:
        regex_list = [r"(?<=Name of Insured-).*(?=Age)", r"(?<=NAME OF INSURED-PATIENT).*(?=AGE)"]
        pname = pname_fun(f, regex_list)
        if (f.find('QUERY ON AUTHORIZATION') != -1):
            subprocess.run(
                ["python", "test_api.py", hg[0], hg[1], hg[2], 'Pa', status, sys.argv[6], sys.argv[1], '', hg[3],
                 pname])
            subprocess.run(["python", "updation.py", "1", "max", "11", 'Yes'])
        else:
            subprocess.run(
                ["python", "test_api.py", hg[0], hg[1], hg[2], 'Cl', status, sys.argv[6], sys.argv[1], '', hg[3],
                 pname])
            subprocess.run(["python", "updation.py", "1", "max", "11", 'Yes'])

    except Exception as e:
        subprocess.run(["python", "updation.py", "1", "max", "11", 'No'])
except Exception as e:
    subprocess.run(["python", "updation.py", "1", "max", "9", 'Yes'])
    subprocess.run(["python", "updation.py", "1", "max", "11", 'No'])
now = datetime.datetime.now()
# s2.cell(row_count_1+1, column=6).value=now
# wbk.save(wbkName)
subprocess.run(["python", "updation.py", "1", "max", "6", str(now)])

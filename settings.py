import pdfkit

platform = 'linux'

if platform != 'linux':
    path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
else:
    config = pdfkit.configuration(wkhtmltopdf='/usr/bin/wkhtmltopdf')

folder, dbname = "files/", "database1.db"

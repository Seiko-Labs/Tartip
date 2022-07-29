import subprocess
import smtplib
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename

cmd_tablet = 'python get_Kpi_tablet.py'
cmd_TS = 'python get_Kpi_auto.py'
cmd_drivers = 'python get_Kpi_driver.py'
cmd_exel = 'python get_prize.py'

p = subprocess.Popen(cmd_tablet, stdout=subprocess.PIPE, shell=True)
out_tablet, err_tablet = p.communicate()
print("Получение фотофиксаций")
print(out_tablet)
print(err_tablet)
p = subprocess.Popen(cmd_TS, stdout=subprocess.PIPE, shell=True)
out_TS, err_TS = p.communicate()
print("Получение Kpi машин")
print(out_TS)
print(err_TS)
p = subprocess.Popen(cmd_drivers, stdout=subprocess.PIPE, shell=True)
out_drivers, err_drivers = p.communicate()
print("Получение Kpi водителей и грузчиков")
print(out_drivers)
print(err_drivers)
p = subprocess.Popen(cmd_exel, stdout=subprocess.PIPE, shell=True)
out_exel, err_exel = p.communicate()
print("Составление ексель файла")
print(out_drivers)
print(err_drivers)

# smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
# smtpObj.ehlo()
# smtpObj.starttls()
# smtpObj.ehlo()
# smtpObj.login('oxana.melnik666@gmail.com','wolf6256256')
#
# smtpObj.sendmail("oxana.melnik666@gmail.com","oxana.melnik666@gmail.com","go to bed!")

def send_mail(to_email, subject, files=[], message='', server='mail.tartyp.kz', from_email='robot@tartyp.kz'):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message))
    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)
    server = smtplib.SMTP(server)
    server.set_debuglevel(1)
    server.login(from_email, '#%Qf4yT{Jx')
    server.send_message(msg)
    server.quit()

now = datetime.now()
month = now.strftime("%m-%Y")
subject = 'Ежедневный отчет по рассчету Kpi за ' + str(month) + ' месяц.'
exel_name = 'Salary/' + 'Kpi_prize_' + str(month) + '.xlsx'
send_mail('oxana.melnik666@gmail.com', 'subject', [exel_name], 'Ежедневный отчет по рассчету Kpi водителей и грузчиков')
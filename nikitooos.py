import csv
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, formatdate
from docxtpl import DocxTemplate
from datetime import datetime

def load_employee_data(file_path='data.csv'):
    employee_data = []
    with open(file_path, encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # Пропускаем заголовок
        for row in reader:
            employee_data.append(row)
    return employee_data

def load_issued_references(file_path='issued_references_log.csv'):
    issued_references = []
    with open(file_path, encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # Пропускаем заголовок
        for row in reader:
            issued_references.append(row)
    return issued_references

def log_issued_reference(reference_number, name, issue_date, log_file='issued_references_log.csv'):
    if not os.path.exists(log_file):
        with open(log_file, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Reference Number', 'Name', 'Issue Date'])
    with open(log_file, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([reference_number, name, issue_date])

def send_email(to_address, subject, body, attachment):
    from_address = "your@mail.ru"
    password = "your_password"
    msg = MIMEMultipart()
    msg['From'] = formataddr(('АО «Теплоконтроль»', from_address))
    msg['To'] = to_address
    msg['Subject'] = subject
    msg['Date'] = formatdate(localtime=True)
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    with open(attachment, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    file_name = os.path.basename(attachment)
    part.add_header('Content-Disposition', 'attachment', filename=file_name)
    msg.attach(part)
    server = smtplib.SMTP('smtp.mail.ru', 587)
    server.starttls()
    server.login(from_address, password)
    text = msg.as_string()
    server.sendmail(from_address, to_address, text)
    server.quit()

def get_next_reference_number(file_path='reference_number.txt'):
    with open(file_path, 'r+') as f:
        current_number = int(f.read().strip())
        new_number = current_number + 1
        f.seek(0)
        f.write(str(new_number))
        f.truncate()
    return f"00-{current_number:04d}"

def process_reference(employee_id):
    with open("data.csv", encoding='utf-8') as r_file:
        file_reader = csv.reader(r_file, delimiter=",")
        next(file_reader)  # Пропускаем заголовок
        found = False
        for row in file_reader:
            if row[0] == employee_id:
                employee_name = row[2]
                found = True
                break
        if not found:
            print("Сотрудник с таким идентификатором не найден.")
        else:
            doc = DocxTemplate("template.docx")
            issue_date = datetime.now().strftime("%d.%m.%Y")
            reference_number = get_next_reference_number()
            context = {
                "group_number": row[1],
                "fio_stud": employee_name,
                "date_protocol": row[3],
                "number_protocol": row[4],
                "issue_date": issue_date,
                "reference_number": reference_number,
            }
            doc.render(context)
            file_name = f"справка-{employee_name}.docx"
            print("Имя файла:", file_name)
            doc.save(file_name)
            send_email("your@gmail.com", "[Кадры] Справки для выдачи", "Необходимо заверить печатью АО Теплоконтроль с момента поступления файл указанный во вложении. Выдача справок происходит в течение 2-х рабочих дней с даты подписания в соответствии с распоряжением директора 32ЛС. С уважением, директор АО Теплоконтроль Пеньковцев Н.Р.." , file_name)
            print("Справка успешно отправлена!")
            log_issued_reference(reference_number, employee_name, issue_date)

def get_employee_info(employee_id):
    with open("data.csv", encoding='utf-8') as r_file:
        file_reader = csv.reader(r_file, delimiter=",")
        next(file_reader)  # Пропускаем заголовок
        for row in file_reader:
            if row[0] == employee_id:
                return row
    return None

def get_employee_ids():
    employee_ids = []
    with open("data.csv", encoding='utf-8') as file:
        file_reader = csv.reader(file, delimiter=",")
        next(file_reader)  # Пропускаем заголовок
        for row in file_reader:
            employee_ids.append(row[0])
    return employee_ids

if __name__ == "__main__":
    process_reference("идентификатор_сотрудника")

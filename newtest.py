from email import encoders
from email.mime.base import MIMEBase

import mysql.connector
import openpyxl
from openpyxl.workbook import Workbook
from datetime import datetime

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

db_config = {
    'host': '185.146.1.107',
    'user': 'admin',
    'password': 'thPZJhRIyMoY23P5ayj7K1k=',
    'database': 'emg-cm-test',
}

try:
    # Устанавливаем соединение
    connection = mysql.connector.connect(**db_config)

    if connection.is_connected():
        print("Успешное подключение к базе данных")

        # код работы с базой данных

    sql_query = '''
        # SELECT 
        #   obj_num, 
        #   COUNT(*) as count_per_day, 
        #   AVG(`0030`) as avg_0030_per_day,
        #   AVG(`0031`) as avg_0031_per_day,
        #   AVG(`0032`) as avg_0032_per_day,
        #   AVG(`0033`) as avg_0033_per_day,
        #   AVG(`0034`) as avg_0034_per_day,
        #   AVG(`0035`) as avg_0035_per_day,
        #   AVG(`0036`) as avg_0036_per_day,
        #   AVG(`0037`) as avg_0037_per_day,
        #   AVG(`003B`) as avg_003B_per_day
        # FROM 
        #   regular
        # WHERE 
        #   DATE(dt_0) = CURDATE()
        # GROUP BY 
        #   obj_num;

        SELECT 
          obj_num,
          DATE(dt_0) as date,
          COUNT(*) as count_per_day,
          AVG(`AB10`) as average_AB10_per_day,
          AVG(`AB11`) as average_AB11_per_day,
          AVG(`AB12`) as average_AB12_per_day,
          AVG(`AB13`) as average_AB13_per_day,
          AVG(`AB14`) as average_AB14_per_day
        FROM 
          regular
        GROUP BY 
          obj_num, date;

    '''

    current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f'report_{current_datetime}.xlsx'

    workbook = Workbook()
    sheet = workbook.active

    cursor = connection.cursor(dictionary=True)
    cursor.execute(sql_query)
    headers = [column[0] for column in cursor.description]
    sheet.append(headers)

    for row in cursor:
        sheet.append(list(row.values()))

    workbook.save(file_name)
    print(f'Название xlsx файла : {file_name}')

except mysql.connector.Error as err:
    print(f"Ошибка при подключении к базе данных: {err}")


finally:
    # закрываем курсор и соединение после использования
    if 'cursor' in locals() and cursor is not None:
        cursor.close()
        print("Курсор закрыт")

    if 'connection' in locals() and connection.is_connected():
        connection.close()
        print("Соединение закрыто")

####  -------- Ниже код отправки файла по почте

# Настройки почтового сервера
email_host = 'smtp.mail.ru'
email_port = 587
email_username = 'bekzat.ablaev99@mail.ru'
email_password = 'imLrrWcQT1NhczpgWQMc'

# Адрес отправителя и получателя
from_email = 'bekzat.ablaev99@mail.ru'
to_email = 'bekzatablaev@gmail.com'

# Создание сообщения
subject = ''
body = ''

message = MIMEMultipart()
message['From'] = from_email
message['To'] = to_email
message['Subject'] = subject
message.attach(MIMEText(body, 'plain'))

attachment = open(file_name, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename= {file_name}')
message.attach(part)

# Подключение к почтовому серверу
try:
    server = smtplib.SMTP(email_host, email_port)
    server.starttls()
    server.login(email_username, email_password)
    server.sendmail(from_email, to_email, message.as_string())
    server.quit()
    print('Письмо успешно отправлено!')
except Exception as e:
    print(f'Ошибка отправки письма: {e}')

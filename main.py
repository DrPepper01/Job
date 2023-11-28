from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import mysql.connector
from openpyxl import Workbook
from datetime import datetime

db_config = {
    'host': '185.146.1.107',
    'user': 'admin',
    'password': 'thPZJhRIyMoY23P5ayj7K1k=',
    'database': 'emg-cm-test',
}

try:
    connection = mysql.connector.connect(**db_config)

    if connection.is_connected():
        print("Успешное подключение к базе данных")

    sql_query = '''
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

#    workbook.save(file_name)
    print(f'Название xlsx файла : {file_name}')

    # Код отправки почты через Gmail
    email_host = 'smtp.gmail.com'
    email_port = 587
    email_username = 'bekzatablaev@gmail.com'
    email_password = 'osgk uuod zhin bvku'

    from_email = 'bekzatablaev@gmail.com'
    to_email = 'bekzat.ablaev99@mail.ru'

    subject = 'Отчет по данным'
    body = 'Добрый день! Пожалуйста, найдите вложенный отчет.'

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

    server = smtplib.SMTP(email_host, email_port)
    server.starttls()
    server.login(email_username, email_password)
    server.sendmail(from_email, to_email, message.as_string())
    server.quit()

    print('Письмо успешно отправлено!')

except mysql.connector.Error as err:
    print(f"Ошибка при подключении к базе данных: {err}")

except Exception as e:
    print(f'Ошибка: {e}')

finally:
    if 'cursor' in locals() and cursor is not None:
        cursor.close()
        print("Курсор закрыт")

    if 'connection' in locals() and connection.is_connected():
        connection.close()
        print("Соединение закрыто")

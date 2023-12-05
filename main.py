import configparser
import datetime
import os

import pandas as pd

config = configparser.ConfigParser()
config.read('config\config.ini')

save_location = config.get("DEFAULT", "save_location")
input_location = config.get("DEFAULT", "input_location")
filename = config.get("DEFAULT", "filename")
send_mail = config.get("SMTP", "send_mail")

send_mail = bool(send_mail)

os.makedirs(save_location, exist_ok=True)

datum = datetime.datetime.now().strftime("%m-%Y")
try:
    csv = pd.read_csv(f'{input_location}', delimiter=';')

    result = pd.DataFrame(
        columns=['Firma', 'Essential', 'Business', 'Enterprise', 'Enterprise-Dial-in-pack', 'Attendant',
                 'Alert', 'Connect', 'Dial in per use', 'Room', 'Business-Custom', 'Enterprise-Custom',
                 'Attendant-Custom', 'Voice-Business-Custom', 'Voice-Enterprise-Custom',
                 'Voice-Attendant-Custom', 'Alert-Custom', 'Room-Custom'])

    for i, row in csv.iterrows():
        customer_name = row['customerName']
        volume = row['volume']
        lizenz_typ = 'NA'

        if 'Essential' in row['rainbowServiceId']:
            lizenz_typ = 'Essential'

        elif 'Business' in row['rainbowServiceId']:
            lizenz_typ = 'Business'

        elif 'Voice-Enterprise-Custom' in row['rainbowServiceId']:
            lizenz_typ = 'Voice-Enterprise-Custom'

        elif 'Enterprise-Custom' in row['rainbowServiceId']:
            lizenz_typ = 'Enterprise-Custom'

        elif 'Enterprise-Dial-in-pack' in row['rainbowServiceId']:
            lizenz_typ = 'Enterprise-Dial-in-pack'

        elif 'Enterprise' in row['rainbowServiceId']:
            lizenz_typ = 'Enterprise'

        elif 'Alert-Custom' in row['rainbowServiceId']:
            lizenz_typ = 'Alert-Custom'

        elif 'Room-Custom' in row['rainbowServiceId']:
            lizenz_typ = 'Room-Custom'

        elif 'Alert' in row['rainbowServiceId']:
            lizenz_typ = 'Alert'

        elif 'Connect' in row['rainbowServiceId']:
            lizenz_typ = 'Connect'

        elif 'Dial-in-per-use' in row['rainbowServiceId']:
            lizenz_typ = 'Dial-in-per-use'

        elif 'Room' in row['rainbowServiceId']:
            lizenz_typ = 'Room'

        elif 'Voice-Attendant-Custom' in row['rainbowServiceId']:
            lizenz_typ = 'Voice-Attendant-Custom'

        elif 'Attendant-Custom' in row['rainbowServiceId']:
            lizenz_typ = 'Attendant-Custom'

        elif 'Attendant' in row['rainbowServiceId']:
            lizenz_typ = 'Attendant'

        if customer_name in result["Firma"]:
            result.loc[result['Firma'] == customer_name, f'{lizenz_typ}'] += volume
        else:
            new_row = {'Firma': customer_name}
            new_row.update({f'{lt}': volume if lt == lizenz_typ else 0 for lt in result.columns[1:]})
            result = result._append(new_row, ignore_index=True)
except Exception as Exc1:
    print(f'Fehler: {Exc1}')
try:
    file = f'{save_location}\{filename}_{datum}'

    result = result.groupby('Firma').sum()
    result.to_csv(f'{file}.csv', sep=';', index=True, encoding='utf-8')
except Exception as Exc2:
    print(f'Fehler: {Exc2}')
try:
    if send_mail is True:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders

        smtp_server = config.get("SMTP", "smtp_server")
        smtp_port = config.get("SMTP", "smtp_port")
        sender_email = config.get("SMTP", "sender_mail")
        sender_password = config.get("SMTP", "sender_password")
        receiver_email = config.get("SMTP", "receiver_mail")
        subject = config.get("SMTP", "subject")
        body = config.get("SMTP", "body")

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        file_path = f'{file}.csv'
        attachment = open(file_path, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % f'Lizenzen_{datum}.csv')

        message.attach(part)

        smtp_port = int(smtp_port)

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, message.as_string())
            server.quit()
        print("E-Mail wurde versendet.")
except Exception as Exc3:
    print(f'Fehler bei SMTP: {Exc3}')

import imaplib
import email
from email.header import decode_header
import yaml
from datetime import datetime
import pytz
from dateutil import tz

with open("CREDENTIALS.yml") as f:
    content = f.read()

my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# información de la cuenta
username = my_credentials["user"]
password = my_credentials["password"]

# crear una instancia de IMAP4 class con SSL
mail = imaplib.IMAP4_SSL("imap.gmail.com")
# autenticar
mail.login(username, password)

# seleccionar la bandeja de entrada
mail.select("inbox")
# buscar todos los correos
result, data = mail.uid('search', None, "ALL")
# obtener la lista de ID de correo
mail_ids = data[0].split()
# recorrer todos los ID de correo en orden inverso (del más reciente al más antiguo)
for i in range(-1, -len(mail_ids), -1):
    # obtener el correo (raw)
    result, data = mail.uid('fetch', mail_ids[i], '(BODY.PEEK[])')
    raw_email = data[0][1].decode("utf-8")
    # convertir a mensaje de correo
    email_message = email.message_from_string(raw_email)

    # obtener la fecha del correo
    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    if date_tuple:
        local_date = datetime.fromtimestamp(
            email.utils.mktime_tz(date_tuple))
        local_message_date = "%s" % (
            str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))

        # obtener la fecha y hora actuales
        now = datetime.now(pytz.timezone('GMT'))

        # convertir la fecha del correo a datetime
        email_date_uno = datetime.strptime(
            local_message_date, "%a, %d %b %Y %H:%M:%S")

        # convertir la fecha del correo a datetime y hacerla consciente de la zona horaria
        email_date = email_date_uno.replace(tzinfo=tz.tzutc())

        # calcular la diferencia en minutos entre la fecha del correo y la fecha actual
        diff = (now - email_date).total_seconds() / 60

        # si la diferencia es menor o igual a 5 minutos, imprimir los detalles del correo
        if diff <= 5:
            print("Fecha:", local_message_date)
            print("De:", decode_header(email_message['From'])[0][0])
            print("Asunto:", decode_header(email_message['Subject'])[0][0])
            print("Mensaje:")
            if email_message.is_multipart():
                for payload in email_message.get_payload():
                    # si la carga útil es de tipo texto o html, entonces imprime el mensaje
                    if payload.get_content_type() == "text/plain" or payload.get_content_type() == "text/html":
                        print(payload.get_payload())
            else:
                print(email_message.get_payload())
            break

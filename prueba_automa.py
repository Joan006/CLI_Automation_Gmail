import imaplib
import email
import yaml
import re
import pandas as pd
from datetime import datetime
from twilio.rest import Client
import time
import logging

# Configurar el sistema de registro
logging.basicConfig(filename='correo_script.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s: %(message)s')


def cargar_credenciales(file_path):
    # funcion para leer las Credenciales en el archivo yaml
    with open(file_path) as f:
        content = f.read()
    return yaml.load(content, Loader=yaml.FullLoader)


def conectar_gmail(user, password, imap_url="imap.gmail.com"):
    # funcion para conectar_gmail con user , password , imap_url
    my_mail = imaplib.IMAP4_SSL(imap_url)
    my_mail.login(user, password)
    # iniciamos serion en gmail con el user y password
    my_mail.select("inbox")
    return my_mail
    # retornamos la bandeja que ocuparemos


def buscar_ultimo_correo(my_mail, key, value):
    # funcion que retorna el ultimo correo , con el key y value
    _, data = my_mail.search(None, key, value)
    # Buscar en la bandeja los correos relacionados con 'value'
    mail_id_list = data[0].split()
    # Seleccionar y separar el numero del correo electr√≥nico
    msgs = []
    if mail_id_list:
        last_mail_id = mail_id_list[-1]
        typ, last_mail_data = my_mail.fetch(last_mail_id, '(RFC822)')
        msgs.append(last_mail_data)
    else:
        logging.info(
            "No se encontraron correos electrónicos que cumplan con los criterios.")
    return msgs
    # retornamos el ultimo correo electronico deseado


def extraer_datos_correo(msgs):
    data = {'Subject': [], 'From': [], 'Body': []}
    # Crear un DataFrame para almacenar los datos
    for msg in msgs[::-1]:
        # Iterar sobre los mensajes recuperados y extraer informacion
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg = email.message_from_bytes(response_part[1])

                # Agregar informacion al DataFrame
                data['Subject'].append(" 📎 : " + my_msg['subject'])
                data['From'].append(" 📬 : " + my_msg['from'])

                body = ''
                for part in my_msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload()
                
                # Codificar y formatear el cuerpo del mensaje modificacion google
                body = body.encode('utf-8').decode('unicode_escape')
                body = re.sub(r'\s+', ' ', body)
                data['Body'].append(body)
    return pd.DataFrame(data)
    # Crear un DataFrame con los datos


def cargar_datos_excel(excel_path, sheet_name):
    try:
        df_existing = pd.read_excel(excel_path, sheet_name=sheet_name)
        # Cargar la hoja de Excel existente en el DataFrame
    except FileNotFoundError:
        # Si el archivo no existe, crea un DataFrame vacío
        df_existing = pd.DataFrame(columns=['Subject', 'From', 'Body'])
    return df_existing


def enviar_mensaje_whatsapp(ultima_fila, nueva_columna_nombre, account_sid, auth_token, numero_celular):
    asunto = ultima_fila['Subject']
    remitente = ultima_fila['From']
    cuerpo = ultima_fila['Body']

    mensaje_whatsapp = (
        f"📩 *Nuevo Correo Electrónico* 📩 \n"
        f"Asunto: {asunto}\n"
        f"Remitente: {remitente}\n"
        f"Cuerpo del Correo: {cuerpo}"
    )

    columnas_seleccionadas = ultima_fila[
        [nueva_columna_nombre, 'Subject', 'From', 'Body']]
    mensaje_whatsapp = columnas_seleccionadas.to_string(index=False)
    mensaje_whatsapp = mensaje_whatsapp.encode('utf-8').decode('utf-8')
    partes_mensaje = [mensaje_whatsapp[i:i+1500]
                      for i in range(0, len(mensaje_whatsapp), 1500)]

    client = Client(account_sid, auth_token)
    message = client.messages.create(
        from_='whatsapp:+14155238886',
        body=partes_mensaje,
        to=numero_celular
    )

    logging.info(f"Mensaje de WhatsApp enviado: {message.sid}")


def procesar_correos(user, password, excel_path, sheet_name, account_sid, auth_token, numero_celular):
    try:
        my_mail = conectar_gmail(user, password)

        key = "from"
        value = "joan_martinez.olivares@hotmail.com"
        msgs = buscar_ultimo_correo(my_mail, key, value)

        if msgs:
            df = extraer_datos_correo(msgs)
            df_existing = cargar_datos_excel(excel_path, sheet_name)
            df_combined = pd.concat([df_existing, df], ignore_index=True)

            nueva_columna_nombre = 'NuevaColumna'
            nuevo_valor = f"📩 NUEVO 📩"
            df_combined[nueva_columna_nombre] = nuevo_valor

            ultima_fila = df_combined.iloc[-1]

            if not df_existing.equals(df_combined):
                enviar_mensaje_whatsapp(
                    ultima_fila, nueva_columna_nombre, account_sid, auth_token, numero_celular)

                df_combined.to_excel(
                    excel_path, sheet_name=sheet_name, index=False)

    except Exception as e:
        logging.error(f"Error: {e}")

    finally:
        try:
            my_mail.close()
        except:
            pass
        my_mail.logout()


if __name__ == "__main__":
    my_credentials = cargar_credenciales("CREDENTIALS.yml")
    user, password = my_credentials["user"], my_credentials["password"]
    excel_path = 'emails.xlsx'
    sheet_name = 'Hoja1'
    account_sid = 'AC3430b29f9d9b9f403701741fafdf1150'
    auth_token = my_credentials["auth_token"]
    numero_celular = my_credentials['numero_celular']

    while True:
        try:
            procesar_correos(
                user, password, excel_path, sheet_name, account_sid, auth_token, numero_celular)

        except Exception as e:
            logging.error(f"Error: {e}")

        logging.info(f"Ejecutado a las {datetime.now()}")
        time.sleep(60)

import imaplib
import email
import yaml
import re
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from twilio.rest import Client
import time
import logging

# Configurar el sistema de registro
logging.basicConfig(filename='correo_script.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s: %(message)s')


# Abrir el archivo donde se encuentran las credenciales
with open("CREDENTIALS.yml") as f:
    content = f.read()

# Trabajar con los datos del archivo YAML
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# Declarar las variables de usuario y contraseña
user, password = my_credentials["user"], my_credentials["password"]

# URL IMAP
imap_url = "imap.gmail.com"

while True:
    try:
        # Conexión a Gmail mediante SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Iniciar sesión en Gmail
        my_mail.login(user, password)

        # Seleccionar la bandeja de entrada
        my_mail.select("inbox")

        key = "from"
        value = "joan_martinez.olivares@hotmail.com"

        # Buscar en la bandeja los correos relacionados con 'value'
        _, data = my_mail.search(None, key, value)

        # Seleccionar y separar el número del correo electrónico
        mail_id_list = data[0].split()

        msgs = []  # Lista para almacenar todos los mensajes

        # Extraer solo el último correo completo deseado
        if mail_id_list:
            last_mail_id = mail_id_list[-1]
            typ, last_mail_data = my_mail.fetch(last_mail_id, '(RFC822)')
            msgs.append(last_mail_data)
        else:
            logging.info(
                "No se encontraron correos electrónicos que cumplan con los criterios.")

        # Crear un DataFrame para almacenar los datos
        data = {'Subject': [], 'From': [], 'Body': []}

        # Iterar sobre los mensajes recuperados y extraer información
        for msg in msgs[::-1]:
            for response_part in msg:
                if type(response_part) is tuple:
                    my_msg = email.message_from_bytes((response_part[1]))

                    # Agregar información al DataFrame
                    data['Subject'].append(" 🔎 : " + my_msg['subject'])
                    data['From'].append(" 📭 : " + my_msg['from'])

                    body = ''
                    for part in my_msg.walk():
                        if part.get_content_type() == 'text/plain':
                            body = part.get_payload()

                    # Codificar y formatear el cuerpo del mensaje
                    body = body.encode('utf-8').decode('unicode_escape')
                    body = re.sub(r'\s+', ' ', body)
                    data['Body'].append(body)

        # Crear un DataFrame con los datos
        df = pd.DataFrame(data)

        # Cargar el archivo Excel existente
        excel_path = 'emails.xlsx'
        book = load_workbook(excel_path)

        # Obtener el nombre de la hoja
        sheet_name = 'Hoja1'  # Cambia esto al nombre de tu hoja existente si es diferente

        # Cargar la hoja de Excel existente en el DataFrame
        df_existing = pd.read_excel(excel_path, sheet_name=sheet_name)

        # Concatenar el nuevo DataFrame con el existente
        df_combined = pd.concat([df_existing, df], ignore_index=True)

        # Agregar nueva columna
        nueva_columna_nombre = 'NuevaColumna'
        nuevo_valor = f"📩 NUEVO 📩"
        df_combined[nueva_columna_nombre] = nuevo_valor

        # Obtener la última fila del DataFrame combinado
        ultima_fila = df_combined.iloc[-1]

        # Personalizar el mensaje de WhatsApp
        asunto = ultima_fila['Subject']
        remite = ultima_fila['From']
        cuerpo = ultima_fila['Body']

        # Construir el mensaje de WhatsApp con emojis y formato estético
        mensaje_whatsapp = (
            f"📩 *Nuevo Correo Electrónico* 📩 \n"
            f"Asunto: {asunto}\n"
            f"Remitente: {remite}\n"
            f"Cuerpo del Correo: {cuerpo}"
        )

        # Enviar mensaje por WhatsApp para la última fila
        columnas_seleccionadas = ultima_fila[[
            nueva_columna_nombre, 'Subject', 'From', 'Body',]]
        mensaje_whatsapp = columnas_seleccionadas.to_string(index=False)
        hora_actual = datetime.now().time()
        mensaje_whatsapp = mensaje_whatsapp.encode('utf-8').decode('utf-8')
        partes_mensaje = [mensaje_whatsapp[i:i+1500]
                          for i in range(0, len(mensaje_whatsapp), 1500)]

        account_sid = 'AC3430b29f9d9b9f403701741fafdf1150'
        auth_token = my_credentials["auth_token"]
        client = Client(account_sid, auth_token)

        message = client.messages.create(
            from_='whatsapp:+14155238886',
            body=partes_mensaje,
            to=my_credentials['numero_celular']
        )

        logging.info(f"Mensaje de WhatsApp enviado: {message.sid}")
        df_combined.to_excel(excel_path, sheet_name=sheet_name, index=False)

    except Exception as e:
        logging.error(f"Error: {e}")

    finally:
        try:
            # Cerrar la conexión
            my_mail.close()
        except:
            pass
        my_mail.logout()

    # Registrar la hora de la ejecución
    logging.info(f"Ejecutado a las {datetime.now()}")

    # Procesar nuevos correos cada 1 minutos (ajusta según tus necesidades)
    time.sleep(60)  # Esperar 1 minutos antes de volver a revisar

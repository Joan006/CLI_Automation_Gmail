import imaplib
import email
import yaml
import re
import pandas as pd
from openpyxl import load_workbook

# Abrir el archivo donde se encuentran las credenciales
with open("CREDENTIALS.yml") as f:
    content = f.read()

# Trabajar con los datos del archivo YAML
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# Declarar las variables de usuario y contraseña
user, password = my_credentials["user"], my_credentials["password"]

# URL IMAP
imap_url = "imap.gmail.com"

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
    print("No se encontraron correos electrónicos que cumplan con los criterios.")

# Crear un DataFrame para almacenar los datos
data = {'Subject': [], 'From': [], 'Body': []}

# Iterar sobre los mensajes recuperados y extraer información
for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg = email.message_from_bytes((response_part[1]))
            
            # Agregar información al DataFrame
            data['Subject'].append(my_msg['subject'])
            data['From'].append(my_msg['from'])
            

            body = ''
            for part in my_msg.walk():
                if part.get_content_type() == 'text/plain':
                    body = part.get_payload()

            # Codificar y formatear el cuerpo del mensaje modificacion google
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

# Escribir el DataFrame combinado en el archivo Excel existente
df_combined.to_excel(excel_path, sheet_name=sheet_name, index=False)



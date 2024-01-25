import imaplib
import email

import yaml
from yaml import parser

# Abrimos el archivo donde se encuentran las credenciales
with open("CREDENTIALS.yml") as f:
    content = f.read()

# es para trabajar con los datos de el archivo yml
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# declaramos las variables con el user  y password
user, password = my_credentials["user"], my_credentials["password"]

# URL IMAP 
imap_url = "imap.gmail.com"

# Conexion a gmail mediante ssl
my_mail = imaplib.IMAP4_SSL(imap_url)

# iniciamos secion en gmail
my_mail.login(user, password)

# seleccionamos la bandeja 
my_mail.select("Inbox")

# 
key = "from"
value = "joan_martinez.olivares@hotmail.com"

_, data = my_mail.search(None, key, value)  # buscamos en la bandeja los relacionados a value

# seleccionamos y separamos el numero del correo electronico
mail_id_list = data[0].split()

msgs = []  # creamos una lista donde estara todos los mensajes


# extraemos los correos completos deseados
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')
    msgs.append(data)

# extraeremos solamente el cuepro del correo electronico que nos interesa

for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg = email.message_from_bytes((response_part[1]))
            print("-----------------------------")
            print("subj:", my_msg['subject'])
            print("from:", my_msg['from'])
            print("body:")
            for part in my_msg.walk():
                print(part.get_content_type())
                if part.get_content_type() == 'text/plain':
                    print(part.get_payload())


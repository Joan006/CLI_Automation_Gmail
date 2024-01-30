import imaplib
import email
import yaml

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

key = "from"
value = "joan_martinez.olivares@hotmail.com"

# Buscamos en la bandeja los relacionados a value
_, data = my_mail.search(None, key, value)

# Seleccionamos y separamos el n√∫mero del correo electr√≥nico
mail_id_list = data[0].split()

msgs = []  # Creamos una lista donde estar√° solo el √∫ltimo mensaje

# Extraemos √∫nicamente el √∫ltimo correo completo deseado
if mail_id_list:
    last_mail_id = mail_id_list[-1]
    typ, last_mail_data = my_mail.fetch(last_mail_id, '(RFC822)')
    msgs.append(last_mail_data)
else:
    print("No se encontraron correos electr√≥nicos que cumplan con los criterios.")

# Ahora msgs contendr√° solo el √∫ltimo correo que cumple con los criterios

# extraeremos solamente el cuepro del correo electronico que nos interesa

for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg = email.message_from_bytes((response_part[1]))
            print("-----------------------------")
            print("subject:", my_msg['subject'])
            print("from:", my_msg['from'])
            print("body:")
            for part in my_msg.walk():
                print(part.get_content_type())
                if part.get_content_type() == 'text/plain':
                    print(part.get_payload())


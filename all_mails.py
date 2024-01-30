# Codigo para extraer todos los correos electronicso deseadods del usuario

# Criterios del correo electronico que solcitamos 
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



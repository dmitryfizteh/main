import spur
from shutil import copyfileobj

username = input("Input username to generate the keys: ")
mail_p = username + "@mail.ru"
mail_i = input("E-mail (по умолчанию " + mail_p + "): ")
if mail_i != "":
    mail = mail_i
else:
    mail = mail_p

print(mail)

# Подключаемся по ssh

shell = spur.SshShell(hostname="*", username="*", password="*")
with shell:
    result = shell.run(["echo", "-n", "hello"])
print(result.output) # prints hello

with shell.open("/etc/openvpn/keys/client.key", "rb") as remote_file:
    with open("./client.key", "wb") as local_file:
        copyfileobj(remote_file, local_file)

exit()


# Вносим изменения в config
# добавляем файлы в архив
# отправляем ключи на почту
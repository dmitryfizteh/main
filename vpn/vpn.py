import spur
from shutil import copyfileobj
import configparser

conf = configparser.RawConfigParser()
conf.read("vpn.ini")

username = input("Input username to generate the keys: ")
mail_p = username + "@" + conf.get("main", "MAIL_SUFF")
mail_i = input("E-mail (по умолчанию " + mail_p + "): ")
if mail_i != "":
    mail = mail_i
else:
    mail = mail_p

print(mail)

# Создаем из шаблона ovpn-файл
with open('client.ovpn') as file_in:
    text = file_in.read()
text = text.replace("username", username)
filename = username + ".ovpn"
with open(filename, "w") as file_out:
    file_out.write(text)

exit()

# Подключаемся по ssh

shell = spur.SshShell(hostname=conf.get("main", "SERVER"), username=conf.get("main", "USER"), password=conf.get("main", "PASSWORD"))
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
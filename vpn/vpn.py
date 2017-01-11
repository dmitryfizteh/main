#!/usr/bin/env python
# -*- coding: utf-8 -*-

import configparser
import paramiko
import subprocess
import os
import sys
import logging
import re
#import ldap3

# Добавляем логгирование событий
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

conf = configparser.RawConfigParser()
conf.read("vpn.ini")
# TODO: убрать пароль из файла и сделать его запрос

# TODO: убрать avedrov из параметров командной строки
# Считываем username из параметров командной строки и запрашиваем только если не указан
if len(sys.argv) > 1:
    username = sys.argv[1]
else:
    username = input("Имя пользователя, для которого создаются ключи доступа: ")
logging.info("Имя пользователя: "+ username)

mail_p = username + "@" + conf.get("mail", "MAIL_SUFF")
mail_i = input("E-mail пользователя, на который будут отправлены ключи доступа(по умолчанию " + mail_p + "): ")
if mail_i != "":
    mail = mail_i
else:
    mail = mail_p
check = re.search(r'^[a-zA-Z0-9]{1,100}[@][a-z]{2,6}\.[a-z]{2,4}', mail)
if check:
    logging.info("E-mail пользователя: " + mail)
else:
    logging.error("Ошибочно введен e-mail пользователя: " + mail)
    exit()

# Подключаемся по ssh, создаем и скачиваем ключи
logging.debug(u'Подключаемся по ssh, создаем и скачиваем ключи')
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

ssh.connect(hostname=conf.get("main", "SERVER"), username=conf.get("main", "USER"), password=conf.get("main", "PASSWORD")) #, key_filename='<path/to/openssh-private-key-file>')

# TODO: реализовать команду создания ключей
stdin, stdout, stderr = ssh.exec_command('ls /etc/openvpn/easy-rsa')
data = stdout.read() + stderr.read()
print(data)

sftp = ssh.open_sftp()

source = '/etc/openvpn/keys/' + username + '.key'
destination_key = username + '.key'
sftp.get(source, destination_key)

source = '/etc/openvpn/keys/' + username + '.crt'
destination_crt = username + '.crt'
sftp.get(source, destination_crt)

sftp.close()
ssh.close()

# Создаем из шаблона ovpn-файл
with open('keys/client.ovpn') as file_in:
    text = file_in.read()
text = text.replace("username", username)
filename = username + ".ovpn"
with open(filename, "w") as file_out:
    file_out.write(text)

# Создаем архив с ключами
# TODO: сделать одноранговые файлы в архиве
rc = subprocess.call(['7z', 'a', '-p12345', '-y', '-mhe=on', 'client.7z'] + ['keys/ca.crt','keys/ta.key', filename, destination_key, destination_crt])

#os.remove(filename)
#os.remove(destination_key)
#os.remove(destination_crt)
#os.remove('client.7z')

# TODO: Добавить пользователя в группу AD

exit()

# TODO: организовать отправку на почту или/и сохранение на рабочий стол
# отправляем ключи на почту
'''
import smtplib
from email.MIMEText import MIMEText

# отправитель
me = conf.get("mail", "SENDER")
# получатель
you = conf.get("mail", "RECIPIENT")
# текст письма
text = 'В приложении инструкция и архив с ключами доступа'
# заголовок письма
subj = 'OpenVPN'

# SMTP-сервер
server = conf.get("mail", "MAIL_SERVER")
port = conf.get("mail", "PORT")
user_name = conf.get("main", "USER")
user_passwd = conf.get("main", "PASSWORD")

# формирование сообщения
msg = MIMEText(text, "", "utf-8")
msg['Subject'] = subj
msg['From'] = me
msg['To'] = you

# отправка
s = smtplib.SMTP(server, port)
#s.ehlo()
#s.starttls()
#s.ehlo()
s.login(user_name, user_passwd)
s.sendmail(me, you, msg.as_string())
s.quit()
'''
print("X:/university/reorganization/Материалы для участников программы развития компетенций/Рабочее пространство участников программы")
exit
import os
os.chdir("X:/university/reorganization/Материалы для участников программы развития компетенций/Рабочее пространство участников программы")

f = open('users.txt')
line = f.readline()
for line in f.readlines():
    line = line[0:-1]
    print(line)
    os.mkdir(line)
f.close()

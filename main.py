import platform
import socket
from getmac import get_mac_address
import getpass
import os
import psutil
if platform.uname().system == 'Windows':
    import winapps
    import wmi
    w = wmi.WMI()
global list
list=[]

"""
def main():
    global path
    path= "c:/systeminfo"
    for BIOSs in w.Win32_ComputerSystem():
        UserNames=BIOSs.Caption
    fileName=path+os.path.sep+UserNames+".txt"
    info()

# Оценить, существует ли папка (путь)
    if not os.path.exists(path):
        print ("не существует")
# Создать папку (путь к файлу)
        os.makedirs(path)
# Записать информацию о файле
        with open(fileName,'w+') as f:
            for li in list:
                print(li)
                l=li+"\n"
                f.write(l)
    else:
        print ("существует")
        with open(fileName,'w+') as f:
            for li in list:
                print(li)
                l=li+"\n"
                f.write(l)
"""
def info():
    list.append("Информация о компьютере")
    list.append("Операционная система: % s" % platform.uname().system)
    list.append("Версия ОС: % s" % platform.uname().version)
    list.append("Имя компьютера: % s" % platform.node())
    list.append("IP адресс ПК: % s" % socket.gethostbyname_ex(socket.gethostname())[-1][-1])
    list.append("MAC адресс: % s" % get_mac_address(ip = socket.gethostbyname_ex(socket.gethostname())[-1][-1]))
    list.append("Текущий пользователь: % s" % getpass.getuser())

    if platform.uname().system == 'Windows':
        indx = 0

        for disk in w.Win32_DiskDrive(InterfaceType="IDE"):
            diskSize = int(disk.size)

            for x in w.Win32_PhysicalMedia():
                tagReplacedSlashes = x.Tag.replace("\\", "")
                devname = "PHYSICALDRIVE" + str(indx)
                if devname in tagReplacedSlashes:
                    sernum = x.SerialNumber

            list.append("Марка диска:% s" % disk.Caption)
            list.append("Серийный номер диска: % s" % sernum)
            list.append("Размер диска:% .2fGB" % (diskSize / 1024 ** 3))

        for li in list:
            print(li)
            l = li + "\n"
        print("\n" + "Установленное ПО:")
        for app in winapps.list_installed():
             print(app)


def main():
    info()

if __name__ == '__main__':
    main()
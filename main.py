import platform
import socket
from getmac import get_mac_address
import getpass
import winapps

global list
list=[]

"""
import wmi
import os
w = wmi.WMI()
global list
list=[]

def info():
    list.append("Информация о компьютере")
    for BIOSs in w.Win32_ComputerSystem():
        list.append ("Имя компьютера:% s"% BIOSs.Caption)
        list.append ("Пользователь:% s"% BIOSs.UserName)
    for address in w.Win32_NetworkAdapterConfiguration(ServiceName = "e1dexpress"):
        list.append ("IP-адрес:% s"% address.IPAddress [0])
        list.append ("MAC-адрес:% s"% address.MACAddress)
    for BIOS in w.Win32_BIOS():
        list.append ("Дата использования:% s"% BIOS.Description)
        list.append ("Модель материнской платы:% s"% BIOS.SerialNumber)
    for processor in w.Win32_Processor():
        list.append ("Модель ЦП:% s"% processor.Name.strip ())
    for memModule in w.Win32_PhysicalMemory():
        totalMemSize=int(memModule.Capacity)
        list.append ("Производитель памяти:% s"% memModule.Manufacturer)
        list.append ("Модель памяти:% s"% memModule.PartNumber)
        list.append ("Размер памяти:% .2fGB"% (totalMemSize / 1024 ** 3))
    for disk in w.Win32_DiskDrive(InterfaceType = "IDE"):
        diskSize=int(disk.size)
        list.append ("Имя диска:% s"% disk.Caption)
        list.append ("Размер диска:% .2fGB"% (diskSize / 1024 ** 3))
    for xk in w.Win32_VideoController():
        list.append ("Имя видеокарты:% s"% xk.name)

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

    for li in list:
        print(li)
        l = li + "\n"

    if platform.uname().system == 'Windows':
        print("\n" + "Установленное ПО:" )
        for app in winapps.list_installed():
            print(app)



def main():
    info()

main()
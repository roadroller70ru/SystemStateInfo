import platform
import socket
from getmac import get_mac_address
import getpass
import pandas as pd
import xlsxwriter
if platform.uname().system == 'Windows':
    import winapps
    import wmi
    w = wmi.WMI()
global list
global dict
list=[]
pc_dict={}

def info():
    pc_dict.clear()
    pc_dict.update({"Операционная система:" : platform.uname().system})
    pc_dict.update({"Версия ОС:" : platform.uname().version})
    pc_dict.update({"Имя компьютера:" : platform.node()})
    pc_dict.update({"IP адресс ПК:": socket.gethostbyname_ex(socket.gethostname())[-1][-1]})
    pc_dict.update({"MAC адресс:": get_mac_address(ip = socket.gethostbyname_ex(socket.gethostname())[-1][-1])})
    pc_dict.update({"Текущий пользователь:": getpass.getuser()})

    if platform.uname().system == 'Windows':
        indx = 0

        for disk in w.Win32_DiskDrive(InterfaceType="IDE"):
            diskSize = int(disk.size)

            for x in w.Win32_PhysicalMedia():
                tagReplacedSlashes = x.Tag.replace("\\", "")
                devname = "PHYSICALDRIVE" + str(indx)
                if devname in tagReplacedSlashes:
                    sernum = x.SerialNumber

            pc_dict.update({"Марка диска:" : disk.Caption})
            pc_dict.update({"Серийный номер диска:" : sernum})
            pc_dict.update({"Размер диска:" : (diskSize / 1024 ** 3)})

        pc_df = pd.DataFrame(pc_dict.items())

        for app in winapps.list_installed():
            list.append(app)

        po_df = pd.DataFrame({'col' : list})

        s_sheets = {'Система': pc_df, 'ПО': po_df}
        writer = pd.ExcelWriter('C:/tmp/test.xlsx', engine='xlsxwriter')

        for sheet_name in s_sheets.keys():
            s_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

        writer.save()



def main():
    info()

if __name__ == '__main__':
    main()
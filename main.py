import platform
import socket
from getmac import get_mac_address
import getpass
import pandas as pd
import re
import xlsxwriter
if platform.uname().system == 'Windows':
    import winapps
    import wmi
    w = wmi.WMI()
global list
global dict
lnameapp=[]
lverapp=[]
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
            #print(app)
            #InstalledApplication(name='Windows Driver Package - DEVLINE LTD (DVRSE) Media  (05/13/2015 )', version='05/13/2015 ', install_date=None, install_location=None, install_source=None, modify_path=None, publisher='DEVLINE LTD', uninstall_string='C:\\PROGRA~1\\DIFX\\B60D12~1\\dpinst.exe /u C:\\Windows\\System32\\DriverStore\\FileRepository\\dvrs.inf_amd64_850a0d787d994d4b\\dvrs.inf')
            res = re.split(r"'", str(app))
            i = 0
            while i < len(res):
                #     print(res[i])

                if re.search(r'name=', res[i]):
                    #print('name= ' + res[i + 1])
                    lnameapp.append(res[i+1])
                    break
                i += 1

            while i < len(res):
                if re.search(r'version=None', res[i]):
                    #print('version= ' + res[i + 1])
                    lverapp.append("None")
                    break
                else:
                    if re.search(r'version=', res[i]):
                        # print('version= ' + res[i + 1])
                        lverapp.append(res[i+1])
                        break
                i += 1
            #list.append(appname)

        #print(len(lnameapp))
        #print(len(lverapp))

        po_df = pd.DataFrame({'AppName' : lnameapp, 'Version' : lverapp})

        s_sheets = {'Система': pc_df, 'ПО': po_df}
        writer = pd.ExcelWriter('C:/tmp/test.xlsx', engine='xlsxwriter')

        for sheet_name in s_sheets.keys():
            s_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

        writer.save()

def main():
    info()

if __name__ == '__main__':
    main()
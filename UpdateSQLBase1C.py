import os
import time
import psutil
import pythoncom
import win32com.client
from pywinauto.application import Application
from pywinauto.keyboard import send_keys, KeySequenceError
import pywinauto

prm1 = [ '"C:\\Program Files\\1cv8\\common\\1cestart.exe"', "CONFIG", "/S" ]
l_fld = "D:\\DST\\1Clog\\"
pathZup = "D:\\DST\\HRM\\1cv8.cfu"
pathBuh = "D:\\DST\\BUH\\1cv8.cfu"
pathUT = "D:\\DST\\UT\\1cv8.cfu"
rez="Обновление конфигурации успешно завершено"
bs = 0
global fl
global cl

class base1C(object):
    def __init__(self, srv, base, usr, passw):
        Conn = "Srvr={0};Ref={1};Usr={2};Pwd={3};".format(srv, base, usr, passw)
        self.DBconnect(Conn)
        self.srv = srv
        self.base = base
        self.usr = usr
        self.passw = passw
        self.conf = self.V83.Metadata.Name
        self.ver = self.V83.Metadata.Version
                 
    def DBconnect(self, Conn):
        pythoncom.CoInitialize()
        self.V83 = win32com.client.Dispatch("V83.COMConnector").Connect(Conn)

    def PathUpd(self):
        if self.conf == "ЗарплатаИУправлениеПерсоналом":
            return pathZup
        elif self.conf == "БухгалтерияПредприятия":
            return pathBuh
        elif self.conf == "УправлениеТорговлей":
            return pathUT
        else:
            return "Неизвестная конфигурация"

def kill_1C():
    for n in psutil.process_iter():
        pn = n.as_dict(attrs = ['pid', 'name', 'username'])
        if (pn['name'] == "1cv8.exe") and (pn['username'] == 'SRV-1C\\admin') :
            pk = psutil.Process(pn['pid'])
            pk.kill()
            print(pk.status())

#def kill_rphost():
#    for n in psutil.process_iter():
#        pn = n.as_dict(attrs = ['pid', 'name', 'username'])
#        if (pn['name'] == "rphost.exe"):
#            pk = psutil.Process(pn['pid'])
#            pk.kill()
#            print(pk.status())

def window0():    #cравнение версий конфигураций  
    try:
        app = Application().connect(title=u'\u041e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u0435 \u043a\u043e\u043d\u0444\u0438\u0433\u0443\u0440\u0430\u0446\u0438\u0438', class_name='V8NewLocalFrameBaseWnd')
        if 'app' in locals():
            vnewlocalframebasewnd = app.V8NewLocalFrameBaseWnd
            vnewlocalframebasewnd.set_focus()
            send_keys("{ENTER}")
            print("подтверждаю версию обновления")
            return True
    except pywinauto.findwindows.ElementNotFoundError:
        print("нет окна сравнения версий конфигурации")
        return False

def window1():  #обновить конфигурацию базы данных?
    try:
        app = Application().connect(title=u'\u041a\u043e\u043d\u0444\u0438\u0433\u0443\u0440\u0430\u0442\u043e\u0440', class_name='V8NewLocalFrameBaseWnd')
        if 'app' in locals():
            vnewlocalframebasewnd = app.V8NewLocalFrameBaseWnd
            vnewlocalframebasewnd.set_focus()
            send_keys("{ENTER}")
            return True
    except pywinauto.findwindows.ElementNotFoundError:
        print("Окно подтверждения не найдено")
        return False

def window2():  #завершить все соединения с базой данных
    try:
        app = Application().connect(title=u'\u041a\u043e\u043d\u0444\u0438\u0433\u0443\u0440\u0430\u0442\u043e\u0440', class_name='V8NewLocalFrameBaseWnd')
        if 'app' in locals():
            vnewlocalframebasewnd = app[u'\u041a\u043e\u043d\u0444\u0438\u0433\u0443\u0440\u0430\u0442\u043e\u0440']
            vnewlocalframebasewnd.set_focus()
            send_keys("{TAB}")
            send_keys("{ENTER}")
            return True
    except pywinauto.findwindows.ElementNotFoundError:
        print("Окно повтора или аварийного завершения всех сеаносв с базой")
        return False

def window3():  #подтвердить аварийное завершение
    try:
        app = Application().connect(title=u'\u041a\u043e\u043d\u0444\u0438\u0433\u0443\u0440\u0430\u0442\u043e\u0440', class_name='V8NewLocalFrameBaseWnd')
        if 'app' in locals():
            vnewlocalframebasewnd = app[u'\u041a\u043e\u043d\u0444\u0438\u0433\u0443\u0440\u0430\u0442\u043e\u0440']
            vnewlocalframebasewnd.set_focus()
            send_keys("{TAB}")
            send_keys("{ENTER}")
            return True
    except pywinauto.findwindows.ElementNotFoundError:
        print("Окно подтверждения аварийного завершения не найдено")
        return False

def window4():  #реорганизация информации
    try:
        app_r = Application().connect(title=u'\u0420\u0435\u043e\u0440\u0433\u0430\u043d\u0438\u0437\u0430\u0446\u0438\u044f \u0438\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u0438', class_name='V8NewLocalFrameBaseWnd')
        if 'app_r' in locals():
            vnewlocalframebasewnd_r = app_r.V8NewLocalFrameBaseWnd
            vnewlocalframebasewnd_r.set_focus()
            send_keys("{ENTER}")
            return True
    except pywinauto.findwindows.ElementNotFoundError:
        print("Окно Реорганизация информации не найдено")
        return False

def AlreadyUpdated():
    try:
        app = Application().connect(title=u'\u0424\u0430\u0439\u043b \u043d\u0435 \u0441\u043e\u0434\u0435\u0440\u0436\u0438\u0442 \u0434\u043e\u0441\u0442\u0443\u043f\u043d\u044b\u0445 \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u0439', class_name='V8NewLocalFrameBaseWnd')
        if 'app' in locals():
            vnewlocalframebasewnd = app.V8NewLocalFrameBaseWnd
            vnewlocalframebasewnd.set_focus()
            vnewlocalframebasewnd.close()    
            return True
    except pywinauto.findwindows.ElementNotFoundError:
        return False

def compareVer(versB, pathB):
    infoB = pathB.replace("1cv8.cfu", "UpdInfo.txt")
    updInfo = open(infoB, "r")
    infoList = updInfo.readlines()
    updVer = (infoList[0].split("="))[1]
    updFrom = (infoList[1].split("="))[1].split(";")
    vers1 = versB.split(".")
    vers2 = updVer.split(".")
    status = 0
    updInfo.close()
    for n in range(len(updFrom)):
        if updFrom[n] == versB :
            status = 1
            break
    if status == 0:    
        print("Неверная версия конфигурации для обновления")
        return False
    if int(vers1[1]) < int(vers2[1]) or int(vers1[2]) < int(vers2[2]) or int(vers1[3]) < int(vers2[3]):
        return True
    else:
        return False
 
listBases = []
bases = open("bases.txt", "r")
lines = bases.readlines()
for i in lines:
    s = i.split(",")
    b = base1C(s[0], s[1], s[2], s[3].rstrip())
    listBases.append(b)
bases.close()

for j in range(len(listBases)):
    print(listBases[j].base, "  ", listBases[j].ver, "  ", listBases[j].conf)
ask1 = input("Введите Y если обновления для баз скопированы в нужные каталоги и можно продолжить \n Нажмите любую другую клавишу чтобы прервать выполнение\n Ответ:")
if ask1 == "y":
    pass
    ask2 = input("Введите Y если нужно закрыть все запущенные платформы 1С на сервере и продолжить \n Нажмите любую другую клавишу чтобы продолжить без закрытия\n Ответ:")
    if ask2 == "y":
        kill_1C()
#    ask3 = input("Введите Y если нужно сбросить все соединения на сервере 1С и продолжить \n Нажмите любую другую клавишу чтобы продолжить без сброса\n Ответ:")
#    if ask3 == "y":
#        kill_rphost()
else:
    quit()


    
for i in range(len(listBases)):
    if compareVer(listBases[i].ver, listBases[i].PathUpd()):  
        res = ' '.join(prm1) + " " + str(listBases[i].srv) + "\\" + str(listBases[i].base) + " /N " + listBases[i].usr + " /P " + listBases[i].passw + " /UpdateCfg " + str(listBases[i].PathUpd()) + " /UpdateDBCfg /OUT " + l_fld + str(listBases[i].base) + ".log"
        print(res, "-------------------------------------------------------------------------------------------------------------------", listBases[i].conf)
        print("начато обновление базы ", listBases[i].base )  
        os.system(res)
        time.sleep(20)
        fl = 1
        cl = 0
        while fl :
            print("начат цикл обоновления, условие = " + str(fl))
            for n in psutil.process_iter():
                pn = n.as_dict(attrs = ['pid', 'name', 'username'])
                if (pn['name'] == "1cv8.exe") and (pn['username'] == 'SRV-1C\\Admin') :
                    time.sleep(10)
                    print("процесс 1C найден")
                    if cl == 0:
                        cl = 1 #запускаем процесс проверки окон
            if cl == 1:
                if window0():
                    cl = 2
                elif AlreadyUpdated():
                    fl = 0
            elif cl == 2:
                if window1():
                    cl = 3
            elif cl == 3:
                if window2():
                    cl = 4
            elif cl == 4:
                if window3():
                    cl = 5
            elif cl == 5:
                if window4():
                    cl = 6
            elif cl == 6:        
                log_n = l_fld + listBases[i].base + ".log"
                log_cur = open(log_n).readlines()
                for str_l in iter(log_cur):
                    if rez in str_l:
                        fl = 0
                        print("обновление", listBases[i].base, "завершено")  
            time.sleep(10)
time.sleep(10)



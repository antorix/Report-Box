#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
Root_path = os.path.dirname(os.path.abspath(__file__)) + "\\" # корневая папка, в которой лежат подпапки возвещателей и файлы программы
Subpath = []
Subpath.append(Root_path + "Возвещатели\\") # подпапки с файлами бланков S-21, здесь их при желании можно переименовать, но их должно быть только 3
Subpath.append(Root_path + "Подсобные пионеры\\")
Subpath.append(Root_path + "Общие пионеры\\")
Version = 1 # это самая новая версия!
Update = True # проверка обновлений, выберите False, чтобы отключить
Filename = Root_path + "publishers.txt" # файл с собственной базой данных возвещателей
Values = []
Max_Name_Len = 0
Docmode = True # документированный режим, в общем доступе всегда True


import webbrowser
from os.path import isfile, join
from shutil import copyfile
import time
import datetime
import requests
import _thread
from sys import argv
if not Docmode: import extras as ex
try:
    import keyboard
except:
    from subprocess import check_call
    from sys import executable
    check_call([executable, '-m', 'pip', 'install', 'keyboard'])
    import keyboard

if "dev" in argv: # переопределение переменных из системных параметров
    Devmode = True
    Docmode = False
    print("Dev mode on")
else: Devmode = False

###

def load():
    """ Загрузка базы возвещателей из файла """
    publishers = []
    #global Max_Name_Len
    if not os.path.exists(Filename):
        print("Файл базы данных не найден.")
    else:       
        with open(Filename, "r", encoding="utf-8") as f: lines=[line for line in f]
        for i in range(len(lines)):
            publishers.append(["", 0, 0, 0])
            try:
                publishers[i][0] = lines[i][0: lines[i].index("\t")]  # присвоение имени
            except:
                publishers[i][0] = lines[i]  # присвоение имени, если нет статистики
            tPos = []
            for y in range(len(lines[i])):  # выяснение положений табуляций
                if lines[i][y] == "\t" or lines[i][y] == "\n": tPos.append(y)
            for length in range(4):
                if len(tPos) == length:
                    for m in range(4 - length):
                        tPos.append(0)
            try: publishers[i][1] = int(lines[i][tPos[0]: tPos[1]].strip())  # часы
            except: publishers[i][1] = 0
            try: publishers[i][2] = int(lines[i][tPos[1]: tPos[2]].strip())  # изучения
            except: publishers[i][2] = 0
            try: publishers[i][3] = int(lines[i][tPos[2]: tPos[3]].strip())  # кредит
            except: publishers[i][3] = 0
            #if len(publishers[i][0]) > Max_Name_Len: Max_Name_Len = len(publishers[i][0])

    return publishers
    
def save():
    """ Выгрузка базы данных в файл """
    try: Pub.sort(key=lambda x: x[0])
    except: pass
    with open(Filename, "w", encoding="utf-8") as datafile:
        for row in Pub:
            datafile.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\n")

def fetch(command):
    """ Интерпретация строки отчета - перевод простой строки в список из имени возвещателя и 4 параметров """
    for delimeter in [" ", "\t"]:
        try:
            if delimeter in command:
                name=command[: command.index(delimeter)]
                stats=command[command.index(delimeter)+1 :]
                tPos=[0]
                for y in range(len(stats)):  # выяснение положений табуляций
                    if stats[y]==delimeter: tPos.append(y)
                if len(tPos)    == 1:   # введены только часы|служение
                    stat1 = stats[tPos[0]: ].strip()            # часы
                    stat2 = stat3 = 0
                elif len(tPos)  == 2:   # введены часы и изучения
                    stat1 = stats[tPos[0]: tPos[1]].strip()     # часы
                    stat2 = stats[tPos[1]: ].strip()            # изучения
                    stat3 = 0
                elif len(tPos)  == 3:   # введены часы, изучения, кредит
                    stat1 = stats[tPos[0]: tPos[1]].strip()     # часы
                    stat2 = stats[tPos[1]: tPos[2]].strip()     # изучения
                    stat3 = stats[tPos[2]: ].strip()     # кредит
                result = [name, int(stat1), int(stat2), int(stat3)]
                break
            else: result = [command.strip(), 0, 0, 0]
        except:
            print("Значения отчета должны содержать только цифры!\n"+\
                  "Попробуйте еще раз.")
            result = [command.strip(), 0, 0, 0]
    return result

def Pub_delete(line):
    """ Удаление возвещателя """
    try:
        os.remove(Pub[line][0])
        del Pub[line]
        save()
    except: print("Удаление не сработало. Возможно, этот файл открыт.\n"+\
                  "Закройте его и попробуйте еще раз.")
    else: print("Возвещатель успешно удален.")

def Pub_add(string):
    """ Создание нового возвещателя """
    S21 = Root_path + "S-21_U.pdf"
    name = string[1:]
    type = getPath()
    newFileName = f"{type}{name}.pdf"
    for p in Pub:
        if p[0] == newFileName:
            print("Уже есть возвещатель такого типа с таким именем!")
            return
    if type is not None:
        while 1:
            if os.path.exists(S21):

                copyfile(S21, newFileName)
                Pub.append([newFileName, 0, 0, 0])
                print(f"Возвещатель '{newFileName}' успешно создан.")
                save()
                break
            else:
                print("Не найден бланк S-21! Укажите местоположение этого\n"+\
                      "бланка в формате PDF...")
                from tkinter import filedialog
                S21 = filedialog.askopenfilename()
                copyfile(S21, f"{Root_path}\\S-21_U.pdf") # копируем пустой бланк S-21 в папку программы на будущее

def search(myinput, process=True):
    """ Поиск возвещателя и ввод его данных. Возвращает индекс строки в массиве Pub с этим возвещателем """
    name=myinput[0]
    print("Ищу %s..." % name)
    line = None
    found=[]

    for i in range(len(Pub)):
        entry = format_title(Pub[i][0].strip(), cut_path=True)
        if name.lower() in entry.lower():
            print(f"Найдено: {'{:>3}│'.format(str(len(found)+1))} {format_title(Pub[i][0].strip())} {format_report_string(Pub[i])}")
            line=i
            found.append(line)

    if len(found)==0: # если запрошенного имени нет
        print("Ничего не найдено.")

    elif len(found)>1: # если запрошенных вариантов несколько
        while 1:
            try:
                value = input("Введите номер варианта или Enter для отмены: ")
                if value=="":
                    cls()
                    return None
                elif not Docmode and (value=="groups" or value=="!"):
                    webbrowser.open("https://docs.google.com/spreadsheets/d/1POVv-2nM4rGd6-MOITwA-0hNHmqHe9QH/edit#gid=1676049882")
                else:
                    line=found[int(value)-1]
                    print(f"Ваш выбор:\n{format_title(Pub[line][0])} {format_report_string(Pub[line])})")
                    break
            except:
                print("Ошибка, попробуйте еще раз.")
                continue

    if line is not None and myinput[1]+myinput[2]+myinput[3] == 0: # открытие возвещателя без ввода статистики, с меню для выбора
        try:
            option = int(input("Что нужно сделать (введите номер варианта):\n1│ Открыть PDF-файл\n2│ Обнулить отчет\n3│ Переименовать/переместить\n4│ Удалить\n"))
        except: option = 0
        if option == 1:
            print("Открываю PDF-файл...")
            webbrowser.open(Pub[line][0])
        elif option == 2:
            Pub[line][1] = Pub[line][2] = Pub[line][3] = 0
            print("Отчет возвещателя обнулен.")
            save()
        elif option == 3:
            new_name = input("Введите новое название файла для этого возвещателя:\n")
            Pub_rename(line, new_name)
        elif option == 4:
            Pub_delete(line)
        else: cls()

    elif line is not None and process: # открытие возвещателя с вводом статистики
        Pub[line] = [Pub[line][0], myinput[1], myinput[2], myinput[3]]
        save()
        auto = ". Поставьте курсор мыши в\nполе «Часы» и нажмите на кнопку «Insert»..." if Auto else "..."
        print(f"База обновлена, открываю PDF-файл{auto}")
        webbrowser.open(Pub[line][0])
        for i in myinput[1:5]: Values.append(i)

    return line

def recreate(confirm=False):
    """ Генерация новой базы данных возвещателей из папок """
    result = False
    choice = 0
    if confirm:
        try: choice = int(input(
            "Внимание! Данная операция заново создает базу данных программы\n"+\
            "путем сканирования папок с PDF-файлами. При этом все отчеты\n"+\
            "за месяц удаляются. Текущий файл базы данных не удаляется, а\n"+\
            "сохраняется в папке «Архив». Обычно это стоит делать в начале\n"+\
            "нового месяца. Продолжать?\n[1] Да\n[0] Нет\n"))
        except: pass
    if not confirm or choice == 1:
        files = []
        for p in Subpath:
            if os.path.exists(p):
                files += [p + f for f in os.listdir(p) if isfile(join(p, f))]
        files.sort(key=lambda x: x[0])
        if os.path.exists(Filename):
            savedTime = time.strftime("%Y-%m-%d_%H%M%S", time.localtime())
            if not os.path.exists(f"{Root_path}\\Архив"):
                os.makedirs(f"{Root_path}\\Архив")
            copyfile(Filename, f"{Root_path}\\Архив\\publishers-{savedTime}.txt")
        with open(Filename, "w", encoding="utf-8") as datafile:
            print("База данных сгенерирована из папок.")
            for row in files:
                datafile.write(f"{row}\t0\t0\t0\n")
        result = True
    return result

def stats(mode=1):
    """ Подсчет и вывод статистики для ввода на jw.org (команда =) """
    global Values
    def printStat(): # вывод статистики
        print("Активные возвещатели: %d\n" % len(Pub))
        print("Возвещатели:\nКоличество отчетов: %d\nЧасы: %d\nИзучения Библии: %d\n" % (countP, countP_Hours, countP_St))
        print("Подсобные пионеры:\nКоличество отчетов: %d\nЧасы: %d\nИзучения Библии: %d\n" % (countAP, countAP_Hours, countAP_St))
        print("Общие пионеры:\nКоличество отчетов: %d\nЧасы: %d\nИзучения Библии: %d" % (countRP, countRP_Hours, countRP_St))
    countP=countP_Hours=countP_St=countAP=countAP_Hours=countAP_St=countRP=countRP_Hours=countRP_St=0
    for row in Pub:
        if "Подсобные" in row[0] and row[1]!="" and row[1]!="0":
            countAP+=1
            countAP_Hours+=float(row[1])
            countAP_St+=int(row[2])
        elif "Общие" in row[0] and row[1]!="" and row[1]!="0":
            countRP+=1
            countRP_Hours+=float(row[1])
            countRP_St+=int(row[2])
        elif "Возвещатели" in row[0] and row[1]!="" and row[1]!="0":
            countP+=1
            countP_Hours+=float(row[1])
            countP_St+=int(row[2])
    if mode==1: # показ статистики
        printStat()
    elif mode==2: # ввод статистики на сайте
        printStat()
        webbrowser.open("https://hub.jw.org/congregation-reports/ru/a1b222ff-99b5-4021-8b12-be9bafe86234/overview")
        print("Жду окончания ввода...")
        Values = [    str(len(Pub)),
                    str(int(countP)),  str(int(countP_Pub)),    str(int(countP_Vid)),   str(int(countP_Hours)),  str(int(countP_Ret)),  str(int(countP_St)),
                    str(int(countAP)), str(int(countAP_Pub)),   str(int(countAP_Vid)),  str(int(countAP_Hours)), str(int(countAP_Ret)), str(int(countAP_St)),
                    str(int(countRP)), str(int(countRP_Pub)),   str(int(countRP_Vid)),  str(int(countRP_Hours)), str(int(countRP_Ret)), str(int(countRP_St))]
        Values = []

def insert_into_PDF():
    """ Вставка данных в PDF-файл """
    global Values
    if Auto:
        for i in range(len(Values)):
            if i == 0: # часы
                keyboard.write(str(Values[i]) if Values[i] != 0 else "")
                if Values[1] != 0 or Values[2] != 0:
                    keyboard.press("\t")
                    keyboard.press("\t") # табулируем еще раз, чтобы пропустить повторные посещения
            elif i == 1: # изучения
                keyboard.write(str(Values[i]) if Values[i] != 0 else "")
                if Values[2] != 0: keyboard.press("\t")
            elif i == 2: # кредит
                keyboard.write(f"кредит {str(Values[i])}" if Values[i] != 0 else "")
    Values = []

def getPath():
    """ Спрашивает категорию возвещателя и возвращает путь к соответствующей папке """
    path = []
    for p in Subpath:
        last = len(p)-1
        path.append(p[ len(Root_path) : last])
    try:
        cat = int(input(f"Выберите тип возвещателя (введите номер варианта):\n1│ {path[0]}\n2│ {path[1]}\n3│ {path[2]}\n"))
    except: cat = 0
    if cat==1 or cat==2 or cat==3: return Subpath[cat-1]
    else:
        cls()
        return None

def Pub_rename(line1, Pub2):
    # Переименовывание возвещателя.
    # Принимает: line1 = номер строки переименуемого возвещателя; Pub2 - новое имя, в которое он будет переименован
    path = getPath()
    if path is not None:
        oldPubPath = Pub[line1][0]
        newPubPath = f"{path}{Pub2}.pdf"
        try:
            os.rename(oldPubPath, newPubPath)
            Pub[line1][0] = newPubPath
        except:
            print(f"Ошибка операции с файлами. Скорее всего, какой-то из файлов открыт.")
            return
        else:
            save()
            print(f"База успешно обновлена, новый файл возвещателя:\n'{newPubPath}'.")

def list():
    """ Вывод полного списка возвещателей """
    count=0
    print(f"Полный список активных возвещателей:")
    for p in Pub:
        print(f"{'{:>3}│'.format(str(count+1))} {format_title(p[0])} {format_report_string(p)}")
        count+=1

def remaining():
    """ Вывод возвещателей, не сдавших отчеты, и их подсчет (команда *) """
    count = 0
    print(f"Не сдали отчеты:")
    for i in range(len(Pub)):
        if Pub[i][1] + Pub[i][2] + Pub[i][3] == 0:
            print(f"{'{:>3}│'.format(str(count+1))} {format_title(Pub[i][0])}")
            count+=1

def format_title(name, cut_path=False):
    """ Форматирование и сокращение пути к файлу с именем возвещателя """
    cut = len(Root_path)
    name = name[cut:]
    string = ""
    if cut_path:
        string = name[ name.index("\\")+1 : name.index(".pdf")]
    else:
        for char, i in zip(name, range(len(name))):
            if name[i:i+4] == ".pdf": break
            elif char != "\\":
                string += char
            else: string += ": "
    #return '{:<25}'.format(string)
    space = ""
    for a in range(Max_Name_Len - len(string)): space += " "
    return string + space

def format_report_string(p):
    """ Показывает расширение строки с возвещателем, если у него есть отчет (если нет, пусто) """
    string = ""
    #if p[1]+p[2]+p[3] > 0: string += space
    string += ("%3dч" % p[1]) if p[1] > 0 else "    "
    string += ("%3dи" % p[2]) if p[2] > 0 else "    "
    string += ("%4dк" % p[3]) if p[3] > 0 else "    "
    return string

def get_pub_reported():
    count = 0
    for p in Pub:
        if p[1] + p[2] + p[3] > 0: count += 1
    return '{:<20}'.format(count)

def cls():
    """ Очистка терминала и вывод приветственной плашки """
    if not Devmode:
        os.system('cls' if os.name == 'nt' else 'clear')
        update = str("{:%d.%m.%Y, %H:%M:%S}".format(datetime.datetime.strptime(time.ctime((os.path.getmtime(Filename))),"%a %b %d %H:%M:%S %Y")))
        print(f"┌───────────────────────────────────────────────────────────────┐")
        print(f"│ Добро пожаловать в Report Box!                                │")
        print(f"│                                                               │")
        print(f"│ Возвещателей в базе: {'{:<27}'.format(len(Pub))}              │")
        print(f"│ C отчетами: {get_pub_reported()}                              │")
        print(f"│ Последнее изменение базы: {update}                │")
        print(f"│                                                               │")
        print(f"│ Введите команду и нажмите Enter:                              │")
        print(f"│                                                               │")
        print(f"│ 1AA         выбор возвещателя (введите название файла)        │")
        print(f"│ 1AA 5       ввести отчет возвещателя (только часы)            │")
        print(f"│ 1AA 5 1     ввести отчет возвещателя (часы, изучения)         │")
        print(f"│ 1AA 5 0 10  ввести отчет возвещателя (часы, изучения, кредит) │")
        print(f"│ +2ББ        создать нового возвещателя                        │")
        print(f"│ :           список всех возвещателей в базе                   │")
        print(f"│ !           кто не сдал отчет                                 │")
        print(f"│ =           статистика собрания для общего отчета             │")
        print(f"│ *           перезагрузка базы данных                          │")
        print(f"│ ?           ничего не понимаю, что здесь происходит?          │")
        print(f"└───────────────────────────────────────────────────────────────┘")

def update():
    """ Проверяем новую версию и при наличии обновляем программу с GitHub """
    def __update(threadName, delay):
        try:
            for line in requests.get("https://raw.githubusercontent.com/antorix/Report-Box/master/version"):
                newVersion = int(line.decode('utf-8').strip())
        except: print("Не удалось проверить обновления.")
        else:
            print("Успешно подключились к серверу обновлений.")
            print(f"Версия на сайте: {newVersion}")
            if newVersion > Version:
                print("Найдена новая версия, скачиваем.")
                """response = requests.get("https://github.com/antorix/Rocket-Ministry/archive/refs/heads/master.zip")
                import tempfile
                file = tempfile.TemporaryFile()
                file.write(response.content)
                file.close()            
                #result = True"""
            else: print("Обновлений нет.")
    if Update: _thread.start_new_thread(__update, ("Thread-Update", 3,))

###

try: # проверяем, будет ли работать автоматизация клавиатуры (не работает на мобильных устройствах)
    keyboard.add_hotkey('insert', insert_into_PDF)
except:
    Auto = False
    print("Не удалось задействовать автоматизацию клавиатуры. Вам придется\n"+\
          "вводить значения в поля PDF-файлов вручную.")
else: Auto = True

for p in Subpath: # создаем папки, если их нет
    if not os.path.exists(p): os.makedirs(p)

print("Ищу базу данных...") # начальная загрузка базы данных
Pub=load()
if Pub==[] and not os.path.exists(Filename):
    print("Генерирую базу данных из папок...")
    recreate()
    Pub=load()

cls()

update()

while 1: # главный цикл программы

    Max_Name_Len = 0 # определение максимальной длины первой колонки в списках возвещателей
    for p in Pub:
        if len(format_title(p[0])) > Max_Name_Len:
            Max_Name_Len = len(format_title(p[0]))

    command=input("> ").strip()
    if command=="":
        cls()
        continue
    elif command=="?": # открытие справки в онлайне
        cls()
        print("Открываю онлайн-справку...")
        webbrowser.open("https://github.com/antorix/secretary/blob/master/readme.md")
    elif command=="=": # статистика
        cls()
        stats()
    elif command=="!": # кто не сдал отчет
        cls()
        remaining()
    elif command==":": # показ всех возвещателей
        cls()
        list()
    elif command[0]=="+" and len(command)>1: # добавление нового возвещателя
        cls()
        Pub_add(command)
    elif command[0]=="*":
        cls()
        result = recreate(confirm=True)
        if result: Pub = load()

    elif "doc=off" in command.lower(): # команды в недокументированном режиме (не описаны в документации)
        Docmode = False
        print("Doc mode off")
    elif command=="=jw" and not Docmode: # ввод статистики на сайте
        stats(mode=2)
    elif command=="file" and not Docmode: # открытие файла данных
        webbrowser.open(Filename)
    elif command=="import" and not Docmode: # ввод статистики из файла
        ex.excel()
    elif command=="load" and not Docmode: # перезагрузка из файла
        Pub=load()
        save()
    elif command=="groups" and not Docmode: # показ групп в онлайне
        webbrowser.open("https://docs.google.com/spreadsheets/d/1POVv-2nM4rGd6-MOITwA-0hNHmqHe9QH/edit#gid=1676049882")
    elif command=="folder" and not Docmode: # открытие папки файлов в проводнике
        webbrowser.open(Root_path)
    elif command=="share" and not Docmode: # открытие ShareFile
        webbrowser.open("https://watchtower.sharefile.com/home/shared/fo64dbb7-f1c2-4683-b2e6-f53959f02435")
    elif command=="picalc" and not Docmode: # подсчет статистики пионеров
        ex.piCalc()

    else:
        cls()
        search(fetch(command)) # обработка команды с указанием возвещателя

#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
Root_path = os.path.dirname(os.path.abspath(__file__)) + "\\" # корневая папка, в которой лежат подпапки возвещателей и файлы программы
Subpath = []
Subpath.append(Root_path + "Возвещатели\\") # подпапки с файлами бланков S-21
Subpath.append(Root_path + "Подсобные пионеры\\")
Subpath.append(Root_path + "Общие пионеры\\")
Version = 1
Filename = Root_path + "publishers.ini" # файл с собственной базой данных возвещателей
Values = []
Bullet = "•" #→●•○▪
Docmode = True # документированный режим, в котором работают только функции, описанные в документации

import webbrowser
from os.path import isfile, join
import shutil
import time
import tkinter
import tkinter.messagebox
import tkinter.filedialog
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
else:
    Devmode = False

###

def load():
    """ Загрузка базы возвещателей из файла """
    publishers = []
    print("Ищу базу данных в файле publishers.ini...")
    if not os.path.exists(Filename):
        print("Файл данных не найден.")
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
        print("Файл данных успешно загружен.")

    return publishers

def save():
    """ Выгрузка базы данных в файл """
    Pub.sort(key=lambda x: x[0])
    with open(Filename, "w", encoding="utf-8") as datafile:
        for row in Pub:
            datafile.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\n")

def fetch(command):
    """ Интерпретация строки отчета - перевод простой строки в список из имени возвещателя и 4 параметров """
    print(f"> {command}\nВот что нашлось по вашему запросу:")
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
    except:
        if os.path.exists(Pub[line][0]):
            print("Удаление не сработало. Возможно, этот файл открыт.\n"+\
                  "Закройте его и попробуйте еще раз.")
            return False
        else:
            del Pub[line] # если файла нет, удаляем только из базы
            print("Возвещатель успешно удален из базы данных, но его\nPDF-файл отсутствует в папках!")
            save()
    else:
        print("Возвещатель успешно удален.")
        save()

def nullify():
    """ Обнуление всех отчетов """
    choice = 0
    try:
        choice = int(input(
            "Внимание! Все отчеты возвещателей за отчетный месяц будут\n"+\
            "удалены. Перед этим копия файла данных будет сохранена в папке\n"+\
            "«Архив». Обычно это нужно делать в начале нового месяца.\n"+\
            "Продолжать?\n[1] Да\n[0] Нет\n"))
    except: cls()
    if choice == 1:
        if os.path.exists(Filename):
            savedTime = time.strftime("%Y-%m-%d_%H%M%S", time.localtime())
            if not os.path.exists(f"{Root_path}\\Архив"):
                os.makedirs(f"{Root_path}\\Архив")
            shutil.copyfile(Filename, f"{Root_path}\\Архив\\publishers-{savedTime}.ini")
        for p in Pub:
            p[1] = p[2] = p[3] = 0
        save()
        print("Отчеты успешно обнулены.")
    else: cls()

def Pub_add(string):
    """ Создание нового возвещателя """
    S21 = Root_path + "S-21_U.pdf"
    name = string[1:].strip()
    type = getPath()
    newFileName = f"{type}{name}.pdf"
    for p in Pub:
        if p[0] == newFileName:
            print("Уже есть возвещатель в этой категории с таким именем!")
            return
    if type is not None:
        while 1:
            if os.path.exists(S21):
                shutil.copyfile(S21, newFileName)
                Pub.append([newFileName, 0, 0, 0])
                print(f"Возвещатель '{newFileName}' успешно создан.")
                save()
                break
            else:
                print("Не найден бланк S-21! Укажите местоположение этого\n"+\
                      "бланка в формате PDF...")
                S21 = tkinter.filedialog.askopenfilename()
                if S21 != "":
                    shutil.copyfile(S21, f"{Root_path}\\S-21_U.pdf") # копируем пустой бланк S-21 в папку программы на будущее
                else:
                    print("Для создания возвещателей программе нужен бланк S-21 в\n"+\
                          "формате PDF. Вы знаете, где его найти.")
                    break

def search(myinput, process=True):
    """ Поиск возвещателя и ввод его данных. Возвращает индекс строки в массиве Pub с этим возвещателем """
    name=myinput[0]
    line = None
    found=[]
    max_len = 0

    for i in range(len(Pub)):
        entry = format_title(Pub[i][0].strip(), cut_path=True)
        if name.lower() in entry.lower(): # сначала просто находим строки и обсчитываем ширину колонки
            string = f"{format_title(Pub[i][0].strip())}"
            if len(string) > max_len: max_len = len(string)
            line=i
            found.append(line)

    for f, i in zip(found, range(len(found))): # выводим результаты
        if len(found) == 1:
            print(f"{Bullet} {format_title(Pub[f][0].strip(), max_len)} {format_report_string(Pub[f])}")
        elif len(found) > 1:
            print(f"{'{:>3}'.format(i+1)} {Bullet} {format_title(Pub[f][0].strip(), max_len)} {format_report_string(Pub[f])}")

    if len(found)==0: # если запрошенного имени нет
        print("Ничего не найдено.")

    elif len(found)>1: # если запрошенных вариантов несколько
        while 1:
            try:
                value = input("Введите номер варианта или Enter для отмены: ").strip()
                if value=="":
                    cls()
                    return None
                elif not Docmode and (value=="groups" or value=="!"):
                    webbrowser.open("")
                else:
                    line=found[int(value)-1]
                    print(f"Ваш выбор:\n{Bullet} {format_title(Pub[line][0])} {format_report_string(Pub[line])}")
                    break
            except:
                continue

    if line is not None and myinput[1]+myinput[2]+myinput[3] == 0: # открытие возвещателя без ввода статистики, с меню для выбора
        print("Меню действий:\n"+\
              "[1] Открыть PDF-файл\n[2] Ввести отчет\n[3] Обнулить отчет\n[4] Переименовать/переместить\n[5] Удалить")
        while 1:
            option = input("Введите номер варианта или Enter для отмены: ").strip()
            if option == "":
                cls()
                break
            else:
                try: option = int(option)
                except: continue
            if option == 1:
                print("Открываю PDF-файл...")
                if os.path.exists(Pub[line][0]):
                    webbrowser.open(Pub[line][0])
                else:
                    print("PDF-файл возвещателя отсутствует в базе!")
                break
            elif option == 2:
                print("Внимание! Отчеты вводятся так, как показано на подсказке по\n"+\
                      "командам выше. Если вам это непонятно, введите «?» и почитайте\n"+\
                      "онлайн-справку.")
                break
            elif option == 3:
                Pub[line][1] = Pub[line][2] = Pub[line][3] = 0
                print("Отчет возвещателя обнулен.")
                save()
                break
            elif option == 4:
                while 1:
                    new_name = input("Переименование возвещателя\nВведите новое название файла или Enter для отмены: ").strip()
                    if new_name != "":
                        Pub_rename(line, new_name)
                        break
                    else:
                        cls()
                        break
                break
            elif option == 5:
                result = Pub_delete(line)
                if result == False: continue # Pub_delete возвращает False, если нельзя удалить файл, потому что он открыт
                else: break

    elif line is not None and process: # открытие возвещателя с вводом статистики
        Pub[line] = [Pub[line][0], myinput[1], myinput[2], myinput[3]]
        save()
        print(f"База обновлена, открываю PDF-файл. Поставьте курсор мыши в\nполе «Часы» и нажмите на кнопку «Insert»...")
        webbrowser.open(Pub[line][0])
        for i in myinput[1:5]: Values.append(i)
    return line

def generate():
    """ Генерация новой базы данных возвещателей из папок """
    print("Генерирую базу данных из папок...")
    files = []
    for p in Subpath:
        if os.path.exists(p):
            files += [p + f for f in os.listdir(p) if isfile(join(p, f))]
    files.sort(key=lambda x: x[0])
    with open(Filename, "w", encoding="utf-8") as datafile:
        for row in files:
            datafile.write(f"{row}\t0\t0\t0\n")
    print("База данных сгенерирована, создан новый файл publishers.ini.")

def scan():
    """ Сканируем базу на предмет новых возвещателей"""
    save_flag = False
    files = []
    for p in Subpath:
        if os.path.exists(p):
            files += [p + f for f in os.listdir(p) if isfile(join(p, f))]
    for f in files:
        for p in Pub:
            if format_title(f) == format_title(p[0]): break
        else:
            Pub.append([f, 0, 0, 0])
            save_flag = True
    if save_flag: save()

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
    print(f"Категория возвещателя:\n[1] {path[0]}\n[2] {path[1]}\n[3] {path[2]}")
    while 1:
        try:
            cat = input("Введите номер варианта или Enter для отмены: ").strip()
            if cat == "":
                cls()
                return None
            else: cat = int(cat)
        except: continue
        if cat==1 or cat==2 or cat==3: return Subpath[cat-1]

def Pub_rename(line1, Pub2):
    # Переименовывание возвещателя.
    # Принимает: line1 = номер строки переименуемого возвещателя; Pub2 - новое имя, в которое он будет переименован
    path = getPath()
    if path is not None:
        oldPubPath = Pub[line1][0]
        newPubPath = f"{path}{Pub2}.pdf"

        for p in Pub:
            if p[0] == newPubPath:
                print("Уже есть возвещатель в этой категории с таким именем!")
                return

        try:
            os.rename(oldPubPath, newPubPath)
            Pub[line1][0] = newPubPath
        except:
            print(f"Не удалось переименовать – скорее всего, открыт PDF-файл.")
            return
        else:
            save()
            print(f"База успешно обновлена, новый файл возвещателя:\n'{newPubPath}'.")

def list():
    """ Вывод полного списка возвещателей """
    count=0
    print(f"Полный список активных возвещателей:")
    max_len = 10 # сначала определяем максимальную ширину колонки
    for p in Pub:
        if len(format_title(p[0])) > max_len:
            max_len = len(format_title(p[0]))
    for p in Pub:
        print(f"{'{:>3})'.format(str(count+1))} {format_title(p[0], max_len=max_len)} {format_report_string(p)}")
        count+=1

def remaining():
    """ Вывод возвещателей, не сдавших отчеты, и их подсчет (команда *) """
    count = 0
    print(f"Не сдали отчеты:")
    for i in range(len(Pub)):
        if Pub[i][1] + Pub[i][2] + Pub[i][3] == 0:
            print(f"{'{:>3})'.format(str(count+1))} {format_title(Pub[i][0])}")
            count+=1

def format_title(name, max_len=5, cut_path=False):
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
    space = ""
    for a in range(max_len - len(string)): space += " "
    return string + space

def format_report_string(p):
    """ Показывает расширение строки с возвещателем, если у него есть отчет (если нет, пусто) """
    string = " →" if p[1]+p[2]+p[3] > 0 else ""
    string += ("%4dч" % p[1]) if p[1] > 0 else "     "
    string += ("%3dи" % p[2]) if p[2] > 0 else "    "
    string += ("%3dк" % p[3]) if p[3] > 0 else "    "
    return string

def cls(command=None):
    """ Очистка терминала и вывод приветственной плашки """
    def __get_pub_reported():
        count = 0
        for p in Pub:
            if p[1] + p[2] + p[3] > 0: count += 1
        return '{:<20}'.format(count)

    if not Devmode:
        os.system('cls' if os.name == 'nt' else 'clear')
        update = str("{:%d.%m.%Y, %H:%M:%S}".format(
            datetime.datetime.strptime(time.ctime((os.path.getmtime(Filename))), "%a %b %d %H:%M:%S %Y")))
        print(f"┌───────────────────────────────────────────────────────────────┐")
        print(f"│ Добро пожаловать в Report Box!                                │")
        print(f"│                                                               │")
        print(f"│ Возвещателей в базе: {'{:<27}'.format(len(Pub))}              │")
        print(f"│ C отчетами: {__get_pub_reported()}                              │")
        print(f"│ Последнее изменение базы: {update}                │")
        print(f"│                                                               │")
        print(f"│ Введите команду и нажмите Enter:                              │")
        print(f"│                                                               │")
        print(f"│ Иван        выбор возвещателя (название файла, даже частично) │")
        print(f"│ Иван 5      ввести отчет возвещателя (только часы)            │")
        print(f"│ Иван 5 1    ввести отчет возвещателя (часы, изучения)         │")
        print(f"│ Иван 5 0 8  ввести отчет возвещателя (часы, изучения, кредит) │")
        print(f"│ +Пётр       создать нового возвещателя                        │")
        print(f"│ :           список всех возвещателей                          │")
        print(f"│ !           кто не сдал отчет                                 │")
        print(f"│ =           статистика собрания                               │")
        print(f"│ *           обнулить все отчеты                               │")
        print(f"│ ?           ничего не понимаю, что здесь происходит?          │")
        print(f"└───────────────────────────────────────────────────────────────┘")
    if command is not None:
        print(f"> {command}")

def update():
    """ Проверяем новую версию и при наличии обновляем программу с GitHub """
    if Devmode: return
    def __update(threadName, delay):
        try:
            for line in requests.get("https://raw.githubusercontent.com/antorix/Report-Box/master/version"):
                newVersion = int(line.decode('utf-8').strip())
        except: pass
        else:
            if newVersion > Version:
                try:
                    response = requests.get("https://github.com/antorix/Report-Box/archive/refs/heads/master.zip")
                    import tempfile
                    import zipfile
                    file = tempfile.TemporaryFile()
                    file.write(response.content)
                    fzip = zipfile.ZipFile(file)
                    fzip.extractall("")
                    file.close()
                    os.remove(f"{Root_path}\\reportbox.py")
                    downloadedFolder = f"{Root_path}\\Report-Box-master"
                    shutil.move(f"{downloadedFolder}\\reportbox.py", Root_path)
                    shutil.rmtree(downloadedFolder)
                except: pass
                else:
                    tkinter.messagebox.showinfo(title="Внимание",
                        message="Найдена новая версия программы, она обновится при следующем запуске!")
            else: pass
    _thread.start_new_thread(__update, ("Thread-Update", 3,))

###

update()
for p in Subpath:
    if not os.path.exists(p): os.makedirs(p) # создаем папки возвещателей, если их нет
if not os.path.exists(Filename): generate()
Pub=load()
scan()
keyboard.add_hotkey('insert', insert_into_PDF) # регистрируем горячую клавишу
cls()

while 1: # главный цикл программы

    command=input("> ").strip()
    scan() # после каждого ввода проверяем папки
    if command=="":
        cls()
        continue
    elif command=="?": # открытие справки в онлайне
        cls(command)
        print("Открываю онлайн-справку...")
        webbrowser.open("https://github.com/antorix/Report-Box/blob/master/README.md#установка-и-начало-работы")
    elif command=="=": # статистика
        cls(command)
        stats()
    elif command=="!": # кто не сдал отчет
        cls(command)
        remaining()
    elif command==":": # показ всех возвещателей
        cls(command)
        list()
    elif command[0]=="+" and len(command)>1: # добавление нового возвещателя
        cls(command)
        Pub_add(command)
    elif command[0]=="*": # обнуление отчетов
        cls(command)
        nullify()

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
        webbrowser.open("")
    elif command=="folder" and not Docmode: # открытие папки файлов в проводнике
        webbrowser.open(Root_path)
    elif command=="share" and not Docmode: # открытие ShareFile
        webbrowser.open("")
    elif command=="picalc" and not Docmode: # подсчет статистики пионеров
        ex.piCalc()

    else:
        cls()
        search(fetch(command)) # обработка команды с указанием возвещателя

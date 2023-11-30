#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
from sys import argv

if "dev" in argv:
    print("Dev mode on")
    Devmode = True
    Root_path = "C:/1/"
    import extras as ex
else:
    Devmode = False
    Root_path = os.path.dirname(os.path.abspath(__file__)) + "/" # корневая папка, в которой лежат подпапки возвещателей и файлы программы

Subpath = []
Subpath.append(Root_path + "Возвещатели/") # подпапки с файлами бланков S-21
Subpath.append(Root_path + "Подсобные пионеры/")
Subpath.append(Root_path + "Общие пионеры/")
Filename = Root_path + "publishers.csv" # файл с собственной базой данных возвещателей
Version = 3
Pub = []
Values = []
Bullet = "•"
Check = "√"
Main_width = len('Активные возвещатели:')+1

import webbrowser
from os.path import isfile, join
import shutil
import time
import csv
import datetime
try:
    import requests
except:
    from subprocess import check_call
    from sys import executable
    check_call([executable, '-m', 'pip', 'install', 'requests'])
    import requests
import _thread
try:
    import keyboard
except:
    from subprocess import check_call
    from sys import executable
    check_call([executable, '-m', 'pip', 'install', 'keyboard'])
    import keyboard

###

def load():
    """ Загрузка базы возвещателей из файла """
    del Pub[:]
    for p in Subpath: # создаем папки возвещателей, если их нет
        if not os.path.exists(p):
            print(f"Папка {p} не найдена, создаю...")
            os.makedirs(p)

    if os.path.exists(Root_path+"publishers.ini"): # конвертация старого формата, удалить в феврале
        print("Найден файл устаревшего формата 'publishers.ini'. Загружаю и удаляю его...")
        with open(Root_path+"publishers.ini", encoding="utf-8", newline='') as file:
            reader = csv.reader(file, delimiter='\t')
            for p in reader:
                Pub.append([p[0], 0, int(p[2]), int(p[1]), int(p[3])])
        save()
        del Pub[:]
        os.remove(Root_path + "publishers.ini")

    if not os.path.exists(Filename): # генерируем файл возвещателей, если его нет
        print("База данных не найдена, генерирую ее из папок...")
        files = []
        for p in Subpath:
            if os.path.exists(p):
                files += [p + f for f in os.listdir(p) if isfile(join(p, f))]
        files.sort(key=lambda x: x[0])
        with open(Filename, "w", encoding="utf-8", newline='') as file:
            writer = csv.writer(file, delimiter=',')
            for row in files: writer.writerow([row, 0, 0, 0, 0])
        print(f"База сгенерирована, создан новый файл '{Filename}'.")

    with open(Filename, encoding="utf-8", newline='') as csvfile: # открываем файл
        reader = csv.reader(csvfile, delimiter=',')
        for p in reader:
            Pub.append([revert(p[0]), int(p[1]), int(p[2]), int(p[3]), int(p[4])])
        print("Файл данных успешно загружен.")

    save()

def save():
    """ Выгрузка базы данных в файл """
    Pub.sort(key=lambda x: x[0])
    with open(Filename, "w", encoding="utf-8", newline='') as file:
        writer = csv.writer(file, delimiter=',')
        for p in Pub:
            if p[1] is None: p[1] = 0
            if p[2] is None: p[2] = 0
            if p[3] is None: p[3] = 0
            if p[4] is None: p[4] = 0
            writer.writerow([p[0], p[1], p[2], p[3], p[4]])

def fetch(command):
    """ Интерпретация строки отчета - перевод простой строки в список из имени возвещателя и 4 параметров """
    if not Devmode: print(f"> {command}\nВот что нашлось по вашему запросу:")
    after_equal = None
    if "=" in command and Devmode: # если есть знак = в конце, добавим его как отдельное значение к возврату из fetch
        after_equal = command[command.index("=") + 1:].strip()
        command = command[ : command.index("=")].strip()
    for delimeter in [" ", "\t"]:
        try:
            if delimeter in command:
                name=command[: command.index(delimeter)].strip()
                stats=command[command.index(delimeter)+1 :].strip()
                tPos=[0]
                for y in range(len(stats)):  # выяснение положений табуляций
                    if stats[y]==delimeter: tPos.append(y)

                if len(tPos)    == 1:   # введено служение
                    stat1 = int(stats[tPos[0]: ].strip())           # служение
                    stat2 = stat3 = stat4 = None

                elif len(tPos)  == 2:   # введены служение, изучения
                    stat1 = int(stats[tPos[0]: tPos[1]].strip())    # служение
                    stat2 = int(stats[tPos[1]:].strip())            # изучения
                    stat3 = stat4 = None

                elif len(tPos)  == 3:   # введены служение, изучения, часы
                    stat1 = int(stats[tPos[0]: tPos[1]].strip())    # служение
                    stat2 = int(stats[tPos[1]: tPos[2]].strip())    # изучения
                    stat3 = int(stats[tPos[2]: ].strip())           # часы
                    stat4 = None

                elif len(tPos)  == 4:   # введены служение, изучения, часы, кредит
                    stat1 = int(stats[tPos[0]: tPos[1]].strip())    # служение
                    stat2 = int(stats[tPos[1]: tPos[2]].strip())    # изучения
                    stat3 = int(stats[tPos[2]: tPos[3]].strip())    # часы
                    stat4 = int(stats[tPos[3]: ].strip())           # кредит
                result = [name, stat1, stat2, stat3, stat4]
                break
            else: result = [command.strip(), None, None, None, None]
        except:
            print("Значения отчета должны содержать только цифры и отличаться от ноля.\n"+\
                  "Попробуйте еще раз.")
            result = None
            break
    if after_equal is not None:
        result.append(after_equal)
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
            if not os.path.exists(f"{Root_path}/Архив"):
                os.makedirs(f"{Root_path}/Архив")
            shutil.copyfile(Filename, f"{Root_path}/Архив/publishers-{savedTime}.csv")
        for p in Pub: p[1] = p[2] = p[3] = p[4]= 0
        save()
        print("Отчеты успешно обнулены.")
    else: cls()

def Pub_add(string):
    """ Создание нового возвещателя """
    S21 = Root_path + "S-21_U.pdf"
    if not "/" in string:
        name = string[1:].strip()
        type = getPath()
    else:
        try:
            name = string[ 1 : string.index("/") ].strip()
            type = string[string.index("/") + 1 : ].strip()
            if type == "В": type = Subpath[0]
            elif type == "П": type = Subpath[1]
            elif type == "О": type = Subpath[2]
            else: 5 / 0
        except:
            print("Что-то не сработало.")
            return
    newFileName = f"{type}{name}.pdf"
    for p in Pub:
        if p[0] == newFileName:
            print("Уже есть возвещатель в этой категории с таким именем!")
            return
    if type is not None:
        while 1:
            if os.path.exists(S21):
                shutil.copyfile(S21, newFileName)
                Pub.append([newFileName, 0, 0, 0, 0])
                print(f"Успешно создано: {format_title(newFileName)}.")
                save()
                break
            else:
                try:
                    import tkinter.filedialog
                    print("Не найден бланк S-21! Укажите местоположение этого\n" + \
                          "бланка в формате PDF...")
                    S21 = tkinter.filedialog.askopenfilename()
                    if S21 != "":
                        shutil.copyfile(S21, f"{Root_path}/S-21_U.pdf") # копируем пустой бланк S-21 в папку программы на будущее
                    else:
                        print("Для создания возвещателей программе нужен бланк S-21 в\n"+\
                              "формате PDF. Вы знаете, где его найти.")
                        break
                except:
                    print("Не найден бланк S-21! Положите файл 'S-21_U.pdf' в папку\n" + \
                          "программы и попробуйте снова. Вы знаете, где его найти.")
                    break

def search(myinput=None, line=None, process=True):
    """ Поиск возвещателя и ввод его данных. Возвращает индекс строки в массиве Pub с этим возвещателем """
    found = []
    if myinput != None: name=myinput[0].strip()
    max_len = 0
    if line is None:
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
                    elif not Devmode and (value=="groups" or value=="!"):
                        webbrowser.open(ex.groups_link)
                    else:
                        line=found[int(value)-1]
                        print(f"Ваш выбор:\n{Bullet} {format_title(Pub[line][0])} {format_report_string(Pub[line])}")
                        break
                except:
                    continue

    if found == [] and line is not None:
        print(f"{Bullet} {format_title(Pub[line][0].strip(), max_len)} {format_report_string(Pub[line])}")

    if line is not None and not report_entered(myinput):
        # открытие возвещателя без ввода статистики, с меню для выбора

        print("Меню действий:\n"+\
              "[1] Открыть PDF-файл\n[2] Ввести отчет\n[3] Обнулить отчет\n[4] Переименовать/переместить\n[5] Удалить")
        if Devmode: print("[=Вася] Переименование в той же категории\n[=Вася/В] Переименование со сменой категории (В: возвещатели; П: подсобные; О: общие)")
        while 1:
            choice = input("Введите номер действия или Enter для отмены: ").strip()
            if choice == "":
                cls()
                return None
            elif Devmode and "=" in choice: # ускоренный синтаксис переименования возвещателя
                try:
                    if "/" in choice:
                        new_name = choice[ choice.index("=")+1 : choice.index("/")].strip()
                        path = choice[ choice.index("/")+1 :].strip()
                        if   path == "В": path = Subpath[0]
                        elif path == "П": path = Subpath[1]
                        elif path == "О": path = Subpath[2]
                        else: 5/0
                    else:
                        new_name = choice[choice.index("=") + 1: ].strip()
                        if new_name == "": 5/0
                        if "Возвещат" in Pub[line][0]: path = Subpath[0]
                        elif "Подсоб" in Pub[line][0]: path = Subpath[1]
                        elif "Общ" in Pub[line][0]: path = Subpath[2]
                    Pub_rename(line1=line, Pub2=new_name, path=path)
                    break
                except: print("Что-то не сработало.")
            else:
                try: choice = int(choice)
                except: continue
            if choice == 1:
                print("Открываю PDF-файл...")
                if os.path.exists(Pub[line][0]): webbrowser.open(Pub[line][0])
                else: print("PDF-файл возвещателя отсутствует в базе!")
                break
            elif choice == 2:
                print("Внимание! Отчеты вводятся так, как показано на подсказке по\n"+\
                      "командам выше. Если вам это непонятно, введите «?» и почитайте\n"+\
                      "онлайн-справку.")
                break
            elif choice == 3:
                Pub[line][1] = Pub[line][2] = Pub[line][3] = Pub[line][4] = 0
                print("Отчет возвещателя обнулен.")
                save()
                break
            elif choice == 4:
                while 1:
                    new_name = input("Переименование возвещателя\nВведите новое название файла или Enter для отмены: ").strip()
                    if new_name != "":
                        Pub_rename(line, new_name)
                        break
                    else:
                        cls()
                        break
                break
            elif choice == 5:
                result = Pub_delete(line)
                if result == False: continue # Pub_delete возвращает False, если нельзя удалить файл, потому что он открыт
                else: break

    elif line is not None and process: # открытие возвещателя с вводом статистики
        print(f"Режим ввода отчета. Открываю PDF-файл...")
        Pub[line] = [Pub[line][0], myinput[1], myinput[2], myinput[3], myinput[4]]
        save()
        webbrowser.open(Pub[line][0])
        del Values[:]
        for val in Pub[line]:
            Values.append(val if val is not None else 0)
        print("Поставьте галочку в поле «Участвовал в служении» отчетного\nмесяца и нажмите на кнопку «Insert».")
        if len(myinput) > 5: pioneer_add(myinput) # получен дополнительный параметр, введенный после знака =
    return line

def pioneer_add(myinput):
    """ Добавляет статистику в файл пионеров с произвольной фамилией """
    pioneers = []
    pioneer = myinput[5].strip()
    pion_file = Root_path + "pioneers.csv"
    if not os.path.exists(pion_file):
        with open(pion_file, "w", encoding="utf-8") as f: pass
    with open(pion_file, encoding="utf-8", newline='') as file:
        reader = csv.reader(file, delimiter=',')
        for row in reader: pioneers.append(row)
    for p in pioneers:
        if p[0] == pioneer:
            print("Такая запись уже есть!")
            break
    else:
        if myinput[3] is None: myinput[3] = 0
        if myinput[4] is None: myinput[4] = 0
        if myinput[3] + myinput[4] <= 55:
            hours = myinput[3] + myinput[4]
        else:
            hours = myinput[3]
            plus = myinput[4]
            if hours + plus >= 55:
                plus = plus - (-(55 - hours) + plus)
                hours += plus
        pioneers.append([pioneer, hours])
        pioneers.sort(key=lambda x: x[0])
        with open(pion_file, "w", encoding="utf-8", newline='') as file:
            writer = csv.writer(file, delimiter=',')
            for p in pioneers:
                writer.writerow(p)
        print(f"Добавлено в файл пионеров: {[pioneer, hours]}")

def scan():
    """ Сканируем базу на предмет новых возвещателей"""
    save_flag = False
    files = []
    for p in Subpath:
        if os.path.exists(p):
            files += [p + f for f in os.listdir(p) if isfile(join(p, f))]
    for f in files:
        for p in Pub:
            if format_title(f) == format_title(p[0]):
                break
        else:
            Pub.append([revert(f), 0, 0, 0, 0])
            save_flag = True
    if save_flag: save()

def revert(p):
    """ Переворачиваем косые черты с \\ на / """
    string = ""
    for char in p:
        if char == "\\": string += "/"
        else: string += char
    return string

def stats(mode=1):
    """ Подсчет и вывод статистики """
    def printStat(): # вывод статистики
        print(  f"{'{:{align}{width}}'.format('Активные возвещатели:'.upper(), align='<', width=Main_width)}{len(Pub)}\n")
        print(   '{:{align}{width}}'.format("Возвещатели".upper(), align='^', width=Main_width))
        print(f"{'{:{align}{width}}'.format('Количество отчетов:', align='<', width=Main_width)}{countP}")
        print(f"{'{:{align}{width}}'.format('Изучения Библии:', align='<', width=Main_width)}{countP_St}\n")
        print('{:{align}{width}}'.format("Подсобные пионеры".upper(), align='^', width=Main_width))
        print(f"{'{:{align}{width}}'.format('Количество отчетов:', align='<', width=Main_width)}{countAP}")
        print(f"{'{:{align}{width}}'.format('Часы:', align='<', width=Main_width)}{countAP_Hours}")
        print(f"{'{:{align}{width}}'.format('Изучения Библии:', align='<', width=Main_width)}{countAP_St}\n")
        print('{:{align}{width}}'.format("Общие пионеры".upper(), align='^', width=Main_width))
        print(f"{'{:{align}{width}}'.format('Количество отчетов:', align='<', width=Main_width)}{countRP}")
        print(f"{'{:{align}{width}}'.format('Часы:', align='<', width=Main_width)}{countRP_Hours}")
        print(f"{'{:{align}{width}}'.format('Изучения Библии:', align='<', width=Main_width)}{countRP_St}")

    countP=countP_St=countAP=countAP_Hours=countAP_St=countRP=countRP_Hours=countRP_St=0
    for p in Pub:
        if "Возвещатели" in p[0] and p[1]+p[2]+p[3]+p[4] > 0:
            countP += 1
            countP_St += p[2]
        elif "Подсобные" in p[0] and p[1]+p[2]+p[3]+p[4] > 0:
            countAP += 1
            countAP_St += p[2]
            countAP_Hours += p[3]
        elif "Общие" in p[0] and p[1]+p[2]+p[3]+p[4] > 0:
            countRP += 1
            countRP_St += p[2]
            countRP_Hours += p[3]
    if mode==1: # показ статистики
        printStat()
    elif mode==2: # ввод статистики на сайте
        printStat()
        webbrowser.open(ex.J_link)
        del Values[:]
        Values.append([ str(len(Pub)),
                        str(int(countP)),                               str(int(countP_St)),
                        str(int(countAP)),  str(int(countAP_Hours)),    str(int(countAP_St)),
                        str(int(countRP)),  str(int(countRP_Hours)),    str(int(countRP_St))
                        ]
        )

def auto_insert():
    """ Вставка данных """
    if len(Values) > 1:
        if "Подсобные" in Values[0] or Values[2] + Values[3] + Values[4] > 0:
            keyboard.press_and_release("\t") # сместились на изучение Библии
            if Values[2] != 0:
                keyboard.write(str(Values[2])) # ввели изучение Библии
            if "Подсобные" in Values[0] or Values[3] + Values[4] > 0:
                keyboard.press_and_release("\t") # сместились на ПП
                if "Подсобные" in Values[0]:
                    keyboard.press_and_release(" ") # поставили галочку на подсобного
                if Values[3] + Values[4] > 0:
                    keyboard.press_and_release("\t") # сместились на часы
                if Values[3] != 0:
                    keyboard.write(str(Values[3] if not "Возвещатели" in Values[0] else "")) # ввели часы
                if Values[4] != 0:
                    keyboard.press_and_release("\t")  # сместились на примечание
                    keyboard.write(f"кредит {Values[4]}")  # ввели кредит
        keyboard.press_and_release("ctrl+s")
        if 0:#Values[4] == 0:
            time.sleep(.1)
            keyboard.press_and_release("ctrl+w")
        del Values[:]
    else:
        try:
            keyboard.write(Values[0][0])
            del Values[0][0]
        except: pass

def getPath():
    """ Спрашивает категорию возвещателя и возвращает путь к соответствующей папке """
    path = []
    for p in Subpath:
        last = len(p)-1
        path.append(p[ len(Root_path) : last])
    print(f"Категория возвещателя:\n[1] {path[0]}\n[2] {path[1]}\n[3] {path[2]}")
    while 1:
        try:
            cat = input("Введите номер категории или Enter для отмены: ").strip()
            if cat == "":
                cls()
                return None
            else: cat = int(cat)
        except: continue
        if cat==1 or cat==2 or cat==3: return Subpath[cat-1]

def Pub_rename(line1, Pub2, path=None):
    # Переименовывание возвещателя.
    # Принимает: line1 = номер строки переименуемого возвещателя; Pub2 - новое имя, в которое он будет переименован
    if path is None: path = getPath()
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
            print(f"Не удалось переименовать, ошибка записи в файл.\nВозможно, открыт PDF-файл возвещателя.")
            return
        else:
            save()
            print(f"Успешно обновлено: {format_title(newPubPath)}.")

def list():
    """ Вывод полного списка возвещателей """
    print(f"АКТИВНЫЕ ВОЗВЕЩАТЕЛИ: {len(Pub)}")
    max_len = 10 # сначала определяем максимальную ширину колонки
    for p in Pub:
        if len(format_title(p[0])) > max_len:
            max_len = len(format_title(p[0]))
    for p in Pub:
        print(f"{format_title(p[0], max_len=max_len)} {format_report_string(p)}")

def remaining():
    """ Вывод возвещателей, не сдавших отчеты, и их подсчет (команда *) """
    count = 0
    found = []
    for i in range(len(Pub)):
        if Pub[i][1] + Pub[i][2] + Pub[i][3] + Pub[i][4] == 0:
            found.append(i)
            count+=1
    print(f"НЕ СДАЛИ ОТЧЕТЫ: {count}")
    for f in found: print(format_title(Pub[f][0]))

def export_groups():
    """ Экспортирует csv-файл с возвещателями, отсортированными по группам """
    def __numberize(name):
        for char in name:
            try: return int(char)
            except: continue
    Pub.sort(key=lambda x: __numberize(format_title(x[0])))
    count = 1
    group = 0
    with open(Root_path + "groups.csv", 'w', encoding="utf-8", newline='') as file:
        writer = csv.writer(file, delimiter='\t')
        for p in Pub:
            if str(group) not in format_title(p[0]):
                count = 1
                group += 1
                writer.writerow("")
                writer.writerow(["", f"\u00A0ГРУППА № {group}", "Служил(а)", "Изучения", "Часы\u00A0\u00A0\u00A0", "Кредит"])
                writer.writerow("")
            else:
                count += 1
            writer.writerow([
                f"{count}",
                f"\u00A0{format_title(p[0])}",
                "☑" if p[1] != 0 else '',
                str(p[2]) if p[2] != 0 else '',
                str(p[3]) if p[3] != 0 and not 'Возвещатели' in p[0] else '',
                str(p[4]) if p[4] != 0 else ''
            ])
    print("Успешно экспортировано, открываю файл...")
    webbrowser.open(Root_path + "groups.csv")
    Pub.sort(key=lambda x: x[0])

def pioneers_show():
    """ Показ всех пионеров """
    found = []
    for i in range(len(Pub)):
        if "Общие" in Pub[i][0]: found.append(i)
    for i in range(len(found)):
        print(f"{i+1} {Bullet} {format_title(Pub[found[i]][0], cut_path=True)} {format_report_string(Pub[found[i]])}")
    while 1:
        try:
            choice = input("Введите номер пионера или Enter для выхода: ").strip()
            if choice == "": break
            choice = int(choice)
            line=found[int(choice)-1] # номер строки этого пионера в общем списке возвещателей
            name = input(f"С какой фамилией связать статистику «{format_title(Pub[line][0], cut_path=True).strip()}»? ")
            if name != "":
                myinput = [Pub[line][0], Pub[line][1], Pub[line][2], Pub[line][3], Pub[line][4], name]
                pioneer_add(myinput)
        except:
            print("Не сработало.")

def format_title(name, max_len=5, cut_path=False):
    """ Форматирование и сокращение пути к файлу с именем возвещателя """
    cut = len(Root_path)
    name = name[cut:]
    string = ""
    if cut_path:
        try: string = name[ name.index("/")+1 : name.index(".pdf")]
        except: pass
    else:
        for char, i in zip(name, range(len(name))):
            if name[i:i+4] == ".pdf": break
            elif char != "/":
                string += char
            else: string += ": "
    space = ""
    for a in range(max_len - len(string)): space += " "
    return string + space

def report_entered(p):
    """ Проверяет поля отчета и возвращает True, если имеется любое значение не None """
    if p is None:
        result = False
    elif p[1] is not None or p[2] is not None or p[3] is not None or p[3] is not None:
        result = True
    else:
        result = False
    return result

def format_report_string(p):
    """ Показывает расширение строки с возвещателем, если у него есть отчет (если нет, пусто) """
    string = " →"# if report_entered(p) else ""
    string += "%5s" % (Check if p[1] is not None and p[1] != 0 else "")
    string += "%5s" % (str(p[2])+"и" if p[2] is not None and p[2] != 0 else "")
    string += "%5s" % (str(p[3])+"ч" if p[3] is not None and p[3] != 0 and not "Возвещат" in p[0] else "")
    string += "%5s" % (str(p[4])+"к" if p[4] is not None and p[4] != 0 else "")
    return string

def cls(command=None, info_force=False):
    """ Очистка терминала и вывод приветственной плашки """
    def __get_pub_reported():
        count = 0
        for p in Pub:
            if p[1] + p[2] + p[3] > 0: count += 1
        return '{:<30}'.format(count)

    if not Devmode or info_force:
        os.system('cls' if os.name == 'nt' else 'clear')
        update = str("{:%d.%m.%Y, %H:%M:%S}       ".format(
            datetime.datetime.strptime(time.ctime((os.path.getmtime(Filename))), "%a %b %d %H:%M:%S %Y")))
        print(f"┌───────────────────────────────────────────────────────────────────────────┐")
        print(f"│ Добро пожаловать в Report Box!                                     (v. {Version}) │")
        print(f"│                                                                           │")
        print(f"│ Возвещателей в базе: {'{:<27}'.format(len(Pub))}                          │")
        print(f"│ C отчетами: {__get_pub_reported()}                                │")
        print(f"│ Последнее изменение базы: {update}                     │")
        print(f"│                                                                           │")
        print(f"│ Введите команду и нажмите Enter:                                          │")
        print(f"│                                                                           │")
        print(f"│ Иван            выбор возвещателя (название файла, даже частично)         │")
        print(f"│ Иван 1          ввести отчет (служил)                                     │")
        print(f"│ Иван 1 2        ввести отчет (служил, 2 изучения)                         │")
        print(f"│ Иван 1 0 50     ввести отчет (служил, 0 изучений, 50 часов)               │")
        print(f"│ Иван 1 0 50 10  ввести отчет (служил, 0 изучений, 50 часов, 10 ч. кредит) │")
        print(f"│ +Пётр           создать нового возвещателя                                │")
        print(f"│ :               список всех возвещателей                                  │")
        print(f"│ !               кто не сдал отчет                                         │")
        print(f"│ =               статистика собрания                                       │")
        print(f"│ *               обнулить все отчеты                                       │")
        print(f"│ ?               ничего не понимаю, что здесь происходит?                  │")
        print(f"└───────────────────────────────────────────────────────────────────────────┘")
    if not Devmode and command is not None:
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
                    os.remove(f"{Root_path}/reportbox.py")
                    downloadedFolder = f"{Root_path}/Report-Box-master"
                    shutil.move(f"{downloadedFolder}/reportbox.py", Root_path)
                    shutil.rmtree(downloadedFolder)
                except: pass
                else:
                    message = "Найдена новая версия программы, она обновится при следующем запуске!"
                    try:
                        import tkinter.messagebox
                        tkinter.messagebox.showinfo(title="Внимание", message=message)
                    except:
                        print(message)
            else: pass
    _thread.start_new_thread(__update, ("Thread-Update", 3,))

###

update()
load()
scan()
cls()

try: keyboard.add_hotkey('insert', auto_insert) # регистрируем горячую клавишу
except:
    if os.name != "nt": print("Не удалось зарегистрировать горячую клавишу Insert для быстрой\n"+\
                              "вставки данных в PDF-файлы! На Unix-системах для этого необходимо\n"+\
                              "работать в режиме суперпользователя.")

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
    elif command == "list":
        for p in Pub: print(p)
    elif command=="file":
        webbrowser.open(Filename)
    elif command=="folder":
        webbrowser.open(Root_path)
    elif command == "export":
        export_groups()

    elif Devmode and command == "import": # dev mode
        ex.excel(Pub)
    elif Devmode and command == "groups":
        webbrowser.open(ex.Groups_link)
    elif Devmode and command == "=jw":
        stats(mode=2)
    elif Devmode and command == "pi":
        pioneers_show()
    elif Devmode and command == "zoom":
        ex.Zoom_count(Root_path)
    elif Devmode and command == "wipe":
        if os.path.exists(Filename):
            os.remove(Filename)
            load()
            save()

    elif Devmode and command == "help": # подсказки по дополнительным командам
        cls(info_force=True)
        print("Дополнительные команды:")
        print("list     полный список базы данных")
        print("file     открытие файла данных")
        print("folder   открытие папки с программой")
        print("export   экспорт групп в CSV")
        print("groups   группы онлайн")
        print("=jw      ввод статистики на сайте")
        print("pi       вывод пионеров")
        print("import   импорт статистики из файла")
        print("zoom     обсчет посещаемости в Zoom")
        print("wipe     пересоздание файла данных")

    else:
        cls()
        result = fetch(command)
        if result is not None: search(result) # обработка команды с указанием возвещателя

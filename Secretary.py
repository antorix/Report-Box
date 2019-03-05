mypath1="D:\Мои документы\Отчеты\Возвещатели" # папки бланков S-21, пропишите здесь полные пути
mypath2="D:\Мои документы\Отчеты\Возвещатели некрещеные"
mypath3="D:\Мои документы\Отчеты\Пионеры"
mypath4="D:\Мои документы\Отчеты\Пионеры подсобные"
PDFViwer="C:\Program Files\Tracker Software\PDF Viewer\PDFXCview.exe" # полный путь к PDF-читалке
Filename="publishers.txt" # файл с возвещателями
Delimeter="\t" # разделитель данных при вводе

import keyboard
import webbrowser
from os import path, listdir
from os.path import isfile, join
import xlrd
import glob
RowCount=0
Values=["", "", "", "", ""]

def load():
    """Загрузка базы возвещателей из файла"""
    global Filename
    publishers=[]
    if path.exists(Filename)==0:
        print("Файл %s не найден." % Filename)
        return None
    else:       
        with open(Filename, "r", encoding="utf-8") as f:
            lines=[line for line in f]            
    for i in range(len(lines)):
        publishers.append(["", "", "", "", "", ""])
        try:
            publishers[i][0] = lines[i][0 : lines[i].index("\t")] # присвоение имени
        except:
            publishers[i][0] = lines[i] # присвоение имени, если нет статистики
        tPos=[]
        for y in range(len(lines[i])):  # выяснение положений табуляций
            if lines[i][y]=="\t" or lines[i][y]=="\n":
                tPos.append(y)
        for length in range(6):
            if len(tPos)==length:
                for m in range(6-length):
                    tPos.append(0)
        publishers[i][1] = lines[i][tPos[0] : tPos[1]].strip() # публикации
        publishers[i][2] = lines[i][tPos[1] : tPos[2]].strip() # видео
        publishers[i][3] = lines[i][tPos[2] : tPos[3]].strip() # часы
        publishers[i][4] = lines[i][tPos[3] : tPos[4]].strip() # ПП
        publishers[i][5] = lines[i][tPos[4] : tPos[5]].strip() # изучения    
    print("Возвещатели успешно загружены.")
    return publishers
    
def save(list):
    """Выгрузка базы данных в файл"""
    global Filename
    try:
        with open(Filename, "w", encoding="utf-8") as datafile:
            for row in list:
                datafile.write(row[0].strip()+"\t"+row[1].strip()+"\t"+row[2].strip()+"\t"+row[3].strip()+"\t"+row[4].strip()+"\t"+row[5].strip()+"\t\n")
    except:
        print("Ошибка сохранения!")

def fetch(command):
    """Интерпретация строки отчета"""
    global Delimeter
    if Delimeter in command:
        name=command[: command.index(Delimeter)]
        stats=command[command.index(Delimeter)+1 :]
        tPos=[0]
        for y in range(len(stats)):  # выяснение положений табуляций
            if stats[y]==Delimeter:
                tPos.append(y)
        stat1 = stats[tPos[0] : tPos[1]].strip() # публикации
        stat2 = stats[tPos[1] : tPos[2]].strip() # видео
        stat3 = stats[tPos[2] : tPos[3]].strip() # часы
        stat4 = stats[tPos[3] : tPos[4]].strip() # ПП
        stat5 = stats[tPos[4] : ]                # изучения    
        return [name, str(int(stat1)), str(int(stat2)), str(float(stat3)), str(int(stat4)), str(int(stat5))]
    else:
        return [command.strip(), None, None, None, None, None]    
    
def search(input, pub):
    """Поиск возвещателя и ввод его данных"""
    global Values
    global RowCount
    for i in range(6):
        if input[i]=="":
            input[i]="0.0"
    if input[0]=="0.0":
        return
    else:
        print("Ищу %s…" % input[0])
    count=line=0
    for i in range(len(pub)):
        if input[0].lower() in pub[i][0].lower():
            print("Найдено: " + extractname(pub[i][0]))
            line=i
            count+=1
    if count==0:
        print("Ничего не найдено.")
    elif count==1:
        if input[1]==None:
            webbrowser.open(pub[line][0])
            return
        else:
            for i in range(6):
                if ".0" in input[i]:
                    input[i]="%d" % float(input[i])
            pub[line]=[pub[line][0], input[1], input[2], input[3], input[4], input[5]]
            save(pub)
            print("База обновлена, открываю файл.\nЖду окончания ввода...")
            webbrowser.open(pub[line][0])
            Values=[input[1].strip(), input[2].strip(), input[3].strip(), input[4].strip(), input[5].strip()]
            RowCount=0
            keyboard.wait('ctrl+w')
    else:
        print("Найдено %s, конкретизируйте." % count)   

def recreate():
    """Генерация новой базы данных возвещателей из папок (команда СБРОС)"""
    global mypath1, mypath2, mypath3, mypath4, Filename
    try:
        files = [mypath1+f for f in listdir(mypath1) if isfile(join(mypath1, f))]        
        files = files + [mypath2+f for f in listdir(mypath2) if isfile(join(mypath2, f))]        
        files = files + [mypath3+f for f in listdir(mypath3) if isfile(join(mypath3, f))]
        files = files + [mypath4+f for f in listdir(mypath4) if isfile(join(mypath4, f))]
        print("Сброс файла данных и перезагрузка из папок…")
    except:
        print("Не найдены нужные папки, проверьте настройки.")
        return
    files.sort(key=lambda x: x[0])    
    with open(Filename, "w", encoding="utf-8") as datafile:
        for row in files:
            datafile.write(row+"\n")

def stats(pub):
    """Подсчет и вывод статистики для ввода на jw.org (команда =)"""
    countP=countP_Hours=countP_Ret=countP_St=countAP=countAP_Hours=countAP_Ret=countAP_St=countRP=countRP_Hours=countRP_Ret=countRP_St=0.0
    for row in pub:
        if("подсобные\\" in row[0] and row[3]!=""):
            countAP+=1
            countAP_Hours+=float(row[3])
            countAP_Ret+=int(row[4])
            countAP_St+=int(row[5])
        elif("\\Пионеры\\" in row[0] and row[3]!=""):
            countRP+=1
            countRP_Hours+=float(row[3])
            countRP_Ret+=int(row[4])
            countRP_St+=int(row[5])
        elif("\\Возвещатели\\" in row[0] and row[3]!=""):
            countP+=1
            countP_Hours+=float(row[3])
            countP_Ret+=int(row[4])
            countP_St+=int(row[5])
    print("Возвещатели:\nЧисло сдавших отчет: %d\nЧасы: %.1f\nПовторные посещения: %d\nИзучения Библии: %d\n" % (countP, countP_Hours, countP_Ret, countP_St))
    print("Подсобные пионеры:\nЧисло сдавших отчет: %d\nЧасы: %.1f\nПовторные посещения: %d\nИзучения Библии: %d\n" % (countAP, countAP_Hours, countAP_Ret, countAP_St))
    print("Пионеры:\nЧисло сдавших отчет: %d\nЧасы: %.1f\nПовторные посещения: %d\nИзучения Библии: %d\n" % (countRP, countRP_Hours, countRP_Ret, countRP_St))
    print("Активные возвещатели: %d" % len(pub))

def reprow():
    """Подстановка данных в буфер обмена"""
    global RowCount
    global Values
    if RowCount<5:
        if Values[RowCount]!="0":
            keyboard.write(Values[RowCount])
        else:
            keyboard.write("")
        RowCount+=1
    else:
        return

def excel(pub):
    """Импорт отчетов из таблицы Excel (команда +)"""
    try:
        filenames = glob.glob('*.xlsx')
        book = xlrd.open_workbook(filenames[0])
        sheet = book.sheet_by_index(0)
    except:
        print("Xlsx-файл не найден или защищен от чтения.")
        return
    line=[]
    for row in range(sheet.nrows):
        line.append([str(format(sheet.cell(row,col).value)) for col in range(6)])
    for r in range(len(line)):
        search(line[r], pub)

def remaining(pub):
    """Вывод возвещателей, не сдавших отчеты, и их подсчет (команда *)"""
    count=0
    for i in range(len(pub)):
        if pub[i][3]=="":
            print(extractname(pub[i][0]))
            count+=1
    print("Не сдали отчеты: %d" % count)

def extractname(string):
    #"""Сокращение имени файла возвещателя до более удобного"""
    #return string[string.index("возвещателей\\")+13 : string.index("S-21.pdf")].strip()
    return string

webbrowser.open(PDFViwer)
pub=load()
if pub==None:
    print("Создаю файл из папок…")
    recreate()
    pub=load()
keyboard.add_hotkey('right', reprow)

# Главный цикл

while 1:    
    command=input("> ")
    if command=="?": # открытие справки в онлайне
        webbrowser.open("https://github.com/antorix/secretary/blob/master/readme.md")
    elif command=="СБРОС":
        recreate()
        pub=load()        
    elif command=="=":
        stats(pub)
    elif command=="\\":
        webbrowser.open(Filename)
    elif command=="+":
        excel(pub)
    elif command=="*":
        remaining(pub)
    else:
        try:
            result1=fetch(command)
            result2=search(result1, pub)
        except:
            print("Команда непонятна, попробуйте еще (введите \"?\" для показа справки).")

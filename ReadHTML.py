# from operator import index
import os
import sys
from math import ceil as math_ceil
import threading
import traceback
import urllib.request
import win32com.client
from time import sleep
from okno_ui import Ui_Form
from PyQt5 import QtWidgets, QtCore
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
import SvodTable

# from rich import print

app = QtWidgets.QApplication(sys.argv)
Form = QtWidgets.QWidget()
Form.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)  # поверх других окон
ui = Ui_Form()
ui.setupUi(Form)
Form.show()

def exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol):
    '''Отправляем данные в диапозон ячеек'''
    sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol)).Formula = data

# def GO(FullName, filename, zond, dayList, datatimes, totals, text0, data, svobod, procDostyp, dostypMemory, allMemor):
def GO(FullName, filename):
    global ZGDDict
    zond = []
    dayList = []
    datatimes = []
    totals = []
    svobod = []
    procDostyp = []
    dostypMemory = []
    allMemor = []
    

    '''url, на котором находится список преподователей'''
    FullName = 'file:///' + FullName
    link = urllib.request.urlopen(FullName)
    lines = []
    for line in link.readlines():
        '''взяли все строки с сотрудниками'''
        '''Каждый сотрудник начинается с тегов <li><a href .....'''
        # if line.find(b'<nobr>') != -1 or (line.find(b'<td colspan="6">') != -1 and b'<td colspan="6"><b></b></td>' not in line):
        # if line.find(b'<nobr>') != -1 or line.find(b'rosneft.ru') != -1:
        if b'<nobr>' in line or b'rosneft.ru' in line:            
            lines.append(line)
    # print(lines)
    link.close()

    '''Переводим bytes в str'''
    for i in range(len(lines)):
        lines[i] = lines[i].decode('utf-8')

    '''# svobod, procDostyp, dostypMemory, allMemor, timeProstoya'''
    def poisk(pered, posle):
        lenpered = len(pered)
        index1 = lines[i].find(pered) + lenpered
        index2 = lines[i].find(posle, index1)
        text = lines[i][index1 : index2]
        if '&nbsp;' in text:
            text = ''
        if '&lt;1 %' in text:
            text = '<1%'
        if '&minus;' in text:
            text = text.replace('&minus;', '-')

        return text
    
    tipdata = ''
    if 'CPU' in filename.upper():
        tipdata = 'CPU'
    if 'HDD' in filename.upper():
        tipdata = 'HDD'
    if 'MEMORY' in filename.upper():
        tipdata = 'MEMORY'
    # tipdata.append(tip)
    
    text0 = ''
    for i in range(len(lines)):

        '''Зонд, группа, устройство'''
        if '.rosneft.ru' in lines[i]:
            # pered = '<td colspan="6">'
            pered = '>'
            # posle = '</td>'
            posle = '.rosneft.ru'
            text0 = poisk(pered, posle)
            # text0 = text0.split('.rosneft.ru')[0]

        '''Дата и время_dayList_datatimes'''
        if '<nobr>' in lines[i]:
            text1 = ''
            text2 = ''
            text3 = ''
            text4 = ''
            text5 = ''
            text6 = ''
            text7 = ''

            pered = '<nobr>'
            posle = '</nobr>'
            textx = poisk(pered, posle)
            textx = textx.split(' ', maxsplit=1)
            text1 = textx[0]
            text2 = textx[-1]

            '''Всего CPU_totals'''
            if tipdata == 'CPU':
                if 'всего">' in lines[i]:
                    pered = 'всего">'
                    posle = '</td>'
                    text3 = poisk(pered, posle)
                
                if 'Не найдено' in lines[i]:
                    text3 = ''
            
            if tipdata == 'HDD':
                '''HDD Свободное пространство_svobod'''
                if 'col-свободное-пространство">' in lines[i]:
                    pered = 'col-свободное-пространство">'
                    posle = '</td>'
                    text4 = poisk(pered, posle)
          
            '''Memoiry Процент доступной памяти_procDostyp'''
            if tipdata == 'MEMORY':
                if 'процент-доступной-памяти">' in lines[i]:
                    pered = 'процент-доступной-памяти">'
                    posle = '</td>'
                    text5 = poisk(pered, posle)
                
                    '''dostypMemory'''
                    if 'доступная-память">' in lines[i]:
                        pered = 'доступная-память">'
                        posle = '</td>'
                        text6 = poisk(pered, posle)
                        text6 = text6.split(' Мбайт')[0]
                
                    '''allMemor'''
                    if 'общая-память">' in lines[i]:
                        pered = 'общая-память">'
                        posle = '</td>'
                        text7 = poisk(pered, posle)
                        text7 = text7.split(' Мбайт')[0]
    
            zond.append(text0)
            dayList.append(text1)
            datatimes.append(text2)
            totals.append(text3)
            svobod.append(text4)
            procDostyp.append(text5)
            dostypMemory.append(text6)
            allMemor.append(text7)


    # print(f'{tipdata}, {filename}')
    # fffeee = [zond, dayList, datatimes, totals, svobod, procDostyp, dostypMemory, allMemor]
    # for i in fffeee:
    #     print(f'{len(i)}')
    
    dlockZGD = []
    for i in range(len(zond)):
        if zond[i] in ZGDDict:
            xxx = ZGDDict[zond[i]]
        else:
            xxx = ''
        dlockZGD.append(xxx)


    data = [[tipdata, filename, zond[i], dlockZGD[i], dayList[i], datatimes[i], totals[i], svobod[i], procDostyp[i], dostypMemory[i], allMemor[i]] for i in range(len(datatimes))]
    data += data

    return data

def on_change_err(s):
    '''Сообщение об ошибке'''
    QtWidgets.QMessageBox.information(Form, 'Excel не отвечает...', s)


def importdata(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Собираем данные из диапозона ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    vals = cel.Formula
    if StartCol == EndCol:
        vals = [vals[i][x] for i in range(len(vals)) for x in range(len(vals[i]))]
    return vals


def start():
    global ZGDDict
    ui.label.setStyleSheet("color: rgb(0, 0, 0);")
    ui.label.setText("Обработка данных . . .")
    directory = str(ui.plainTextEdit_10.toPlainText())
    # directory = r'C:\Users\vvkhomutskiy\Desktop'
    if directory == '':
        on_change_err('Не указана папка с файлами в формате "html"')
        return
    direct = os.listdir(directory)

    pythoncomCoInitializeEx(0)
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = 0

    '''================================================'''
    '''Блок ЗГД'''
    wbZGD = Excel.Workbooks.Open(os.getcwd() + "\ZGD.xlsx")
    sheet = wbZGD.Worksheets("Свод")
    sheet.Activate()
    count_row = sheet.UsedRange.Rows.Count
    StartRow = 2
    EndRow = StartRow + count_row - 1

    col_1 = importdata(sheet, StartRow, 1, EndRow, 1)
    col_2 = importdata(sheet, StartRow, 3, EndRow, 3)

    ZGDDict = {}
    for i in range(len(col_1)):
        ZGDDict[col_1[i]] = col_2[i]

    '''Закрыть файл без сохранения'''
    # print(col_1)
    # print(ZGDDict)
    wbZGD.Close(False)
    '''================================================'''

    Excel.Visible = 1
    wb = Excel.Workbooks.Open(os.getcwd() + "\Метрики ПК_шаблон.xltx")
    TypeList = ['CPU_data', 'HDD_data', 'MEMORY_data']
    sheet = wb.Worksheets("CPU_data")
    sheet.Activate()

    data = []
    NameColums = [[
            'Тип',
            'Имя файла',
            'Устройство',

            'Блок ЗГД',

            'Дата',
            'Время',
            'Всего\nCPU',
            'Свободное\nпространство',
            'Процент\nдоступной\nпамяти',
            'Доступная\nпамять,\nМбайт',
            'Общая\nпамять,\nМбайт'
            ]]


    for filename in direct:
        FullName = os.path.join(directory, filename)
        if os.path.isfile(FullName) and ".html" in filename:
            # data = data + GO(FullName, filename, zond, dayList, datatimes, totals, text0, data, svobod, procDostyp, dostypMemory, allMemor)
            data = data + GO(FullName, filename)

    '''Отправляем в Excel'''
    def otpravka(sheet, data):
        StartRow, StartCol = 1, 1
        EndCol = 10
        len_data_full = len(data)
        maxCountRow = 100000

        NameColums
        exportdata(NameColums, sheet, 1, 1, 1, EndCol)
        # StartRow = 2
        if len_data_full > maxCountRow:
            xxx = len_data_full / maxCountRow
            counrange = math_ceil(xxx)
            
            for i in range(1, counrange + 1):
                EndRow = maxCountRow * i if maxCountRow * i < len_data_full else len_data_full
                print(f'EndRow = {EndRow}')
                print(f'{StartRow - 1} :--: {EndRow}')
                dataX = data[StartRow - 1 : EndRow]
                exportdata(dataX, sheet, StartRow + 1, StartCol, EndRow, EndCol)
                StartRow = EndRow
                sleep(1)
        else:
            EndRow = len(data)
            EndCol = 10
            exportdata(data, sheet, StartRow + 1, StartCol, EndRow, EndCol)
            sleep(1)
        
        '''Создание таблицы со стилем'''
        cels = sheet.Range(sheet.Cells(1, 1), sheet.Cells(EndRow, EndCol))
        cels.Select
        sheet.ListObjects.Add(1, cels, True, 1)        
        '''Выравнивание ячеек'''
        sheet.Rows(1).HorizontalAlignment = 3
        sheet.Rows(1).VerticalAlignment = 2
        ColSelect = sheet.Range(sheet.Columns(1), sheet.Columns(EndCol))
        ColSelect.EntireColumn.AutoFit()
        ColSelect.HorizontalAlignment = 3
        ColSelect.VerticalAlignment = 2

    '''Разбиваем общий список на типы данных'''
    CPUList = []
    HDDList = []
    MEMList = []
    for i in data:
        if i[0] == 'CPU':
            CPUList.append(i)
        if i[0] == 'HDD':
            HDDList.append(i)
        if i[0] == 'MEMORY':
            MEMList.append(i)

    AllListdata = [CPUList, HDDList, MEMList]
    for i in range(len(TypeList)):
        if AllListdata[i] != []:
            sheet = wb.Worksheets(TypeList[i])
            sheet.Activate()
            sleep(0.3)
            otpravka(sheet, AllListdata[i])
    
    '''Формирование финальных вкладок с диаграммами и сводными таблицами'''
    FiList = ['CPU', 'HDD', 'RAM']
    for i in range(len(FiList)):
        try:
            SvodTable.GO(wb.Worksheets(FiList[i]), wb, TypeList[i])
            sleep(0.5)
        except:
            pass

    ui.label.setStyleSheet("color: rgb(0, 170, 0);")
    ui.label.setText("Сборка файла завершена . . .")


'''Отслеживаем сигнал в plainTextEdit на изменение данных и удаляем не нужный текст'''
def ChangedPT(plainTextEdit):
    '''Удаления ненужного текста в plainTextEdit_3'''
    directory = plainTextEdit.toPlainText()
    if "file:///" in directory:
        xxx = directory.rfind("file:///")
        directory = directory[xxx + 8:]
        try:
            directory = directory.replace("/", "\\")
        except:
            pass
        plainTextEdit.setPlainText(rf"{directory}")
ui.plainTextEdit_10.textChanged.connect(lambda : ChangedPT(ui.plainTextEdit_10))

# ui.plainTextEdit_10.clear()

def thread(my_func):
    '''Обертка функции в потопк (декоратор)'''
    def wrapper():
        threading.Thread(target=my_func, daemon=True).start()
    return wrapper

@thread
def pysk():
    try:
        start()
    except:
        errortext = traceback.format_exc()
        print(errortext)
        text = f"Ошибка работы, повторите попытку \n\n{errortext}"
        on_change_err(text)

@thread
def redactZGD():
    '''Получаем доступ к определенному файлу'''
    pythoncomCoInitializeEx(0)
    Excel = win32com.client.Dispatch("Excel.Application")
    # Excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    Excel.Visible = 1
    Excel.Workbooks.Open(os.getcwd() + "\ZGD.xlsx")

ui.pushButton_4.clicked.connect(pysk)
ui.pushButton_5.clicked.connect(redactZGD)

if __name__ == "__main__":
    sys.exit(app.exec_())

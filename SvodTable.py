from time import sleep
from unicodedata import category
import win32com.client

def ifErr(formula):
    '''Убираем ошибку при пустом значении'''
    '''Удаляем из фомрулы знак (=)'''
    formula = formula.lstrip('=')
    iferror = f"IFERROR({formula},\"\")"
    text = f"=IF({iferror}=0,\"\",{iferror})"
    return text

def GO(sheet, wb, page):
    sheet.Activate()
    sleep(0.3)
    '''Подключаемся к сводной таблице 1'''
    PivotTab = sheet.PivotTables(1)
    print(PivotTab.Name)

    '''Количество занимаемых таблицей строк'''
    count_row_istok_tab = wb.Worksheets(page).UsedRange.Rows.Count
    PivotTab.ChangePivotCache(wb.PivotCaches().Create(1, f"{page}!R1C1:R{count_row_istok_tab}C10", 5))
    count_col_end = sheet.UsedRange.Columns.Count
    count_row_end = sheet.UsedRange.Rows.Count
    niztab = count_row_end - 42
    sred = ifErr(f"=AVERAGE(R[5]C:R[{niztab}]C)")
    stoka41 = [sred] * (count_col_end - 1)
    sheet.Range(sheet.Cells(41, 2), sheet.Cells(41, count_col_end)).Formula = stoka41
    sleep(0.5)

    '''Верхняя таблица для диаграмм'''
    countColVerh = (count_col_end - 2) * 0.5
    countColVerh = int(countColVerh)


    '''Формулы для верхней таблицы'''
    stoka2 = [ifErr(r"=INDIRECT(ADDRESS(43,(COLUMN()-1)*2))")] * countColVerh
    stoka3 = [ifErr(r"=INDIRECT(ADDRESS(44,(COLUMN()-1)*2))")] * countColVerh
    stoka4 = [ifErr(r"=INDIRECT(ADDRESS(41,(COLUMN()-1)*2))")] * countColVerh
    stoka5 = [ifErr(r"=INDIRECT(ADDRESS(41,(COLUMN()-1)*2+1))")] * countColVerh
    stoka6 = [ifErr("=GETPIVOTDATA(\"Макс\",R42C1,\"Устройство\",R[-3]C)")] * countColVerh
    stoka7 = [ifErr("=GETPIVOTDATA(\"Мин\",R42C1,\"Устройство\",R[-4]C)")] * countColVerh


    def otpravraVerx(row1, row2, data):
        sheet.Range(sheet.Cells(row1, 2), sheet.Cells(row2, countColVerh + 1)).Formula = data

    rowVer = [2, 3, 4, 5, 6, 7]
    strokaList = [stoka2, stoka3, stoka4, stoka5, stoka6, stoka7]
    for i in rowVer:
        xxx = rowVer.index(i)
        otpravraVerx(i, i, strokaList[xxx])

    '''Границы'''
    cel = sheet.Range(sheet.Cells(2, 1), sheet.Cells(7, countColVerh + 1))
    cel.Borders.Weight = 2


    def redactChart(wb, NameChart, Rows):
        '''Активируем диаграмму'''
        sheet.ChartObjects(NameChart).Activate()
        '''Задаем произвольный диапозон ячеек, после автоматом изменится'''
        cel = sheet.Range(sheet.Cells(2, 1), sheet.Cells(5, countColVerh + 1))
        wb.ActiveChart.SetSourceData(cel)

        def redactRow(nomerRow, category, Name, Values):
            '''Элементы легенды (ряды) - Имя'''
            wb.ActiveChart.FullSeriesCollection(nomerRow).Name = Name
            '''Элементы легенды (ряды) - Значение'''
            wb.ActiveChart.FullSeriesCollection(nomerRow).Values = Values
            '''Подписи горизонтальной оси (категория)'''
            wb.ActiveChart.FullSeriesCollection(nomerRow).XValues = category

        category = sheet.Range(sheet.Cells(2, 2), sheet.Cells(3, countColVerh + 1))
        for row in Rows:
            nomerRow = Rows.index(row) + 1
            if nomerRow > 1:
                '''Создаем 2ой рад, 1ый уже есть, но не отображается'''
                wb.ActiveChart.SeriesCollection().NewSeries()
            '''Задаем имя ряда'''
            Name = f"={sheet.Name}!A{row}"
            Values = sheet.Range(sheet.Cells(row, 2), sheet.Cells(row, countColVerh + 1))
            redactRow(nomerRow, category, Name, Values)


    NameChart = "Диаграмма 2"
    Rows = [4, 5]
    redactChart(wb, NameChart, Rows)

    NameChart = "Диаграмма 1"
    Rows = [6, 7]
    redactChart(wb, NameChart, Rows)

    '''Ширина диаграммы'''
    sheet.Shapes("Диаграмма 2").Width = 1700
    sheet.Shapes("Диаграмма 1").Width = 1700
    sleep(1)

    '''Ширина всех колонок'''
    sheet.Cells.ColumnWidth = 11

if __name__ == "__main__":
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = 1
    wb = Excel.ActiveWorkbook
    '''Получаем доступ к определенному файлу'''
    # wb = Excel.Workbooks.Open(r"C:\vxvproj\tnnc-ReadHTML\Метрики ПК_шаблон.xltx")
    # sheet = wb.ActiveSheet
    sheet = wb.Worksheets("CPU")
    sheet.Activate()

    GO(sheet, wb, 'CPU_data')
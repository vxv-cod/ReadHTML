# aaa = '<tr class="odd"><td class="col-datetime"><nobr>22.09.2022 22:45:00 − 23:00:00</nobr></td><td class="col-value col-процент-доступной-памяти">88 %</td><td class="col-value col-доступная-память">43 218 Мбайт</td><td class="col-value col-общая-память">49 135 Мбайт</td><td class="col-value col-время-простоя">0 %</td><td class="col-coverage">100 %</td></tr>'

# fff = aaa.find('<td', 17)
# print(fff)

# import win32com
# print(win32com.__gen_path__)


fff = ['CPU', 'HDD', 'MEMORY']
ddd = '27 PC MEMORY _ Отчет _ TNNC-WORKSTATION-MONITORING'
# for i in fff:
if fff not in ddd:
    tipdata = ''
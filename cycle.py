# -*- coding: utf-8 -*-
# цикл который заполняет списки информацией которую ввожу!
# сделать что бы запоняло по строкам!
# сделать что бы в кс писалось что нужно вводить

from openpyxl import *

wb = Workbook()
ws = wb.active

ws.title = "список"

a = []
b = ['ПІБ: > ','Адреса: > ','Виконавець: > ', 'Дата та час виклику: > ',
     'Дата та час прибуття на місце виклику: > ','Час у дорозі: > ',
     'Причина аварії: > '
    ]

for j in range(1, 3):
    print ("new 'j' = ", j)
    for i in range(1, 4):
        a.append(input("a.%d.%d > " % (j, i)))
    ws.append(a)
    a = []
    
#for i in range(1, len(b)):
#    ws.append(b[i-1])

wb.save('fun.xlsx')
 


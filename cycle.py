from openpyxl import *
from datetime import *

wb = Workbook()
ws = wb.active

ws.title = "список"

a, b = [], []
column = ['ПІБ абонента: > ','Адреса: > ']

worker = input("Старший слюсар зміни:> ")   
print("Введіть дату")    
d = date(
     int(input("рік> ")), 
     int(input("місяць> ")), 
     int(input("день> "))
    )
q = int(input("скільки %s було заявок по витокам ? > " % d))      

j = True
while j:           
    if q > 0:               # цикл отвечающий за количество заявок
        for i in range(1, 3):            
            #создание списка с инф. о утечках и абонетах
            a.append(input("%s" % (column[i-1])))   
            
        a.append(worker) 
        #список дат нужно будет для прикрепления к инормации по заявкам, 
        # а так же для вычисления времени потраченого в пути 
        print("час отримання заявки:")  
        t1 = time(
                   int(input("годин> ")),
                   int(input("хвилин> "))
                 )
        a.append(str(datetime.combine(d, t1)))
        
        print("час виїзду на заявку:")  
        t2 = time(
                   int(input("годин> ")),
                   int(input("хвилин> "))
                 )
        a.append(str(datetime.combine(d, t2)))
        
        #разница между временем получения и выездом
        a.append(abs(datetime.combine(d, t2) - datetime.combine(d, t1)))
        #вставка пустої комірки
        a.append('')
        #причина аварии
        a.append(input("введіть причину витоку газу> "))

        ws.append(a)                    #append prev list to active list
        a = []                          # очистка списка данными по абонентам
        q -= 1      #cчетчик заявок в один день
    
    elif q == 0:
        print ("змінити дату - ’д’, закінчити - ’к’")
        j = input("> ")
        if j == 'д':
            d = date(
                     int(input("рік> ")), 
                     int(input("місяць> ")), 
                     int(input("день> "))
                    )
            q = int(input("скільки %s було заявок по витокам ? > " % d))             
            j = True
        elif j == 'к':
            wb.save('fun.xlsx')
            j = False
        else:   
            print('Ти ввів щось не так, але файл всерівно зберігся. ;)')
            wb.save('fun.xlsx')
            j = False
 
















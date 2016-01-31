from openpyxl import *
from datetime import *

wb = Workbook()
ws = wb.active

ws.title = "список"

a, b = [], []
column = ['ПІБ абонента: > ','Адреса: > ']

while True:
    try:
        print("Введіть дату")    
        d = date(
             int(input("рік> ")), 
             int(input("місяць> ")), 
             int(input("день> "))
            )
        break
    except ValueError:
        print("Не вірна дата! Пробуй ще")

while True:
    try:
        q = int(input("скільки %s було заявок по витокам ? > " % d))
        num_ask = int(input("Ведіть номер першої заявки цього дня> "))      
        break
    except ValueError:
        print("Помилка с кількістю заявок, давай ще.")

while True:                 #кількість заявок
    if q > 0:               
        while True:         #номер заявки
            try:
                num_ask = int(input("Попередня заявка - {0}, нова > ".format(num_ask)))
                break
            except ValueError:
                print("Не вірно ввів номер заявки, давай ще раз! ;)")
        
        #ПІБ та адреса
        for i in range(1, 3):            
            a.append(input("%s" % (column[i-1])))   
            
        
        while True:         #вибір слюсаря   
            worker_list = ['Стасюк О.А','Кравчук А.М.',
                           'Кучер О.В.','Бортяний П.А.'
                          ]
            print("Виберіть слюсаря, який виконував заявку")
            j = 0
            for i in worker_list:
                print("Якщо {0} натисніть - {1}".format(i, j))
                j += 1
            w = int(input("> "))   
            if w == 0:
                worker = worker_list[0]
                break
            elif w == 1:
                worker = worker_list[1]
                break
            elif w == 2:
                worker = worker_list[2]
                break
            elif w == 3:
                worker = worker_list[3]
                break
            elif ((w < 0) or (w > 3)):
                print("Ви ввели не вірне число, спробуйте ще")
            else:
                print("Ви ввели який текст, спробуйте ще")

        #додавання слюсаря до списку
        a.append(worker) 

        while True:
            while True:
                try:
                    print("час отримання заявки:")  
                    t1 = time(
                               int(input("годин> ")),
                               int(input("хвилин> "))
                             )
                    break
                except ValueError:
                    print("Не вірний час отримання, вводь ще")            
            a.append(str(datetime.combine(d, t1)))
            
            while True:
                try:
                    print("час прибуття на заявку:")  
                    t2 = time(
                               int(input("годин> ")),
                               int(input("хвилин> "))
                             )
                    break
                except ValueError:
                    print("Не вірний час прибуття, вводь ще")
            a.append(str(datetime.combine(d, t2)))
            
            dt1 = datetime.combine(d, t1)
            dt2 = datetime.combine(d, t2)
            dts = dt2.minute - dt1.minute
            
            if abs(dts) <= 40:
                break
            elif abs(dts) == 0:
                print("Різниця в хвилинах між заявками 0, шукай помилку")
            else:
                print("Між отриманням та прибуттям більше 40 хвилин, вводь ще раз")                     
            
            
        

        
        a.append(abs(datetime.combine(d, t2) - datetime.combine(d, t1)))
        
        
        
        
        #вставка пустої комірки
        a.append('')
        #причина аварии
        while True:
            print("""Де іде витік газу(виберіть цифру):
            1 - стояк н.т;
            2 - стояк с.т;
            3 - ОК;
            4 - ПГ;
            5 - ГК;
            6 - конвектор;
            7 - флянець;
            8 - РДГ;
            9 - лічильник газу
            0 - інша причина """)
            reason = int(input("> "))
            if reason == 1:
                reason = int(input("""вибірть цифру:
                1 - На самому стояку
                2 - На крані стояка
                3 - На продувній заглушці
                > """))
                if reason == 1:
                    a.append("""Витік газу на стояку н.т. """)
                    break
                elif reason == 2:
                    a.append("""Витік газу на крані стояка н.т. """)
                    break
                elif reason == 3:
                    a.append("""Витік газу на продувній заглушці стояка н.т. """)
                    break
                else:
                    print("Ви допустили помилку - вводьте ще раз!")
            elif reason == 2:       #витік на стояку с.т.
                reason = int(input("""вибірть цифру:
                1 - На самому стояку
                2 - На крані стояка """))
                if reason == 1:
                    a.append("""Витік газу на стояку c.т. """)
                    break
                elif reason == 2:
                    a.append("""Витік газу на крані стояка c.т """)
                    break
                else:
                    print("Ви допустили помилку - вводьте ще")
            elif reason == 3:       #витік на ОК
                reason = int(input("""вибірть цифру:
                1 - На автоматиці до ОК
                2 - На крані опуску до ОК
                3 - На різьбових з’єднаннях опуску до ОК
                4 - На гнучкому газопроводі до ОК
                > """))
                if reason == 1:
                    a.append("""Витік газу на атоматиці ОК""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на крані опуску до ОК""")
                    break
                elif reason == 3:
                    a.append("""Витік газу на різбовому з’днанні опуску до ОК""")
                    break
                elif reason == 4:
                    a.append("""Витік газу на гнучкому газопроводі до ОК""")
                    break
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 4:       #витік на ПГ
                reason = int(input("""вибірть цифру:
                1 - На ПГ-4
                2 - На крані опуску до ПГ-4
                3 - На різьбових з’єднаннях опуску до ПГ-4
                4 - На гнучкому газопроводі до ПГ-4
                > """))
                if reason == 1:
                    a.append("""Витік газу на ПГ-4""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на крані опуску до ПГ-4""")
                    break
                elif reason == 3:
                    a.append("""Витік газу на різбовому з’днанні опуску до ПГ-4""")
                    break
                elif reason == 4:
                    a.append("""Витік газу на гнучкому газопроводі до ПГ-4""")
                    break
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 5:       #витік на ГК
                reason = int(input("""вибірть цифру:
                1 - На ГК
                2 - На крані опуску до ГК
                3 - На різьбових з’єднаннях опуску до ГК
                4 - На гнучкому газопроводі до ГК
                > """))
                if reason == 1:
                    a.append("""Витік газу на ГК""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на крані опуску до ГК""")
                    break
                elif reason == 3:
                    a.append("""Витік газу на різбовому з’днанні опуску до ГК""")
                    break
                elif reason == 4:
                    a.append("""Витік газу на гнучкому газопроводі до ГК""")
                    break
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 6:       #витік на конвекторі
                reason = int(input("""вибірть цифру:
                1 - На конвекторі
                2 - На крані опуску до конвекторі
                3 - На різьбових з’єднаннях опуску до конвекторі
                4 - На гнучкому газопроводі до конвектора
                > """))
                if reason == 1:
                    a.append("""Витік газу на конвекторі""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на крані опуску до конвектора""")
                    break
                elif reason == 3:
                    a.append("""Витік газу на різбовому з’днанні опуску до конвектора""")
                    break
                elif reason == 4:
                    a.append("""Витік газу на гнучкому газопроводі до конвектора""")
                    break
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 7:       #витік на флянці
                reason = int(input("""вибірть цифру:
                1 - На флянці стояка н.т.
                2 - На флянці с.т. 
                > """))

                if reason == 1:
                    a.append("""Витік газу на флянці стояка н.т.""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на флянці стояка с.т.""")
                    break
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 8:       #витік на РДГ
                reason = int(input("""вибірть цифру:
                1 - На РДГ
                2 - На вході в РДГ 
                3 - На виході з РДГ
                > """))

                if reason == 1:
                    a.append("""Витік газу на РДГ""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на вході в РДГ""")
                    break
                elif reason == 3:
                    a.append("""Витік газу на виході в РДГ""")
                    break                
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 9:       #витік на ЛГ
                reason = int(input("""вибірть цифру:
                1 - На самому лічильнику
                2 - На штуцерах ЛГ
                > """))

                if reason == 1:
                    a.append("""Витік газу на лічильнику""")
                    break
                elif reason == 2:
                    a.append("""Витік газу на штуцерах лічильника""")
                    break
                else:
                    print("Ви допустили помилку - вводьте ще") 
            elif reason == 0:
                a.append(input("Ведіть іншу причину витоку: >"))
                break
            else:
                print("Ви ввели щось невірно, вводьте ще")

        ws.append(a)                    #append prev list to active list
        wb.save('autosave.xlsx')
        a = []                          # очистка списка данными по абонентам
        q -= 1      #cчетчик заявок в один день
    
    elif q == 0:
        print ("змінити дату - ’д’, закінчити - ’к’")
        j = input("> ")
        if j == 'д':
            while True:
                try:
                    print("Введіть дату")    
                    d = date(
                         int(input("рік> ")), 
                         int(input("місяць> ")), 
                         int(input("день> "))
                        )
                    break
                except ValueError:
                    print("Не вірна дата! Пробуй ще")

            while True:
                try:
                    q = int(input("скільки %s було заявок по витокам ? > " % d))      
                    break
                except ValueError:
                    print("Помилка с кількістю заявок, давай ще.")         
            continue
        elif j == 'к':
            name_wb = input("Введіть ім’я файла для збереження>")
            wb.save(name_wb + '.xlsx')
            break
        #break
        else:
            print("Ви ввели не те, спробуйте ще!")
        
















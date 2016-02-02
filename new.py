# coding: cp1251
a = ['яблука','нас','не пруть']
b = a.encode('cp1251').decode('utf-8')
f = open('text.txt', 'w')
f.write(a)
f.close()
#print(a.encode('utf-8').decode('cp1251'))

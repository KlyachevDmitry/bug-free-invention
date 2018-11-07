import openpyxl
import socket
import sys

#Создаем объект wb
wb = openpyxl.load_workbook(filename = 'C:/Users/Дмитрий/AppData/Local/Programs/Python/Python36/testX.xlsx')

#активируем лист в эксель
ws = wb.active

#Выбираем страницу
sheet = wb['test']

#считываем значения из ячеек
vals = [sheet.cell(row=1,column=i).value for i in range(1,11)]

#преобразуем массив в строку
vals=str(vals)

#создаем сокет "conn"
conn = socket.socket()

#подключаемся к машине
conn.connect( ("192.168.1.102", 14900) )

#декодируем строки в байты и отправляем 
conn.send(b"Hello! \n")
conn.send(vals.encode('utf-8'))

while 1:
    a=input("Commands: ")
    conn.send(a.encode('utf-8'))
    if (a=="exit"):
        conn.send(b"exit")
        break
    a=b""
conn.close()


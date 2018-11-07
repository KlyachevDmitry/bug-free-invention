import socket
import openpyxl

#Создаем сокет sock
sock=socket.socket()

#Для привязки используется функция bind сокета, которая принимает массив, содержащий два элемента: хост и порт.
sock.bind(("",14900))

#Слушаем соединения
sock.listen(10)

#принимаем соединения с помощью функции accept
conn, addr=sock.accept()

# функцией recv контролируем еколичество получаемых байт
data=conn.recv(1000)
udata=data.decode("utf-8")
print(udata)

#декодируем и отправялем данные
exel=conn.recv(1024)
uexel=exel.decode("utf-8")

#закрываем сокет
conn.close()

#
wb = openpyxl.load_workbook(filename = 'C:/Users/deman/AppData/Local/Programs/Python/Python36/Projects/testX.xlsx')
#
ws = wb.active

#открываем страницу
sheet = wb['test']
uexel = eval(uexel)

#записываем последовательность
i = 1
for rec in uexel:
    rec=float(rec)
    sheet.cell(row=1, column=i).value = rec
    i += 1
#сохраняем файл
wb.save('C:/Users/deman/AppData/Local/Programs/Python/Python36/Projects/testX3.xlsx')

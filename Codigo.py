import pyodbc
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from tkinter import *
from tkinter import filedialog

def exit():
 exit = win.withdraw()
#seleciona o arquivo mdb
def file_dba():
  file_db = filedialog.askopenfilename(defaultextension=".mdb")
  return file_db
#salva o caminho em uma variavel
def save_caminho():
    global DB
    DB = file_dba()

def conect_saves():
# estabelecer conexão com o banco Access
 conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ="+ DB
    )
 conn = pyodbc.connect(conn_str)
 cursor = conn.cursor()
 
# selecionar dados de uma tabela
 cursor.execute('SELECT Id, NomeProfissional FROM Cadastro_Profissionais')
 global rows
 rows = cursor.fetchall()
 
# fechar a conexão
 conn.close()
 win.withdraw()
 # seleciona onde vai o arquivo pdf salvo
 file_save = filedialog.asksaveasfilename(defaultextension=".pdf")
 cnv = canvas.Canvas(file_save, pagesize=A4)
 x = 50
 y = 750
# Adiciona as informações extraídas da tabela ao documento PDF
 for row in rows:
  Id, NomeProfissional = row
  cnv.drawString(x, y, f"ID: {Id}")
  cnv.drawString(x, y-20, f"Nome: {NomeProfissional}")
    
#Criação de retângulos e estão separados por localização
 cnv.rect(48,745,30,19)
 cnv.rect(48,723,90,19)
 cnv.save()


# Janela 
win = Tk()
win.title('Informações')
win.geometry("300x300")


Btn = Button(win, width=10, text="Press", command=save_caminho)
Btn.pack()
Btn.place(x=10, y=10)

Btnex = Button(win, width=10, text="Exit", command=exit)
Btnex.pack()
Btnex.place(x=200, y=10)

Btnsv = Button(win, width=10, text="Conect/Save", command=conect_saves)
Btnsv.pack()
Btnsv.place(x=10, y=200)

win.mainloop()


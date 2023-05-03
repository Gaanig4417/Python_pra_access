DB = input('Local base ')
import pyodbc
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# estabelecer conexão com o banco Access
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ="+DB + r"\base.mdb"
    )
  
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
# selecionar dados de uma tabela
cursor.execute('SELECT Id, NomeProfissional FROM Cadastro_Profissionais')
rows = cursor.fetchall()
# fechar a conexão
conn.close()
cnv = canvas.Canvas("arquivo.pdf")
x = 50
y = 750

# Adiciona as informações extraídas da tabela ao documento PDF
for row in rows:
    Id, NomeProfissional = row
    cnv.drawString(x, y, f"ID: {Id}")
    cnv.drawString(x, y-20, f"Nome: {NomeProfissional}")
    y -= 60


cnv.save()


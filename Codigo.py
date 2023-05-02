DB = input('Local base ')
import pyodbc
from fpdf import FPDF

# estabelecer conexão com o banco Access
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ="+DB
    )
  
conn = pyodbc.connect(conn_str)

# selecionar dados de uma tabela
cursor = conn.cursor()
cursor.execute('SELECT Id, NomeProfissional FROM Cadastro_Profissionais')
columns_info = cursor.description

# Obtém o cabeçalho das colunas
columns_header = [column[0] for column in columns_info]
rows = cursor.fetchall()

# Cria um objeto FPDF
pdf = FPDF()
pdf.add_font('DejaVuSans', '', 'DejaVuSans.ttf', uni=True)
pdf.set_font('DejaVuSans', '', 12)
pdf.add_page()

data = {}
# Adiciona as informações extraídas da tabela ao documento PDF
for row in rows:
    Id = row 
    NomeProfissional = row
    data = [Id] = {"ID": Id}
    data = [NomeProfissional] = {"Nome Profissional": NomeProfissional}
    pdf.cell(10, 10, str(Id), border=1)
    pdf.cell(45, 10, str(NomeProfissional), border=1)
    pdf.ln()
    pdf.cell(0, 10, str(row), border=1)
    pdf.ln()
   

# Salva o documento PDF
pdf.output('arquivo.pdf')

# fechar a conexão
conn.close()

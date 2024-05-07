import os
import tkinter as tk
from tkinter import filedialog, scrolledtext

import openpyxl
import pandas as pd
import pyodbc

from sql import carteira, cliente_query, delete, insert, query, vendedor_query

# # Construir a string de conexão
# connection_string = (
#     r'DRIVER={SQL Server};'
#     r'SERVER=192.168.1.22;'
#     r'DATABASE=TESTE;'
#     r'UID=sa;'
#     r'PWD=Moitgt2526;'
# )

connection_string = (
    r'DRIVER={SQL Server};'
    r'SERVER=192.168.2.10;'
    r'DATABASE=MOINHO;'
    r'UID=sa;'
    r'PWD=moitgt2526;'
)

# # Consulta SQL para verificar a existência da combinação (cd_clien, cd_vend)
# query = "SELECT * FROM dbo.clientelev WHERE cd_clien = ? AND cd_vend = ?"
# insert = "INSERT INTO dbo.clientelev (cd_clien,cd_vend) VALUES (?, ?);"
# delete = "DELETE FROM dbo.clientelev WHERE cd_clien = ? AND cd_vend = ?"
# cliente_query = "SELECT * FROM dbo.cliente WHERE cd_clien = ?"
# vendedor_query = "SELECT * FROM dbo.vendedor where cd_vend = ?"
# carteira = """select
#                     c.cd_clien,
#                     c.cd_vend,
#                     '' as '           ',
#                     cl.cd_clien codCliente,
#                     cl.nome cliente,
#                     ec.estado,
#                     vd.cd_vend codVendedor,
#                     vd.nome vendedor,
#                     eq.descricao equipe
#                 from
#                     clientelev c
#                     left join cliente cl on cl.cd_clien = c.cd_clien
#                     left join vendedor vd on vd.cd_vend = c.cd_vend
#                     left join end_cli ec on ec.cd_clien = cl.cd_clien and ec.tp_end = 'FA'
#                     left join equipe eq on eq.EquipeId = vd.EquipeID"""

# Função para conectar ao banco de dados e verificar a existência da combinação (cd_clien, cd_vend)
def verificar_combinacao(cd_clien, cd_vend, aba):
    # Converter os valores para string
    cd_clien_str = str(cd_clien)
    cd_vend_str = str(cd_vend)

    # Conectar ao banco de dados
    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    cursor.execute(query, (cd_clien_str, cd_vend_str))
    row = cursor.fetchall()

    cursor.execute(cliente_query, cd_clien)
    cliente_row = cursor.fetchall()

    # Verificar se o cd_cliente existe
    if not cliente_row:
        text_box.insert(tk.END, f'\n(cd_clien = {cd_clien_str}, cd_vend = {cd_vend_str}) - cd_clien {cd_clien_str}: INVÁLIDO')
        # Fechar a conexão com o banco de dados
        return

    cursor.execute(vendedor_query, cd_vend)
    vend_row = cursor.fetchall()

    if not vend_row:
        text_box.insert(tk.END, f'\n(cd_clien = {cd_clien_str}, cd_vend = {cd_vend_str}) - cd_vend {cd_vend_str}: INVÁLIDO')
        # Fechar a conexão com o banco de dados
        return

    # Verificar se a combinação existe
    if row:
        mensagem = f"\n(cd_clien = {cd_clien_str}, cd_vend = {cd_vend_str}) já existe."
        text_box.insert(tk.END, mensagem)
        if aba == "Remover":
            # mensagem_delete = "Executando DELETE..."
            # text_box.insert(tk.END, mensagem_delete)
            cursor.execute(delete, (cd_clien_str, cd_vend_str))
            conn.commit()
            mensagem = f"\n - DELETADO COM SUCESSO."
            text_box.insert(tk.END, mensagem)
        else:
            mensagem = f"\n(cd_clien = {cd_clien_str}, cd_vend = {cd_vend_str}) - Já existe."
            text_box.insert(tk.END, mensagem)
    else:
        mensagem = f"\n(cd_clien = {cd_clien_str}, cd_vend = {cd_vend_str}) - não existe."
        text_box.insert(tk.END, mensagem)
        if aba == "Incluir":
            cursor.execute(insert, (cd_clien_str, cd_vend_str))
            conn.commit()
            mensagem = f"\n - INCLUÍDO  COM SUCESSO."
            text_box.insert(tk.END, mensagem)
        else:
            mensagem = f"\n(cd_clien = {cd_clien_str}, cd_vend = {cd_vend_str}) - Não existe."
            text_box.insert(tk.END, mensagem)
    
def extrair_carteira():
    # Executar a consulta para obter a carteira de televendas
    conn = pyodbc.connect(connection_string)
    carteira_df = pd.read_sql_query(carteira, conn)

    # Fechar a conexão com o banco de dados
    conn.close()
    
    # Abrir uma janela para escolher o local para salvar o arquivo
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel", "*.xlsx")],initialfile=f"Carteira Televendas Nexfar.xlsx")
    if filename:
        # Salvar o resultado em um arquivo Excel
        carteira_df.to_excel(filename, index=False)
        # print("Carteira de Televendas Nexfar salva com sucesso em:", filename)

# Função para selecionar uma planilha do computador
def selecionar_planilha():
    # Limpar a tela de mensagens
    text_box.delete('1.0', tk.END)
    
    # Abrir uma janela para selecionar um arquivo
    filename = filedialog.askopenfilename(filetypes=[("Planilha Excel", "*.xlsx")])
    if filename:
        try:
            # Exibir o nome do arquivo lido
            nome_arquivo = os.path.basename(filename)
            text_box.insert(tk.END, f"Arquivo: {nome_arquivo}")

            # Ler a planilha
            workbook = openpyxl.load_workbook(filename)
            for sheet_name in workbook.sheetnames:
                text_box.insert(tk.END, f"\n\nAba: {sheet_name}\n")
                sheet = workbook[sheet_name]
                # Verificar se as colunas esperadas estão presentes
                expected_columns = ("cd_clien", "cd_vend")  # Colunas esperadas
                header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
                if header_row is None or not all(column_name in header_row for column_name in expected_columns):
                    raise ValueError("A planilha não possui as colunas esperadas.")
                # Variável para pular a primeira linha
                primeira_linha = True
                for row in sheet.iter_rows(min_row=2, values_only=True):  # Começar da segunda linha
                    # Verificar se a linha está vazia
                    if any(cell is not None for cell in row):
                        # Verificar a existência da combinação (cd_clien, cd_vend)
                        verificar_combinacao(row[0], row[1], sheet_name)
            # Fechar a planilha após o uso
            workbook.close()
        except FileNotFoundError:
            text_box.insert(tk.END, "Erro: Arquivo não encontrado.")
        except ValueError as e:
            text_box.insert(tk.END, f"Erro: {str(e)}")


# Criar a janela tkinter
root = tk.Tk()
root.title("Selecionar Planilha")

# Componente de texto com barra de rolagem para exibir mensagens
text_box = scrolledtext.ScrolledText(root, height=20, width=100)
text_box.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

# Botão para selecionar a planilha
button_selecionar_planilha = tk.Button(root, text="Selecionar Planilha", command=selecionar_planilha)
button_selecionar_planilha.pack(padx=20, pady=10)

# Botão para extrair carteira
button_extrair_carteira = tk.Button(root, text="Extrair Carteira", command=extrair_carteira)
button_extrair_carteira.pack(padx=20, pady=10, side=tk.LEFT)

root.mainloop()
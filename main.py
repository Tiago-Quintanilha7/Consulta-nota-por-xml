import tkinter as tk
from tkinter import messagebox
import xmltodict
import os
import pandas as pd
from tkinter import filedialog


def escolher_pasta():
    pasta = filedialog.askdirectory()
    if pasta:
        print(f'Pasta selecionada: {pasta}')
       


def pegar_infos(nome_arquivo, valores):
    with open(f'nfs/{nome_arquivo}', 'rb') as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

    if 'NFe' in dic_arquivo:
        infos_nf = dic_arquivo['NFe']['infNFe']
    else:
        infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']

    numero_nota = infos_nf['@Id']
    empresa_emissora = infos_nf['emit']['xNome']
    nome_cliente = infos_nf['dest']['xNome']
    endereco = infos_nf['dest']['enderDest']

    if 'vol' in infos_nf['transp']:
        peso = infos_nf['transp']['vol']['pesoB']
    else:
        peso = 'não informado'

    valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])


def processar_xmls():
    if not os.path.exists('nfs'):
        messagebox.showerror("Erro", "A pasta 'nfs' não foi encontrada.")
        return

    arquivos = os.listdir('nfs')
    if not arquivos:
        messagebox.showwarning("Aviso", "Nenhum arquivo XML encontrado na pasta 'nfs'.")
        return

    colunas = ['numero_nota', 'empresa_emissora', 'nome_cliente', 'endereco', 'peso']
    valores = []

    for arquivo in arquivos:
        try:
            pegar_infos(arquivo, valores)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar {arquivo}:\n{e}")
            return

    df = pd.DataFrame(columns=colunas, data=valores)
    df.to_excel('NotasFiscais.xlsx', index=False)
    messagebox.showinfo("Sucesso", "Notas processadas e salvas em 'NotasFiscais.xlsx'")


# Interface
janela = tk.Tk()
janela.title("Leitor de Notas Fiscais XML")
janela.geometry("300x180")

titulo = tk.Label(janela, text="Processador de Notas Fiscais", font=("Arial", 14))
titulo.pack(pady=10)

botao = tk.Button(janela, text="Processar XMLs", command=processar_xmls, bg="#4CAF50", fg="white", height=2, width=20)
botao.pack(pady=10)

botao_pasta = tk.Button(janela, text='Selecionar Pasta', command=escolher_pasta)
botao_pasta.pack(pady=5) 


janela.mainloop()

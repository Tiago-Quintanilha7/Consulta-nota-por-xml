import tkinter as tk
from tkinter import messagebox, filedialog
import xmltodict
import os
import pandas as pd

# Variável global para armazenar a pasta selecionada
pasta_selecionada = None

def escolher_pasta():
    global pasta_selecionada
    pasta = filedialog.askdirectory()
    if pasta:
        pasta_selecionada = pasta
        messagebox.showinfo("Pasta Selecionada", f"Pasta selecionada:\n{pasta}")
    else:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada.")


def pegar_infos(caminho_arquivo, valores):
    with open(caminho_arquivo, 'rb') as arquivo_xml:
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
    if not pasta_selecionada:
        messagebox.showerror("Erro", "Nenhuma pasta foi selecionada.")
        return

    arquivos = [f for f in os.listdir(pasta_selecionada) if f.endswith('.xml')]
    if not arquivos:
        messagebox.showwarning("Aviso", "Nenhum arquivo XML encontrado na pasta selecionada.")
        return

    colunas = ['numero_nota', 'empresa_emissora', 'nome_cliente', 'endereco', 'peso']
    valores = []

    for arquivo in arquivos:
        caminho_completo = os.path.join(pasta_selecionada, arquivo)
        try:
            pegar_infos(caminho_completo, valores)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar {arquivo}:\n{e}")
            return

    df = pd.DataFrame(columns=colunas, data=valores)
    caminho_saida = os.path.join(pasta_selecionada, 'NotasFiscais.xlsx')
    df.to_excel(caminho_saida, index=False)
    messagebox.showinfo("Sucesso", f"Notas processadas e salvas em:\n{caminho_saida}")


# Interface
janela = tk.Tk()
janela.title("Leitor de Notas Fiscais XML")
janela.geometry("320x200")

titulo = tk.Label(janela, text="Processador de Notas Fiscais", font=("Arial", 14))
titulo.pack(pady=10)

botao_pasta = tk.Button(janela, text='Selecionar Pasta', command=escolher_pasta, bg="#2196F3", fg="white", height=2, width=20)
botao_pasta.pack(pady=5)

botao = tk.Button(janela, text="Processar XMLs", command=processar_xmls, bg="#4CAF50", fg="white", height=2, width=20)
botao.pack(pady=10)

janela.mainloop()

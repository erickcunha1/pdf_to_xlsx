import fitz  # PyMuPDF
import pandas as pd
import os
import shutil
import sys


def substituir_extensao(nome_arquivo, nova_extensao, complemento=""):
    root, extensao = os.path.splitext(nome_arquivo)
    novo_nome = f"{root}{complemento}.{nova_extensao}"
    return novo_nome

def extrair_texto(caminho_do_pdf):
    dados = {'Página': [], 'mes_ref': [], 'matricula': [], 'endereco_1': [], 'endereco_2': [], 'consumo': [], 'vencimento': [], 'valor_total': [], 'tipo': []}
    documento_pdf = fitz.open(caminho_do_pdf)

    for numero_pagina in range(documento_pdf.page_count):
        pagina = documento_pdf[numero_pagina]

        texto = pagina.get_text()
        dados['Página'].append(numero_pagina + 1)
        dados['mes_ref'].append(retornar_item_da_nota(texto, "Rot. Leitura", 5))
        dados['matricula'].append(retornar_item_da_nota(texto, "Rot. Leitura", 6)) 
        dados['endereco_1'].append(retornar_item_da_nota(texto, "Rot. Leitura", 27)) 
        dados['endereco_2'].append(retornar_item_da_nota(texto, "Rot. Leitura", 28)) 
        dados['consumo'].append(retornar_item_da_nota(texto, "CONSUMO ÁGUA", 1)) 
        dados['vencimento'].append(retornar_item_da_nota(texto, "INFORMAÇÕES DE", 10)) 
        dados['valor_total'].append(retornar_item_da_nota(texto, "INFORMAÇÕES DE", 11)) 
        dados['tipo'].append('AGUA') 

    documento_pdf.close()

    return pd.DataFrame(dados)

def processar_arquivos(origem, destino, pasta_excel):
    arquivos = os.listdir(origem)
    dados_combinados = pd.DataFrame()

    for arquivo in arquivos:
        caminho_completo_origem = os.path.join(origem, arquivo)

        if os.path.isfile(caminho_completo_origem):
            df_temp = extrair_texto(caminho_completo_origem)
            dados_combinados = pd.concat([dados_combinados, df_temp], ignore_index=True)

            caminho_completo_destino = os.path.join(destino, arquivo)

            shutil.move(caminho_completo_origem, caminho_completo_destino)
            print(f'Arquivo movido: {arquivo}')

    caminho_excel_existente = os.path.join(pasta_excel, 'dados_combinados.xlsx')
    if os.path.exists(caminho_excel_existente):
        dados_existente = pd.read_excel(caminho_excel_existente)
        dados_combinados = pd.concat([dados_existente, dados_combinados], ignore_index=True)

    if not dados_combinados.empty:
        nome_excel = 'dados_combinados.xlsx'
        caminho_completo_excel = os.path.join(pasta_excel, nome_excel)
        dados_combinados.to_excel(caminho_completo_excel, index=False)
        print(f'Dados combinados salvos em: {caminho_completo_excel}')

def retornar_item_da_nota(texto, ponto_inicial, qtd_linhas):
    linhas = texto.split('\n')

    encontrou = False
    marco_zero = 0

    for numero_linha, linha in enumerate(linhas):
        if ponto_inicial.lower() in linha.lower() and encontrou == False:
            encontrou = True

        if encontrou:
            marco_zero += 1
            if marco_zero == qtd_linhas:
                return linha

if __name__ == '__main__':
    pasta_origem = os.path.join(os.getcwd(), "PDF's Agua")
    pasta_destino = os.path.join(os.getcwd(), "PDF's Concluidos")
    pasta_excel = os.path.join(os.getcwd(), 'Excel Processado')
    processar_arquivos(pasta_origem, pasta_destino, pasta_excel)
import fitz
import pandas as pd
import os
import shutil
import sys
import re


def substituir_extensao(nome_arquivo, nova_extensao, complemento="") -> str:
    root, extensao = os.path.splitext(nome_arquivo)
    novo_nome = f"{root}{complemento}.{nova_extensao}"
    return novo_nome

def extrair_texto(caminho_do_pdf):
    dados = {
        'Página': [], 
        'Mes Referencia': [], 
        'Matricula': [], 
        'Endereço': [],
        'Municipio/Bairro': [], 
        'CEP': [], 
        'Consumo': [], 
        'Vencimento': [], 
        'Valor Total': [], 
        'Tipo': []
    }
    documento_pdf = fitz.open(caminho_do_pdf)

    for numero_pagina in range(documento_pdf.page_count):
        pagina = documento_pdf[numero_pagina]

        texto = pagina.get_text()
        if 'CONSUMO FATURADO' in texto:
            dados['Página'].append(numero_pagina + 1)
            dados['Mes Referencia'].append(retornar_item_da_nota(texto, "REF:", 2))
            dados['Matricula'].append(retornar_item_da_nota(texto, " NOME DO CLIENTE:", 11)) 
            dados['Endereço'].append(retornar_item_da_nota(texto, " NOME DO CLIENTE:", 5)) 
            dados['Municipio/Bairro'].append(retornar_item_da_nota(texto, " NOME DO CLIENTE:", 7))
            dados['CEP'].append(numero_encontrado)
            dados['Consumo'].append(retornar_item_da_nota(texto, "Consumo-TE", 3)) 
            dados['Vencimento'].append(retornar_item_da_nota(texto, "VENCIMENTO", 2)) 
            dados['Valor Total'].append(retornar_item_da_nota(texto, "TOTAL A PAGAR R$", 2)) 
            dados['Tipo'].append('LUZ')

    documento_pdf.close()
    return pd.DataFrame(dados)

def processar_arquivos_luz(origem, destino, pasta_excel):
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
        dados_existentes = pd.read_excel(caminho_excel_existente)
        dados_combinados = pd.concat([dados_existentes, dados_combinados], ignore_index=True)

    if not dados_combinados.empty:
        nome_excel = 'dados_combinados.xlsx'
        caminho_completo_excel = os.path.join(pasta_excel, nome_excel)
        dados_combinados.to_excel(caminho_completo_excel, index=False)
        print(f'Dados combinados salvos em: {caminho_completo_excel}')

def retornar_item_da_nota(texto, ponto_inicial, qtd_linhas):
    linhas = texto.split('\n')
    encontrou = False
    marco_zero = 0
    cep_pattern = r'(\d{5}-\d{3})'
    
    global numero_encontrado
    numero_encontrado = None

    for numero_linha, linha in enumerate(linhas):
        correspondencia = re.sub(cep_pattern, '', linha)
        cep = re.search(cep_pattern, linha)

        if cep:
            numero_encontrado = cep.group()

        if ponto_inicial.lower() in linha.lower() and encontrou == False:
            encontrou = True

        if encontrou:
            marco_zero += 1
            if marco_zero == qtd_linhas:
                if linha == 'CÓDIGO DO CLIENTE':
                    linha = linhas[numero_linha + 1]
                elif correspondencia:
                    return correspondencia
                return linha


if __name__ == '__main__':
    pasta_origem = os.path.join(os.getcwd(), "PDF's Luz")
    pasta_destino = os.path.join(os.getcwd(), "PDF's Concluidos")
    pasta_excel = os.path.join(os.getcwd(), 'Excel Processado')
    processar_arquivos_luz(pasta_origem, pasta_destino, pasta_excel)
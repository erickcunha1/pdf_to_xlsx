import fitz  # PyMuPDF
import pandas as pd
import os
import shutil
import sys

def substituir_extensao(nome_arquivo, nova_extensao, complemento=""):
    root, extensao = os.path.splitext(nome_arquivo)
    novo_nome = f"{root}{complemento}.{nova_extensao}"
    return novo_nome

def extrair_texto(caminho_do_pdf, pasta_excel, arquivo):
    dados = {'Página': [], 'mes_ref': [], 'matricula': [], 'endereco_1': [], 'endereco_2': [], 'consumo': [], 'vencimento': [], 'valor_total': [], 'tipo': []}
    documento_pdf = fitz.open(caminho_do_pdf)

    for numero_pagina in range(documento_pdf.page_count):
        pagina = documento_pdf[numero_pagina]
        
        texto = pagina.get_text()
        
        if 'CONSUMO FATURADO' in texto:
            dados['Página'].append(numero_pagina + 1)
            dados['mes_ref'].append(retornar_item_da_nota(texto, "REF:", 2))
            dados['matricula'].append(retornar_item_da_nota(texto, " NOME DO CLIENTE:", 11)) 
            dados['endereco_1'].append(retornar_item_da_nota(texto, " NOME DO CLIENTE:", 5)) 
            dados['endereco_2'].append(retornar_item_da_nota(texto, " NOME DO CLIENTE:", 7)) 
            dados['consumo'].append(retornar_item_da_nota(texto, "Consumo-TE", 3)) 
            dados['vencimento'].append(retornar_item_da_nota(texto, "VENCIMENTO", 2)) 
            dados['valor_total'].append(retornar_item_da_nota(texto, "TOTAL A PAGAR R$", 2)) 
            dados['tipo'].append('LUZ')
        else:
            continue

    documento_pdf.close()

    df = pd.DataFrame(dados)

    nome_excel = substituir_extensao(arquivo, 'xlsx')
    caminho_completo_excel = os.path.join(pasta_excel, nome_excel)
    df.to_excel(caminho_completo_excel, index=False)
    print(f'Dados salvos em {caminho_completo_excel}')

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
                if linha == 'CÓDIGO DO CLIENTE':
                    linha = linhas[numero_linha + 1]
                return linha

def processar_arquivos_luz(origem, destino, pasta_excel):
    arquivos = os.listdir(origem)

    for arquivo in arquivos:
        caminho_completo_origem = os.path.join(origem, arquivo)

        if os.path.isfile(caminho_completo_origem):
            extrair_texto(caminho_completo_origem, pasta_excel, arquivo)
            
            caminho_completo_destino = os.path.join(destino, arquivo)

            shutil.move(caminho_completo_origem, caminho_completo_destino)
            print(f'Arquivo movido: {arquivo}')

pasta_origem = os.path.join(os.getcwd(), "PDF's Luz")
pasta_destino = os.path.join(os.getcwd(), "PDF's Concluidos")
pasta_excel = os.path.join(os.getcwd(), 'Excel Processado')

if __name__ == '__main__':
    processar_arquivos_luz(pasta_origem, pasta_destino, pasta_excel)

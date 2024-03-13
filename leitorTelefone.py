import fitz
import pandas as pd
import os
import shutil


def substituir_extensao(nome_arquivo, nova_extensao, complemento="") -> str:
    root, extensao = os.path.splitext(nome_arquivo)
    novo_nome = f"{root}{complemento}.{nova_extensao}"
    return novo_nome

def extrair_texto(caminho_do_pdf):
    dados = {
        'Página': [], 
        'Mes Referencia': [],
        'Contrato': [], 
        'Endereço': [],
        'Localidade': [], 
        'Meio de acesso (tel)': [], 
        'Data Emissão': [], 
        'Valor Faturado': [],
        'Valor Imposto': [], 
        'Tipo': []
    }
    documento_pdf = fitz.open(caminho_do_pdf)
    
    pagina_contrato = documento_pdf[0]

    texto_contrato = pagina_contrato.get_text()

    numero_contrato = retornar_item_apos_palavra_chave(texto_contrato, 'NÚMERO DO CONTRATO', 1)
    endereco = retornar_item_da_nota(texto_contrato, 'DATA DE EMISSÃO', 3)
    data_emissao = retornar_item_da_nota(texto_contrato, 'NÚMERO DO CONTRATO', 1)
    mes_ref = retornar_item_da_nota(texto_contrato, 'PÁGINA', 1)

    for numero_pagina in range(documento_pdf.page_count):
        pagina = documento_pdf[numero_pagina]

        texto = pagina.get_text()
        
        locais = retornar_item_da_nota(texto, 'TOTAL ICMS', 8)
        numeros = retornar_item_da_nota(texto, 'TOTAL ICMS', 7)
        valores_faturados = retornar_item_da_nota(texto, 'TOTAL ICMS', 4)
        valor_imposto = retornar_item_da_nota(texto, 'TOTAL ICMS', 6)
        
        for valor in zip(locais, numeros, valores_faturados, valor_imposto):
            if valor[0] not in ['VALOR ICMS', 'CLIENTE: DEFENSORIA PUBLICA DO ESTADO DA BAHIA']:
                dados['Página'].append(numero_pagina + 1)
                dados['Localidade'].append(valor[0])
                dados['Meio de acesso (tel)'].append(valor[1])
                dados['Valor Faturado'].append(valor[2])
                dados['Valor Imposto'].append(valor[3])
                dados['Contrato'].append(numero_contrato[0])
                dados['Endereço'].append(endereco[0])
                dados['Data Emissão'].append(data_emissao[0])
                dados['Mes Referencia'].append(mes_ref[0])
                dados['Tipo'].append('TELEFONE')

            else:
                continue

    documento_pdf.close()
    return pd.DataFrame(dados)

def processar_arquivos_telefone(origem, destino, pasta_excel):
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

def retornar_item_da_nota(texto, palavra_chave, distancia):
    linhas = texto.strip().split('\n')
    
    # Inicializar lista para armazenar os itens encontrados
    itens_encontrados = []
    
    # Iterar sobre as linhas do texto
    for i, linha in enumerate(linhas):
        # Verificar se a linha contém a palavra-chave
        if palavra_chave in linha:
            # Calcular o índice do item anterior baseado na distância
            indice_item = i - distancia
            # Verificar se o índice está dentro dos limites
            if 0 <= indice_item < len(linhas):
                # Adicionar o item encontrado à lista de itens
                itens_encontrados.append(linhas[indice_item])
    
    # Retornar a lista de itens encontrados
    return itens_encontrados

def retornar_item_apos_palavra_chave(texto, palavra_chave, distancia):
    linhas = texto.strip().split('\n')
    
    # Inicializar lista para armazenar os itens encontrados
    itens_encontrados = []
    
    # Iterar sobre as linhas do texto
    for i, linha in enumerate(linhas):
        # Verificar se a linha contém a palavra-chave
        if palavra_chave in linha:
            # Calcular o índice do item seguinte baseado na distância
            indice_item = i + distancia
            # Verificar se o índice está dentro dos limites
            if 0 <= indice_item < len(linhas):
                # Adicionar o item encontrado à lista de itens
                itens_encontrados.append(linhas[indice_item])
    
    # Retornar a lista de itens encontrados
    return itens_encontrados


if __name__ == '__main__':
    pasta_origem = os.path.join(os.getcwd(), "PDF telefone")
    pasta_destino = os.path.join(os.getcwd(), "PDF Concluidos")
    pasta_excel = os.path.join(os.getcwd(), 'Excel Processado')
    processar_arquivos_telefone(pasta_origem, pasta_destino, pasta_excel)
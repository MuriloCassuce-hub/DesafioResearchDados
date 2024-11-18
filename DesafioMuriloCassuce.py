import os
import datetime
import requests
import pandas as pd
import matplotlib.pyplot as plt


#Baixar arquivos do site.
def baixar_arquivo(dia, mes, ano):
    meses = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    mes_str = meses[mes]
    ano_str = str(ano)[-2:]
    
    #url que realiza o download.
    url = f"https://www.anbima.com.br/informacoes/merc-sec-debentures/arqs/d{ano_str}{mes_str}{dia:02d}.xls"
    #Nome adicionado conforme solicitado.
    nome_arquivo = f"{ano}{mes:02d}{dia:02d}.xls"
    caminho_arquivo = os.path.join("Daily Prices", nome_arquivo)
    
    #Verificar se o arquivo já existe.
    if os.path.exists(caminho_arquivo):
        print(f"Arquivo {nome_arquivo} encontrado no diretório atual.")
        return nome_arquivo
    
    #Tenta realizar o request.
    try:
        resposta = requests.get(url)
        #Caso consiga realizar o request.
        if resposta.status_code == 200:
            pasta = "Daily Prices"
            if not os.path.exists(pasta):
                os.makedirs(pasta)
            
            nome_arquivo = f"{ano}{mes:02d}{dia:02d}.xls"
            caminho_arquivo = os.path.join(pasta, nome_arquivo)
            
            with open(caminho_arquivo, 'wb') as f:
                f.write(resposta.content)
            
            print(f"Arquivo {nome_arquivo} baixado com sucesso!")
            return nome_arquivo
        
        #Caso contrário.
        else:
            print(f"Arquivo {ano_str}{mes_str}{dia:02d}.xls não encontrado no site")
            return None
    #Caso não consiga realizar o request.
    except Exception as e:
        print(f"Erro ao tentar baixar o arquivo: {e}")
        return None

#Baixar os últimos 5 arquivos a partir da data atual do objeto date.
def baixar_ultimos_5_arquivos():
    hoje = datetime.datetime.now().date()
    arquivos_baixados = []
    
    dia_atual = hoje
    while len(arquivos_baixados) < 5:
        nome_arquivo = baixar_arquivo(dia_atual.day, dia_atual.month, dia_atual.year)
        if nome_arquivo:
            arquivos_baixados.append(nome_arquivo)
        dia_atual -= datetime.timedelta(days=1)
    
    return arquivos_baixados

#Lê os arquivos baixados e combina os dados.
def processar_arquivos(arquivos):
    paginas = ['DI_PERCENTUAL', 'DI_SPREAD', 'IPCA_SPREAD']
    indexadores = ['% do DI', 'DI +', 'IPCA +']
    
    arquivos_comb = []
    
    for arquivo in arquivos:
        # Extrai a data do nome do arquivo e formata para aaaa/mm/dd.
        data_arquivo = f"{arquivo[:4]}/{arquivo[4:6]}/{arquivo[6:8]}"
        caminho_arquivo = os.path.join("Daily Prices", arquivo)
        
        for i in range(len(paginas)):
            pagina = paginas[i]
            indexador = indexadores[i]
            
            """
            Lê a partir da coluna ideal para extrair os dados (lê a 8 e ignora a 9) e
            garante que ele consiga ser extraído para o Power Bi
            """
            try:
                arquivo = pd.read_excel(caminho_arquivo, sheet_name=pagina, header=7)
                arquivo = arquivo.iloc[1:].reset_index(drop=True)
                colunas_filtradas = ['Código', 'Nome', 'Taxa de Compra', 'Taxa de Venda', 'Taxa Indicativa','PU']
                arquivo = arquivo[colunas_filtradas]
                #Substitui todo -- ou N/D por 0.
                arquivo.replace('--', 0, inplace=True)
                arquivo.replace('N/D', 0, inplace=True)
                arquivo = arquivo[arquivo['Nome'].notna()]
                arquivo['Indexador'] = indexador
                arquivo['Data'] = data_arquivo #Adiciono a coluna data para facilitar na hora de plotar o gráfico, como forma de consultar.
                arquivos_comb.append(arquivo)
            except Exception as e:
                print(f"Erro ao ler {pagina} no arquivo {arquivo}: {e}")
    
    #Concateno todos os datasets em um só.
    if arquivos_comb:
        arquivo_finalizado = pd.concat(arquivos_comb, ignore_index=True)
        return arquivo_finalizado
    return pd.DataFrame()

#Chama a função para baixar os arquivos.
arquivos_baixados = baixar_ultimos_5_arquivos()
if arquivos_baixados:
    #Assim que baixa, processa eles conforme as instruções.
    dados_combinados = processar_arquivos(arquivos_baixados)
else:
    print("Nenhum arquivo foi baixado.")

#Salva o dataset no formato de excel.
dados_combinados.to_excel('DataSetDailyPrices.xlsx', index=False)

#Criei uma ordem dos indexadores para organizar.
ordem_indexadores = ['% do DI', 'DI +', 'IPCA +']
dados_combinados['Indexador'] = pd.Categorical(dados_combinados['Indexador'], categories=ordem_indexadores, ordered=True)
#Pego a média da taxa indicativa.
data_indexador_media = dados_combinados.groupby(['Data', 'Indexador'])['Taxa Indicativa'].mean().reset_index()
#Ordeno conforme o indexador, para facilitar no plot.
data_indexador_media = data_indexador_media.sort_values(by=['Indexador', 'Data'])


#Plotagem do gráfico conforme a data
fig, axs = plt.subplots(3, 1, figsize=(10, 15))

#Ajusta o hspae para melhorar a visualização
plt.subplots_adjust(hspace=0.7)

#Cria um gráfico para cada indexador
for i, indexador in enumerate(ordem_indexadores):
    dados_filtrados = data_indexador_media[data_indexador_media['Indexador'] == indexador]
    axs[i].plot(dados_filtrados['Data'], dados_filtrados['Taxa Indicativa'], marker='o', label=indexador)
    axs[i].set_title(f'Média da Taxa Indicativa - {indexador}')
    axs[i].set_xlabel('Data')
    axs[i].set_ylabel('Taxa Indicativa')
    axs[i].grid(True)
plt.show()






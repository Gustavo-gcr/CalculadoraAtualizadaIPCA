import sys
import pandas as pd
from datetime import datetime,timedelta
import requests
import calendar
import os
import time

# Função para preencher a coluna 'dia' com datas de 01/06/2016 até 10 anos após a data final
def preencher_coluna_dia(mes_fim, ano_fim):
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        print("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    # Calcula a data inicial e final
    data_inicial = datetime(2016, 6, 30)
    data_final = datetime(ano_fim + 10, mes_fim, 1) - timedelta(days=1)
    
    # Gera todas as datas entre data_inicial e data_final
    todas_as_datas = pd.date_range(start=data_inicial, end=data_final, freq='D')
    
    # Carrega a planilha
    planilha = pd.read_excel(caminho_arquivo, engine='openpyxl')
    
    # Ajusta o número de datas para coincidir com o número de linhas da planilha
    num_linhas = len(planilha)
    if len(todas_as_datas) > num_linhas:
        todas_as_datas = todas_as_datas[:num_linhas]
    elif len(todas_as_datas) < num_linhas:
        todas_as_datas = todas_as_datas.append(pd.Series([''] * (num_linhas - len(todas_as_datas))))

    # Substitui a coluna 'dia' com as novas datas
    planilha['dia'] = todas_as_datas.strftime('%d/%m/%Y')
    
    # Salva as mudanças na planilha
    planilha.to_excel(caminho_arquivo, index=False)
    print("Coluna 'dia' preenchida com todas as datas de 30/06/2016 até 10 anos após o mês final informado.")


# Função para buscar o IPCA do Banco Central
def buscar_ipca(mes, ano, tentativas=3, timeout=10):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.16121/dados?formato=json&dataInicial=01/{mes:02d}/{ano}&dataFinal={ultimo_dia:02d}/{mes:02d}/{ano}"

    for tentativa in range(tentativas):
        try:
            resposta = requests.get(url, timeout=timeout)
            resposta.raise_for_status()
            dados = resposta.json()
            if len(dados) > 0:
                valor_ipca = dados[0]['valor']
                return float(valor_ipca) / 100
            else:
                print(f"IPCA para {mes}/{ano} não encontrado.")
                return None
        except requests.exceptions.RequestException as err:
            print(f"Tentativa {tentativa + 1} falhou: {err}")
            if tentativa < tentativas - 1:
                print("Repetindo tentativa em 5 segundos...")
                time.sleep(5)
            else:
                print(f"Erro ao consultar o Banco Central após {tentativas} tentativas.")
                return None
    return None

def limpar_colunas():
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        print("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    # Adiciona o argumento engine='openpyxl' para abrir arquivos .xlsx
    planilha = pd.read_excel(caminho_arquivo, engine='openpyxl')
    planilha['TAXA DIA'] = ''
    planilha['taxa 100'] = ''
    planilha['TotalPorcentagem'] = ''
    planilha.to_excel(caminho_arquivo, index=False)
    print("Colunas 'TAXA DIA', 'taxa 100' e 'TotalPorcentagem' foram limpas com sucesso.")


def preencher_planilha_ipca(ipca_mensal, mes, ano):
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        print("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    planilha = pd.read_excel(caminho_arquivo, dtype=str, engine='openpyxl')
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    ipca_diario = ipca_mensal / dias_no_mes
    taxa_100 = 1 + ipca_diario

    ipca_diario_formatado = f"{ipca_diario:.3%}".replace(".", ",")
    taxa_100_formatada = f"{taxa_100:.3%}".replace(".", ",")

    for i, data in planilha['dia'].items():
        # A data já está no formato '%d/%m/%Y', então vamos separar
        dia, mes_data, ano_data = data.split('/')
        
        # Verifica se o mês e o ano correspondem
        if int(mes_data) == mes and int(ano_data) == ano:
            planilha.at[i, 'TAXA DIA'] = str(ipca_diario_formatado)
            planilha.at[i, 'taxa 100'] = str(taxa_100_formatada)

    planilha.to_excel(caminho_arquivo, index=False)
    print(f"Planilha preenchida com sucesso para {mes}/{ano}.")
    
# Função para preencher todas as taxas de julho de 2016 até a data fornecida
def preencher_intervalo_ipca(mes_fim, ano_fim):
    mes_atual = mes_fim
    ano_atual = ano_fim
    ipca_anterior = None

    for ano in range(2016, ano_fim + 1):
        for mes in range(1, 13):
            if ano == 2016 and mes < 7:
                continue  # Ignorar meses antes de julho de 2016
            if ano == ano_fim and mes > mes_fim:
                break  # Parar ao atingir o mês final fornecido

            ipca_mensal = buscar_ipca(mes, ano)
            if ipca_mensal is None:
                if ipca_anterior is not None:
                    print(f"IPCA para {mes}/{ano} não encontrado. Usando IPCA do mês anterior: {ipca_anterior:.4f}")
                    ipca_mensal = ipca_anterior
                else:
                    print(f"Não foi possível encontrar o IPCA para {mes}/{ano} e não há mês anterior para usar.")
                    continue

            ipca_anterior = ipca_mensal
            preencher_planilha_ipca(ipca_mensal, mes, ano)

def calcular_total_porcentagem():
    caminho_arquivo = 'IPCA-Teste.xlsx'
    if not os.path.exists(caminho_arquivo):
        print("Erro: Arquivo IPCA-Teste.xlsx não encontrado.")
        return

    planilha = pd.read_excel(caminho_arquivo, dtype=str)

    # Inicializa o valor acumulado como 1 para começar a multiplicação
    valor_acumulado = 1.0

    # Itera de forma decrescente sobre os dias mais recentes para os mais antigos
    for i in reversed(range(len(planilha))):
        if pd.notnull(planilha.at[i, 'taxa 100']):
            taxa_100 = float(planilha.at[i, 'taxa 100'].replace(',', '.').replace('%', '')) / 100
            valor_acumulado *= taxa_100  # Multiplica o valor acumulado pela taxa 100
            total_porcentagem = valor_acumulado * 100  # Converte para percentual
            planilha.at[i, 'TotalPorcentagem'] = f"{total_porcentagem:.2f}%".replace('.', ',')  # Formata como percentual

    planilha.to_excel(caminho_arquivo, index=False)
    print("Coluna 'TotalPorcentagem' preenchida com sucesso.")
    
    
# Função principal
def main():
    if len(sys.argv) != 3:
        print("Erro: Você deve fornecer o mês e ano como argumentos.")
        return
    
    mes_fim = int(sys.argv[1])
    ano_fim = int(sys.argv[2])

    # Primeiro, preencher a coluna 'dia' com todas as datas necessárias
    preencher_coluna_dia(mes_fim, ano_fim)

    # Em seguida, limpar as colunas
    limpar_colunas()
    
    # Preencher todas as colunas 'TAXA DIA' e 'taxa 100' desde julho de 2016 até a data fornecida
    preencher_intervalo_ipca(mes_fim, ano_fim)

    # Após todas as taxas estarem preenchidas, calcular a coluna 'TotalPorcentagem'
    calcular_total_porcentagem()

if __name__ == "__main__":
    main()
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os

# Função para carregar os dados do Excel a partir da mesma pasta que o código Python
def carregar_dados_excel(nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.dirname(__file__), nome_arquivo)
        df = pd.read_excel(caminho_arquivo)
        return df
    except Exception as e:
        st.write(f"Erro ao carregar a planilha: {e}")
        return None

# Função para calcular o valor ajustado pelo juros
def calcular_valor_ajustado(valor_inicial, taxa_porcentagem):
    return valor_inicial * (taxa_porcentagem / 100)

# Interface do Streamlit
st.title("Calculadora de Juros IPCA")

# Carregar os dados do Excel
dados_ipca = carregar_dados_excel('IPCA-Teste.xlsx')

if dados_ipca is not None:
    # Input do usuário para o valor
    valor_input = st.number_input("Digite o valor", min_value=0.0)

    # Data mínima permitida (01/07/2016) e data máxima como o último dia do mês anterior
    hoje = datetime.today()
    data_minima = datetime(2016, 7, 1)  # Data do último registro
    data_maxima = hoje.replace(day=1) - timedelta(days=1)  # Último dia do mês anterior

    # Input de data com restrição de qualquer mês anterior até 01/07/2016
    data_input = st.date_input("Selecione a data (Dia, Mês, Ano)", 
                               min_value=data_minima, max_value=data_maxima)

    if data_input and valor_input > 0:
        # Converter a data selecionada para string no formato dia/mês/ano
        data_selecionada_str = data_input.strftime('%d/%m/%Y')
   
        # Verificar se as colunas corretas estão no dataframe
        if 'dia' in dados_ipca.columns and 'TotalPorcentagem' in dados_ipca.columns:
            # Filtrar o dataframe para encontrar a data correspondente
            linha_dados = dados_ipca[dados_ipca['dia'] == data_selecionada_str]
        else:
            st.write("Colunas 'dia' ou 'TotalPorcentagem' não encontradas. Verifique o nome exato das colunas no arquivo Excel.")
            st.stop()

        if not linha_dados.empty:
            # Pega a taxa de juros correspondente à data selecionada
            taxa_porcentagem = float(linha_dados['TotalPorcentagem'].values[0].replace('%', '').replace(',', '.'))
            st.write(f"Taxa de juros para dia {data_selecionada_str}: {taxa_porcentagem:.2f}%")
            # Calcular o valor ajustado com base na taxa de juros
            valor_ajustado = calcular_valor_ajustado(valor_input, taxa_porcentagem)
            st.write(f"Valor ajustado: R$ {valor_ajustado:.2f}")
        else:
            st.write(f"Não há dados disponíveis para a data {data_input}.")
else:
    st.write("Erro ao carregar os dados do Excel.")

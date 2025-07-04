import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Configuração da página do Streamlit
st.set_page_config(layout="wide")
st.title('Relatório Mensal TomTicket')

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type="xlsx")

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    
    # Leitura da aba "Worksheet" pulando as 5 primeiras linhas
    worksheet_df = pd.read_excel(xls, sheet_name='Worksheet', header=5)
    
    required_columns = ['Categoria', 'Atendente', 'Origem do Chamado', 'Última Situação']
    if all(col in worksheet_df.columns for col in required_columns):
        # Criar abas
        tab1, tab2, tab3, tab_detalhamento, tab4, tab_origem = st.tabs([
            "Análise por Categoria",  
            "Análise por Atendente",
            "Painel do Atendente",
            "Detalhamento por Atendente",
            "Situação",
            "Origem do Chamado"
        ])
        
        with tab1:
            category_counts = worksheet_df['Categoria'].value_counts()
            col1, col2 = st.columns(2)
            col1.subheader('Contagem Chamados por Categorias')
            col1.write(category_counts)

            col2.subheader('Gráfico de Categorias')
            if not category_counts.empty:
                fig, ax = plt.subplots()
                bars = category_counts.plot(kind='bar', ax=ax)
                ax.set_xlabel('Categoria')
                ax.set_ylabel('Contagem')
                for bar in bars.patches:
                    ax.annotate(format(bar.get_height(), '.0f'), 
                                (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                                ha='center', va='center', 
                                size=10, xytext=(0, 8), 
                                textcoords='offset points')
                col2.pyplot(fig)
            else:
                col2.info('Nenhuma categoria encontrada.')

        with tab2:
            attendant_counts = worksheet_df['Atendente'].value_counts()
            col3, col4 = st.columns(2)
            col3.subheader('Contagem de Chamados por cada Atendente')
            col3.write(attendant_counts)
            col4.subheader('Gráfico de Atendentes')
            if not attendant_counts.empty:
                fig, ax = plt.subplots()
                bars = attendant_counts.plot(kind='bar', ax=ax)
                ax.set_xlabel('Atendente')
                ax.set_ylabel('Contagem')
                for bar in bars.patches:
                    ax.annotate(format(bar.get_height(), '.0f'), 
                                (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                                ha='center', va='center', 
                                size=10, xytext=(0, 8), 
                                textcoords='offset points')
                col4.pyplot(fig)
            else:
                col4.info('Nenhum chamado por atendente encontrado.')

        with tab3:
            painel_df = worksheet_df[
                (worksheet_df['Atendente Criador'] == worksheet_df['Atendente']) & 
                (worksheet_df['Origem do Chamado'] == 'Painel do Atendente')
            ]
            painel_attendant_counts = painel_df['Atendente'].value_counts()
            col1, col2 = st.columns(2)
            col1.subheader('Contagem de Chamados Criados por cada Atendente')
            col1.write(painel_attendant_counts)

            col2.subheader('Gráfico de Atendentes')
            if not painel_attendant_counts.empty:
                fig, ax = plt.subplots()
                bars = painel_attendant_counts.plot(kind='bar', ax=ax, color='skyblue', edgecolor='black')
                ax.set_xlabel('Atendente')
                ax.set_ylabel('Contagem')
                ax.set_title('Quantidade de chamados')
                for bar in bars.patches:
                    ax.annotate(format(bar.get_height(), '.0f'),
                                (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                                ha='center', va='bottom',
                                size=10, xytext=(0, 5),
                                textcoords='offset points')
                col2.pyplot(fig)
            else:
                col2.info('Nenhum chamado encontrado para os filtros selecionados.')

        with tab_detalhamento:
            st.subheader("Detalhamento de cada Chamado Criado por cada Atendente")
            painel_df = worksheet_df[
                (worksheet_df['Atendente Criador'] == worksheet_df['Atendente']) & 
                (worksheet_df['Origem do Chamado'] == 'Painel do Atendente')
            ]
            attendants_options = painel_df['Atendente'].unique()
            if len(attendants_options) > 0:
                selected_attendant = st.selectbox("Selecione o Atendente", attendants_options)
                df_attendant = painel_df[painel_df['Atendente'] == selected_attendant]
                category_counts = df_attendant['Categoria'].value_counts()
                col1, col2 = st.columns(2)
                col1.subheader(f"Contagem de Categorias ({selected_attendant})")
                col1.write(category_counts)
                col2.subheader(f"Gráfico de Categorias ({selected_attendant})")
                if not category_counts.empty:
                    fig, ax = plt.subplots()
                    bars = category_counts.plot(kind='bar', ax=ax, color='skyblue', edgecolor='black')
                    ax.set_xlabel("Categoria")
                    ax.set_ylabel("Contagem")
                    ax.set_title(f"Chamados de {selected_attendant}")
                    for bar in bars.patches:
                        ax.annotate(format(bar.get_height(), '.0f'),
                                    (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                                    ha='center', va='bottom',
                                    size=10, xytext=(0, 5),
                                    textcoords='offset points')
                    col2.pyplot(fig)
                else:
                    col2.info('Nenhuma categoria encontrada para este atendente.')
            else:
                st.info('Nenhum atendente encontrado para o filtro.')

        with tab4:
            situacao_counts = worksheet_df['Última Situação'].value_counts()
            col7, col8 = st.columns(2)
            col7.subheader('Contagem por Situação dos Chamados')
            col7.write(situacao_counts)
            col8.subheader('Gráfico de Última Situação')
            if not situacao_counts.empty:
                fig, ax = plt.subplots()
                bars = situacao_counts.plot(kind='bar', ax=ax)
                ax.set_xlabel('Última Situação')
                ax.set_ylabel('Contagem')
                for bar in bars.patches:
                    ax.annotate(format(bar.get_height(), '.0f'), 
                                (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                                ha='center', va='center', 
                                size=10, xytext=(0, 8), 
                                textcoords='offset points')
                col8.pyplot(fig)
            else:
                col8.info('Nenhuma situação encontrada.')

        with tab_origem:
            origem_counts = worksheet_df['Origem do Chamado'].value_counts()
            col9, col10 = st.columns(2)
            col9.subheader('Contagem por Origem do Chamado')
            col9.write(origem_counts)
            col10.subheader('Gráfico de Origem do Chamado')
            if not origem_counts.empty:
                fig, ax = plt.subplots()
                bars = origem_counts.plot(kind='bar', ax=ax, color='lightgreen', edgecolor='black')
                ax.set_xlabel('Origem do Chamado')
                ax.set_ylabel('Contagem')
                ax.set_title('Distribuição por Origem do Chamado')
                for bar in bars.patches:
                    ax.annotate(format(bar.get_height(), '.0f'),
                                (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                                ha='center', va='bottom',
                                size=10, xytext=(0, 5),
                                textcoords='offset points')
                col10.pyplot(fig)
            else:
                col10.info('Nenhuma origem encontrada.')

else:
    st.warning('Por favor, faça o upload de um arquivo Excel.')

import streamlit as st
import pandas as pd
import time
import operacao
import csv
import datetime
import os
import warnings
warnings.filterwarnings("ignore")
# Defina a variável de ambiente PYTHONUTF8 para o encoding padrão UTF-8
os.environ["PYTHONUTF8"] = "1"

st.set_page_config(
    page_title="Gerador de Conceito",
    page_icon=":chart_with_upwards_trend:",  # Ícone opcional
    layout="centered"  # Layout opcional, pode ser "wide" ou "centered"
)

st.title('Layout de Importação Conceito')

# Crie um expander com um rótulo
with st.expander("IMPOSTOS"):
    uploaded_file = st.file_uploader("Carregue o Arquivo:",key="impostos")
    if uploaded_file is not None:
        try:
            df_opcoes = pd.read_excel(uploaded_file,dtype={'Centro de Custo': str})
            st.write('Dados Carregados com Sucesso!!')
        except Exception as e:
            st.error(f"Erro ao carregar os dados!!")
        time.sleep(1)
        opcao_sheet = st.selectbox('Selecione o Conceito:',('1º Conceito', '2º Conceito','3º Conceito'))
        st.subheader('Definições de Cabeçalho do Arquivo')
        number = st.number_input('Código do Lote', value = 2023, key='int1')
        title = st.text_input('Título do Lançamento', 'Importacao de lancamento contabil manual',key='titulo')
        select_data = st.date_input("Data de Inclusão:")
        # Verifique se a data foi selecionada
        if select_data:
            data_formatada = select_data.strftime("%d/%m/%Y")
        cabecalho = ['M', number, title, data_formatada]
        st.text(f"Cabeçalho: {cabecalho}")
        if st.button("Transformar Dados"):
            try:
                # Mostra o spinner enquanto os dados estão sendo carregados e processados
                with st.spinner("Processando os dados..."):
                    time.sleep(3)
                    df_filtro = pd.read_excel(uploaded_file, dtype={'Centro de Custo': str}, sheet_name=f'{opcao_sheet}')
                    zip_file_path = operacao.transforma_dados_e_cria_zip(df_filtro, cabecalho)
            except Exception as e:
                st.error(f"Erro ao carregar os dados: {e}")
            else:
                # Use a função st.download_button para permitir o download do arquivo ZIP
                with open(zip_file_path, 'rb') as f:
                    zip_data = f.read()
                st.download_button(
                    label="Clique para baixar o arquivo ZIP",
                    data=zip_data,
                    key="dados_download.zip",
                    file_name=f"Import_{opcao_sheet}.zip",
                    mime="application/zip"
                )
    else:
        pass

st.divider()

with st.expander("LAYOUT GERAL"):
    uploaded_file2 = st.file_uploader("Carregue o Arquivo:",key="geral")

    if uploaded_file2 is not None:
        try:
            df_opcoes = pd.read_excel(uploaded_file2,dtype={'Centro de Custo': str})
            st.write('Dados Carregados com Sucesso!!')
        except Exception as e:
            st.error(f"Erro ao carregar os dados!!")
        time.sleep(1)
        opcoes = sorted(list(df_opcoes['Coligada'].unique()))
        opcao_coligada = st.selectbox('Selecione a Coligada',(opcoes))
        st.subheader('Definições de Cabeçalho do Arquivo')
        number = st.number_input('Código do Lote', value = 2023, key='int2')
        title2 = st.text_input('Título do Lançamento', 'Importacao de lancamento contabil manual',key='titulo2')
        select_data4 = st.date_input("Data de Inclusão:", key='data2')
        # Verifique se a data foi selecionada
        if select_data4:
            data_formatada = select_data4.strftime("%d/%m/%Y")
        cabecalho2 = ['M', number, title2, data_formatada]
        st.text(f"Coligada Selecionada: {opcao_coligada}")
        st.text(f"Cabeçalho: {cabecalho2}")
        try:
            df_filtro = pd.read_excel(uploaded_file2,dtype={'Centro de Custo': str})
            df_filtro = df_filtro.query(f'Coligada == {opcao_coligada}')
            df = operacao.transforma_dados(df_filtro,cabecalho2)
        except Exception as e:
            st.error(f"Erro ao carregar os dados: {e}")
    else:
        pass
    if uploaded_file2 is not None:
        if st.button('Gerar Layout de Importação') and df is not None:
            with st.spinner('Transformação de dados em Progresso...'):
                try:
                    time.sleep(3)
                    st.success('Feito! O Layout está pronto para Download.')
                except Exception as e:
                    st.error(f"Erro durante a transformação dos dados: {e}")
            if df is not None:
                st.download_button(
                    label="Download data as CSV",
                    data=open(df, 'rb').read(),
                    file_name=f'import_Colig_{opcao_coligada}.csv',
                    mime='text/csv',
                )

st.divider()

with st.expander("RATEIOS"):
    uploaded_file3 = st.file_uploader("Carregue o Arquivo:", key="rateio")

    if uploaded_file3 is not None:
        try:
            df_rateio = pd.read_excel(uploaded_file3, sheet_name='NOTAS CRÉD-DÉB')
            st.write('Dados Carregados com Sucesso!!')
        except Exception as e:
            st.error(f"Erro ao carregar os dados!!")
        time.sleep(1)
        select_data2 = st.date_input("Data de Emissão:")
        if select_data2:
            data_emissao = select_data2.strftime("%d/%m/%Y")
            data_emissao = datetime.datetime.strptime(data_emissao, "%d/%m/%Y")
        select_data3 = st.date_input("Data de Vencimento:")
        if select_data3:
            data_vencimento = select_data3.strftime("%d/%m/%Y")
            data_vencimento = datetime.datetime.strptime(data_vencimento, "%d/%m/%Y")
        select_tipo = st.selectbox('Selecione o Conceito:',('Contas a Pagar', 'Contas a Receber'))
        if select_tipo:
            dataselect_tipo_vencimento = str(select_tipo)

        if st.button("Transformar Dados",key='Download2'):
            try:
                with st.spinner("Processando os dados..."):
                    time.sleep(3)
                    df_rateio = pd.read_excel(uploaded_file3, sheet_name='NOTAS CRÉD-DÉB')
                    df_final = operacao.transforma_rateio(df_rateio)
                    zip_rateio = operacao.cria_zip_rateio(df_final, str(data_emissao), str(data_vencimento),select_tipo)                
            except Exception as e:
                st.error(f"Erro ao carregar os dados: {e}")
            else:   
                if zip_rateio is not None:
                    with open(zip_rateio, 'rb') as f:
                        zip_data2 = f.read()
                    st.download_button(
                        label="Baixar Arquivo ZIP",
                        data=zip_data2,
                        key="rateio_download.zip",
                        file_name=f"Rateio_{select_tipo}.zip",
                        mime="application/zip"
                    )
    else:
        pass


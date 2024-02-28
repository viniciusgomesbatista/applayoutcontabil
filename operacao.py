import csv
import zipfile
import os
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

def transforma_dados_e_cria_zip(df,cabecalho):
    df = df.dropna()
    df['Partida'] = '*P'
    df['Centro de Custo'] = df['Centro de Custo'].astype('str')
    df['Nº Documento'] = ''
    df['Conta de Contra Partida'] = ''
    df['Coligada'] = df['Coligada'].astype('int')
    df['Filial'] = df['Filial'].astype('int')
    df['Código do Histórico'] = df['Código do Histórico'].astype('int')
    df['Código do Histórico'] = df['Código do Histórico'].astype('string')
    df['Código do Histórico'] = df['Código do Histórico'].astype(str).str.zfill(3)
    # Use o método replace para substituir os pontos por vírgulas
    df['Valor'] = df['Valor'].map('{:.2f}'.format).str.replace('.', ',')

    df = df[['Coligada','Partida','Nº Documento','Conta de Débito','Conta de Crédito','Conta de Contra Partida','Valor','Código do Histórico','Complemento do Histórico','Filial','Centro de Custo']]
    
    coligadas = sorted(list(df['Coligada'].unique()))
    
    # Criar um arquivo ZIP para salvar os CSVs
    with zipfile.ZipFile('dados.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
        for x in coligadas:
            df_filtrado = df[df['Coligada'] == x].copy()  # Crie uma cópia do DataFrame filtrado
            df_filtrado.drop(columns=['Coligada'], inplace=True)
            dados = []
            for indice, linha in df_filtrado.iterrows():
                lista_valores = linha.tolist()
                dados.append(lista_valores)

            # Nome do arquivo CSV que vamos salvar
            nome_do_arquivo = f'Col{x}' + '_import.csv'

            # Escreva o cabeçalho no arquivo CSV
            with open(nome_do_arquivo, 'w', newline='') as arquivo_csv:
                escritor_csv = csv.writer(arquivo_csv, delimiter=';')
                escritor_csv.writerow(cabecalho)

            # Abrir o arquivo novamente para adicionar os dados
            with open(nome_do_arquivo, 'a', newline='') as arquivo_csv:
                escritor_csv = csv.writer(arquivo_csv, delimiter=';')
                escritor_csv.writerows(dados)
            
            # Adicione o arquivo CSV ao arquivo ZIP
            zipf.write(nome_do_arquivo, arcname=nome_do_arquivo)
            
    return 'dados.zip'  # Retorna o nome do arquivo ZIP gerado

def transforma_dados(df, cabecalho):
    df['Partida'] = '*P'
    df['Nº Documento'] = ''
    df['Conta de Contra Partida'] = ''
    df['Centro de Custo'] = df['Centro de Custo'].astype('string')
    df['Código do Histórico'] = df['Código do Histórico'].astype('int')
    df['Código do Histórico'] = df['Código do Histórico'].astype('string')
    df['Código do Histórico'] = df['Código do Histórico'].astype(str).str.zfill(3)
    df = df.dropna()
    # Use o método replace para substituir os pontos por vírgulas
    df['Valor'] = df['Valor'].map('{:.2f}'.format).str.replace('.', ',')
    df = df[['Partida','Nº Documento','Conta de Débito','Conta de Crédito','Conta de Contra Partida','Valor','Código do Histórico','Complemento do Histórico','Filial','Centro de Custo']]

    dados = []
    for indice, linha in df.iterrows():
        lista_valores = linha.tolist()
        dados.append(lista_valores)

     # Nome do arquivo CSV que vamos salvar
    nome_do_arquivo = 'importacao.csv'

    # Escreva o cabeçalho no arquivo CSV
    with open(nome_do_arquivo, 'w', newline='') as arquivo_csv:
        escritor_csv = csv.writer(arquivo_csv, delimiter=';')
        escritor_csv.writerow(cabecalho)

    # Abrir o arquivo novamente para adicionar os dados
    with open(nome_do_arquivo, 'a', newline='') as arquivo_csv:
        escritor_csv = csv.writer(arquivo_csv, delimiter=';')
        escritor_csv.writerows(dados)
    return nome_do_arquivo

def padronizar_centro_custo(centro_custo):
        return centro_custo.zfill(5) 
def adicionar_ponto_na_segunda_posicao(centro_custo):
    if centro_custo[2] != '.':
        return centro_custo[:2] + '.' + centro_custo[2:]
    else:
        return centro_custo

def transforma_rateio(df):
    palavra_chave = 'CONTABILIZAÇÃO'
    resultado_filtrado = df[df['Unnamed: 1'].str.contains(palavra_chave, case=False, na=False)]
    resultado_filtrado2 = df[df['Unnamed: 13'].str.contains(palavra_chave, case=False, na=False)]
    # Ajuste Titulo de Lançamento
    dic1 = resultado_filtrado['Unnamed: 1'].to_dict()
    dic2 = resultado_filtrado2['Unnamed: 13'].to_dict()
    novo_dicionario = {}
    indices = range(1, 19)  # Números de 1 a 18

    for indice, chave in zip(indices, dic1.keys()):
        novo_dicionario[indice] = dic1[chave]
    # Ajuste Titulo de Lançamento
    novo_dicionario2 = {}
    indices = range(1, 19)  # Números de 1 a 18

    for indice, chave in zip(indices, dic2.keys()):
        novo_dicionario2[indice] = dic2[chave]
    df = df.iloc[2:]
    # Definir a primeira linha como cabeçalho de coluna
    df.columns = df.iloc[0]
    # Redefinir o índice
    df = df.reset_index(drop=True)
    # Remover a primeira linha (agora usada como cabeçalho)
    df = df.iloc[1:]
    titulo = "TITULO"
    df.columns.values[0] = titulo
    df['TITULO'] = df['TITULO'].fillna(method="ffill")
    df['TITULO'] = df['TITULO'].astype('int')
    df['TITULO2'] = df['TITULO']
    df['TITULO'] = df['TITULO'].map(novo_dicionario)
    df['TITULO2'] = df['TITULO2'].map(novo_dicionario2)
    df1 = df.iloc[:, :12]
    novos_nomes = ['TITULO', 'DATA', 'EMPRESA1', 'FILIAL1', 'CONTA DÉBITO', 'CRESP1',
       'EMPRESA2', 'FILIAL2', 'CONTA CRÉDITO', 'CRESP2', 'VALOR', 'HISTÓRICO']
    df1.columns = novos_nomes
    filtro = df1['EMPRESA1'].isin(['RRPM','EDITORA'])
    df1 = df1[filtro]
    df1 = df1.query("VALOR != 0")
    df1 = df1.dropna()

    # Padronização centro de custo
    df1['CRESP1'] = df1['CRESP1'].astype('string').apply(padronizar_centro_custo)
    df1['CRESP1'] = df1['CRESP1'].astype('string').apply(adicionar_ponto_na_segunda_posicao)
    df1['CRESP2'] = df1['CRESP2'].astype('string').apply(padronizar_centro_custo)
    df1['CRESP2'] = df1['CRESP2'].astype('string').apply(adicionar_ponto_na_segunda_posicao)
    df2 = df.iloc[:, 13:]
    novos_nomes2 = ['DATA', 'EMPRESA1', 'FILIAL1', 'CONTA DÉBITO', 'CRESP1',
       'EMPRESA2', 'FILIAL2','CONTA CRÉDITO', 'CRESP2', 'VALOR', 'HISTÓRICO','TITULO']
    df2.columns = novos_nomes2
    filtro2 = df2['CONTA CRÉDITO'].isin(['2.1.1.11.0002'])
    df2 = df2[filtro2]
    df2 = df2.query("VALOR != 0")
    df2 = df2.dropna()
    # Padronização de centro de custo
    df2['CRESP1'] = df2['CRESP1'].astype('string').apply(padronizar_centro_custo)
    df2['CRESP1'] = df2['CRESP1'].astype('string').apply(adicionar_ponto_na_segunda_posicao)
    df2['CRESP2'] = df2['CRESP2'].astype('string').apply(padronizar_centro_custo)
    df2['CRESP2'] = df2['CRESP2'].astype('string').apply(adicionar_ponto_na_segunda_posicao)

    #Junção Final
    df_final = pd.concat([df1,df2])
    df_final['VALOR'] = pd.to_numeric(df_final['VALOR'], errors='coerce')
    df_final = df_final[pd.notna(df_final['VALOR'])]
    dicionario_substituicao = {'02.213..':'02.213',"02.217..":"02.217",'02.213.':'02.213',
                               "02.217.":"02.217", "1..108":"01.108","'1..102":"01.102"}
    df_final['CRESP1'] = df_final['CRESP1'].replace(dicionario_substituicao)
    df_final['CRESP2'] = df_final['CRESP1']
    #Carrega os dados das empresas
    df_empresas = pd.read_excel('empresa.xlsx', dtype={'CODIGO CLIENTE': str,'COLIGADA':str,'CODFILIAL':str,'CNPJ':str})
    # Carrega os dados com o codigo do cliente
    #df_codigo_cliente = pd.read_excel('cod_cliente.xlsx',dtype={'CODIGO CLIENTE': str,'CONTA_BANCO':str})
    df_codigo_cliente = pd.read_csv('cod_cliente.csv', sep=";", dtype={'COD_CLIENTE':str,'CENTRO_CUSTO':str}, encoding='UTF-8')
    # Mapeamento para trazer cnpj, cod_cliente, coligada e cod da filial.
    df_final['IDENTIFICADOR_01'] =  df_final['EMPRESA1'] + df_final['FILIAL1']
    df_final['IDENTIFICADOR_02'] =  df_final['EMPRESA2'] + df_final['FILIAL2']
    dic_coligada = dict(zip(df_empresas['IDENTIFICADOR'], df_empresas['COLIGADA']))
    df_final['COLIGADA1'] = df_final['IDENTIFICADOR_01'].map(dic_coligada)
    df_final['COLIGADA2'] = df_final['IDENTIFICADOR_02'].map(dic_coligada)
    dic_filial = dict(zip(df_empresas['IDENTIFICADOR'], df_empresas['CODFILIAL']))
    df_final['COD_FILIAL1'] = df_final['IDENTIFICADOR_01'].map(dic_filial)
    df_final['COD_FILIAL2'] = df_final['IDENTIFICADOR_02'].map(dic_filial)
    dic_cnpj = dict(zip(df_empresas['IDENTIFICADOR'], df_empresas['CNPJ']))
    df_final['CNPJ1'] = df_final['IDENTIFICADOR_01'].map(dic_cnpj)
    df_final['CNPJ2'] = df_final['IDENTIFICADOR_02'].map(dic_cnpj)
    dic_cod_cliente = dict(zip(df_codigo_cliente['TITULO'], df_codigo_cliente['COD_CLIENTE']))
    dic_conta_banco = dict(zip(df_codigo_cliente['TITULO'], df_codigo_cliente['CONTA_BANCO']))
    dic_tipo_rateio = dict(zip(df_codigo_cliente['TITULO'], df_codigo_cliente['RECEBER_PAGAR']))
    dic_centro_custo = dict(zip(df_codigo_cliente['TITULO'], df_codigo_cliente['CENTRO_CUSTO']))
    df_final['COD_CLIENTE'] = df_final['TITULO'].map(dic_cod_cliente)
    df_final['CONTA_BANCO'] = df_final['TITULO'].map(dic_conta_banco)
    df_final['RECEBER_PAGAR'] = df_final['TITULO'].map(dic_tipo_rateio)
    df_final['CENTRO_CUSTO'] = df_final['TITULO'].map(dic_centro_custo)
    return df_final

def cria_zip_rateio(df_final, data_emissao, data_vencimento,select_tipo):
    # Nome do arquivo ZIP de saída
    nome_arquivo_zip = 'rateios.zip'
    empresas = df_final['IDENTIFICADOR_01'].unique()

    if select_tipo == "Contas a Receber":
        df_final[df_final['RECEBER_PAGAR'] == 'Receber']
        # Criar um arquivo ZIP
        with zipfile.ZipFile(nome_arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for emp in empresas:
                df_base = df_final[df_final['IDENTIFICADOR_01'] == emp]
                nome_coligada = df_base['EMPRESA1'].iat[0]
                nome_filial = df_base['FILIAL1'].iat[0]
                nome_arquivo = f"{nome_coligada}_{nome_filial}.txt"
                lanc_rateios = df_base['TITULO'].unique()

                # Inicializar uma string para armazenar os dados
                dados_formatados = ''

                for x in lanc_rateios:
                    # filtro no dataframe
                    df_provisorio = df_base[df_base['TITULO'] == x]

                    # Parametros
                    cnpj = df_provisorio['CNPJ1'].iat[0]
                    codigo_cliente = df_provisorio['COD_CLIENTE'].iat[0]
                    cod_filial = df_provisorio['COD_FILIAL1'].iat[0]
                    cod_banco = df_provisorio['CONTA_BANCO'].iat[0]
                    centro_custo = df_provisorio['CENTRO_CUSTO'].iat[0]

                    # Agrupamento
                    df_rateio = df_provisorio.groupby(['CRESP1', 'CRESP2']).sum().reset_index()
                    df_rateio['CRESP2'] = 'U' + df_rateio['CRESP1'].astype(str)
                    df_rateio['CRESP1'] = 'B' + df_rateio['CRESP1'].astype(str)
                    df_rateio['HISTÓRICO'] = 'H' + df_provisorio['HISTÓRICO'].iat[0]
                    df_rateio['DATA'] = df_provisorio['DATA'].iat[0]

                    # Formatação de valores
                    soma_valor = df_rateio['VALOR'].sum()
                    soma_formatada = f'{soma_valor:.2f}'.replace('.', ',')

                    df = df_rateio[df_rateio['VALOR'] != 0]
                    df['VALOR'] = df['VALOR'].apply(lambda x: f'{x:.2f}'.replace('.', ','))

                    # Inicialização das variáveis
                    cresp1 = ''
                    cresp2 = ''
                    historico = ''
                    linha_final = ''

                    # Configuração dos Valores cresp1 e cresp2
                    for index, row in df.iterrows():
                        cresp1 += f"{row['CRESP1']:<30}{'1':<21}{row['VALOR']:<70}\n"
                        cresp2 += f"{row['CRESP2']:<26}{row['VALOR']:<345}\n"

                    historico = f"{row['HISTÓRICO']:<256}\n"
                    linha_final = f"]{soma_formatada:<20}{'0,00':<20}{'0,00':<20}\n"

                    # Configuração da Linha I
                    li0 = 'I80.00000.12072'
                    li1 = '1'
                    li2 = '1,00'
                    li3 = soma_formatada
                    li4 = '0,00'
                    li5 = '0,00'
                    li6 = '0,00'
                    li7 = f'{data_emissao}'
                    li8 = '8.01'
                    li9 = 'UN'
                    li10 = '1,00'
                    li11 = '0,00'
                    li12 = f'{centro_custo}'
                    li13 = '0,00'
                    li14 = '2.11.001'
                    li15 = '0,00'
                    li16 = '0,00'
                    li17 = 'N0,00'
                    li18 = '0,00'
                    li19 = '1001'
                    li20 = soma_formatada
                    linha = f"{li0:<35}{li1:<11}{li2:<20}{li3:<20}{li4:<40}{li5:<40}{li6:<20}{li7:<145}{li8:<65}{li9:<5}{li10:<70}{li11:<20}{li12:<125}{li13:<148}{li14:<10}{li15:<20}{li16:<20}{li17:<31}{li18:<44}{li19:<16}{li20:<56}\n"

                    # Configuração do Cabeçalho
                    arg0 = 'M'
                    arg1 = cod_filial
                    arg2 = '001'
                    arg3 = f'{codigo_cliente:010}'
                    arg4 = '999903'
                    arg5 = 'A'
                    arg6 = '2.1.70'
                    arg7 = f'AA000{data_emissao}{data_emissao}'
                    arg45 = f'{data_vencimento}'
                    arg8 = '0,00'
                    arg9 = '001'
                    arg10 = soma_formatada
                    arg11 = soma_formatada
                    arg12 = soma_formatada
                    arg13 = '0,00'
                    arg14 = '0,00'
                    arg15 = '0,00'
                    arg16 = '0,00'
                    arg17 = '0,00'
                    arg18 = '0,00'
                    arg19 = '0,00'
                    arg20 = '1.02.18.001'
                    arg21 = '8.01'
                    arg22 = 'R$'
                    arg23 = f'{data_emissao}'
                    arg24 = '1'
                    arg25 = '0'
                    arg26 = '1'
                    arg27 = 'leandro.souza'
                    arg28 = '0,00'
                    arg29 = f'{centro_custo}'
                    arg30 = f'{cod_banco}'
                    arg31 = '0'
                    arg32 = '651698'
                    arg33 = f'{data_emissao}'
                    arg34 = cnpj
                    arg35 = f'{data_emissao}'
                    arg36 = '153025'
                    arg37 = f'{data_emissao}'
                    arg38 = '0'
                    arg39 = '1'
                    arg40 = '0'
                    arg41 = '0,00'
                    arg42 = '0,00'
                    arg43 = '0,00'
                    arg44 = '0,00'
                    cabecalho = f"{arg0:<5}{arg1:<31}{arg2:<45}{arg3:<50}{arg4:<35}{arg5:<8}{arg6:<10}{arg7:<35}{arg45:<25}{arg8:<40}{arg9:<10}{arg10:<20}{arg11:<20}{arg12:<160}{arg13:<20}{arg14:<20}{arg15:<20}{arg16:<116}{arg17:<57}{arg18:<20}{arg19:<115}{arg20:<75}{arg21:<350}{arg22:<20}{arg23:<14}{arg24:<5}{arg25:<5}{arg26:<17}{arg27:<98}{arg28:<96}{arg29:<29}{arg30:<45}{arg31:<10}{arg32:<36}{arg33:<65}{arg34:<20}{arg35:<64}{arg36:<155}{arg37:<75}{arg38:<5}{arg39:<5}{arg40:<195}{arg41:<21}{arg42:<20}{arg43:<20}{arg44:<20}\n"

                    # Adicionar os dados formatados à string
                    dados_formatados += cabecalho + linha + cresp1 + cresp2 + historico + linha_final

                # Adicionar a string de dados diretamente ao arquivo ZIP
                zipf.writestr(nome_arquivo, dados_formatados)
        
        return nome_arquivo_zip
    else:
        df_final[df_final['RECEBER_PAGAR'] == 'Pagar']
        # Criar um arquivo ZIP
        with zipfile.ZipFile(nome_arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for emp in empresas:
                df_base = df_final[df_final['IDENTIFICADOR_01'] == emp]
                nome_coligada = df_base['EMPRESA1'].iat[0]
                nome_filial = df_base['FILIAL1'].iat[0]
                nome_arquivo = f"{nome_coligada}_{nome_filial}.txt"
                lanc_rateios = df_base['TITULO'].unique()

                # Inicializar uma string para armazenar os dados
                dados_formatados = ''

                for x in lanc_rateios:
                    # filtro no dataframe
                    df_provisorio = df_base[df_base['TITULO'] == x]

                    # Parametros
                    cnpj = df_provisorio['CNPJ1'].iat[0]
                    codigo_cliente = df_provisorio['COD_CLIENTE'].iat[0]
                    cod_filial = df_provisorio['COD_FILIAL1'].iat[0]
                    cod_banco = df_provisorio['CONTA_BANCO'].iat[0]
                    centro_custo = df_provisorio['CENTRO_CUSTO'].iat[0]

                    # Agrupamento
                    df_rateio = df_provisorio.groupby(['CRESP1', 'CRESP2']).sum().reset_index()
                    df_rateio['CRESP2'] = 'U' + df_rateio['CRESP1'].astype(str)
                    df_rateio['CRESP1'] = 'B' + df_rateio['CRESP1'].astype(str)
                    df_rateio['HISTÓRICO'] = 'H' + df_provisorio['HISTÓRICO'].iat[0]
                    df_rateio['DATA'] = df_provisorio['DATA'].iat[0]

                    # Formatação de valores
                    soma_valor = df_rateio['VALOR'].sum()
                    soma_formatada = f'{soma_valor:.2f}'.replace('.', ',')

                    df = df_rateio[df_rateio['VALOR'] != 0]
                    df['VALOR'] = df['VALOR'].apply(lambda x: f'{x:.2f}'.replace('.', ','))

                    # Inicialização das variáveis
                    cresp1 = ''
                    cresp2 = ''
                    historico = ''
                    linha_final = ''

                    # Configuração dos Valores cresp1 e cresp2
                    for index, row in df.iterrows():
                        cresp1 += f"{row['CRESP1']:<30}{'1':<21}{row['VALOR']:<70}\n"
                        cresp2 += f"{row['CRESP2']:<26}{row['VALOR']:<345}\n"

                    historico = f"{row['HISTÓRICO']:<256}\n"
                    linha_final = f"]{soma_formatada:<20}{'0,00':<20}{'0,00':<20}\n"

                    # Configuração da Linha I
                    # Valores originais
                    li_cp01 = 'I80.00000.12072'
                    li_cp02 = '1'
                    li_cp03 = '1,00'
                    li_cp04 = soma_formatada
                    li_cp05 = '0,00'
                    li_cp06 = '0,00'
                    li_cp07 = f'{data_emissao}'
                    li_cp08 = '1.02.19.001'
                    li_cp09 = '002'
                    li_cp10 = '8.01'
                    li_cp11 = 'UN'
                    li_cp12 = '1,00'
                    li_cp13 = '0,00'
                    li_cp14 = f'{centro_custo}'
                    li_cp15 = '0,00'
                    li_cp16 = '2.11.001'
                    li_cp17 = '0,00'
                    li_cp18 = '0,00'
                    li_cp19 = 'N0,00'
                    li_cp20 = '0,00'
                    li_cp21 = '1001'
                    li_cp22 = soma_formatada

                    # Linha formatada
                    linha = f"{li_cp01:<35}{li_cp02:<11}{li_cp03:<20}{li_cp04:<20}{li_cp05:<40}{li_cp06:<60}{li_cp07:<70}" \
                    f"{li_cp08:<25}{li_cp09:<50}{li_cp10:<65}{li_cp11:<5}{li_cp12:<70}{li_cp13:<20}{li_cp14:<125}" \
                    f"{li_cp15:<148}{li_cp16:<10}{li_cp17:<20}{li_cp18:<20}{li_cp19:<31}{li_cp20:<44}{li_cp21:<16}{li_cp22:<56}\n"
                    # Configuração do Cabeçalho
                    arg_cp01 = 'M'
                    arg_cp02 = cod_filial
                    arg_cp03 = '001'
                    arg_cp04 = f'{codigo_cliente:010}'
                    arg_cp05 = '00999903'
                    arg_cp06 = 'D'
                    arg_cp07 = '1.2.52'
                    arg_cp08 = f'AA000{data_emissao}{data_emissao}'
                    arg_cp09 = '0,00'
                    arg_cp10 = '001'
                    arg_cp11 = soma_formatada
                    arg_cp12 = soma_formatada
                    arg_cp13 = soma_formatada
                    arg_cp14 = '0,00'
                    arg_cp15 = '0,00'
                    arg_cp16 = '0,00'
                    arg_cp17 = '0,00'
                    arg_cp18 = '0,00'
                    arg_cp19 = '0,00'
                    arg_cp20 = '0,00'
                    arg_cp21 = '0,00'
                    arg_cp22 = '0,00'
                    arg_cp23 = '1.02.19.001'
                    arg_cp24 = '002'
                    arg_cp25 = '8.01'
                    arg_cp26 = 'R$'
                    arg_cp27 = f'{data_emissao}'
                    arg_cp28 = '1'
                    arg_cp29 = '0'
                    arg_cp30 = '1'
                    arg_cp31 = 'leandro.souza'
                    arg_cp32 = '0,00'
                    arg_cp33 = f'{centro_custo}'
                    arg_cp34 = f'{cod_banco}'
                    arg_cp35 = '1'
                    arg_cp36 = '0'
                    arg_cp37 = '651698'
                    arg_cp38 = f'{data_emissao}'
                    arg_cp39 = cnpj
                    arg_cp40 = f'{data_emissao}'
                    arg_cp41 = '153025'
                    arg_cp42 = f'{data_emissao}'
                    arg_cp43 = f'{data_emissao}'
                    arg_cp44 = '0'
                    arg_cp45 = '0,00'
                    arg_cp46 = '0,00'
                    arg_cp47 = '0,00'
                    arg_cp48 = '0,00'

                    cabecalho = f"{arg_cp01:<5}{arg_cp02:<31}{arg_cp03:<45}{arg_cp04:<50}{arg_cp05:<35}{arg_cp06:<8}{arg_cp07:<10}" \
                                f"{arg_cp08:<60}{arg_cp09:<40}{arg_cp10:<10}{arg_cp11:<20}{arg_cp12:<20}{arg_cp13:<80}" \
                                f"{arg_cp14:<20}{arg_cp15:<60}{arg_cp16:<20}{arg_cp17:<20}{arg_cp18:<20}{arg_cp19:<116}{arg_cp20:<57}" \
                                f"{arg_cp21:<20}{arg_cp22:<115}{arg_cp23:<25}{arg_cp24:<50}{arg_cp25:<350}{arg_cp26:<20}{arg_cp27:<14}" \
                                f"{arg_cp28:<5}{arg_cp29:<5}{arg_cp30:<17}{arg_cp31:<98}{arg_cp32:<96}{arg_cp33:<29}{arg_cp34:<15}" \
                                f"{arg_cp35:<30}{arg_cp36:<10}{arg_cp37:<36}{arg_cp38:<65}{arg_cp39:<20}{arg_cp40:<64}{arg_cp41:<155}" \
                                f"{arg_cp42:<40}{arg_cp43:<35}{arg_cp44:<205}{arg_cp45:<21}{arg_cp46:<20}{arg_cp47:<20}{arg_cp48:<20}\n"

                    # Adicionar os dados formatados à string
                    dados_formatados += cabecalho + linha + cresp1 + cresp2 + historico + linha_final

                # Adicionar a string de dados diretamente ao arquivo ZIP
                zipf.writestr(nome_arquivo, dados_formatados)
        
        return nome_arquivo_zip
 

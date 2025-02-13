import re
import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import os
import calendar 

# Funções de tratamento de data
def tratar_data_credito(data):
    if 'TRI' in data:
        trimestres = {
            '1º TRI': ('01-01', '03-31'),
            '2º TRI': ('04-01', '06-30'),
            '3º TRI': ('07-01', '09-30'),
            '4º TRI': ('10-01', '12-31')
        }
        for tri, (inicio, fim) in trimestres.items():
            if tri in data:
                ano = data.split('/')[-1]
                data_inicial = f'{ano}-{inicio}'
                data_final = f'{ano}-{fim}'
                return data_inicial, data_final

    elif any(x in data for x in ['ANUAL', 'Anual']):
        try:
            ano = data.split('/')[-1]
            data_inicial = f'{ano}-01-01'
            data_final = f'{ano}-12-31'
            return data_inicial, data_final
        except Exception as e:
            return None, None

    else:
        try:
            data = pd.to_datetime(data, dayfirst=True)
            ano = data.year
            mes = data.month
            primeiro_dia = f'{ano}-{mes:02d}-01'
            ultimo_dia = f'{ano}-{mes:02d}-{calendar.monthrange(ano, mes)[1]:02d}'
            return primeiro_dia, ultimo_dia
        except Exception as e:
            return None, None

def tratar_data_competencia(data):
    try:
        data = data.strip()

        if data.startswith('13/'):
            ano = data.split('/')[1].strip()
            return f'{ano}-12-31'

        elif '/' in data:
            mes, ano = data.split('/')
            mes = mes.strip()
            ano = ano.strip()
            return f'{ano}-{mes.zfill(2)}-01'

        elif data.replace('.0', '').isdigit() and len(data.replace('.0', '')) == 4:
            ano = data.replace('.0', '')
            return f'{ano}-01-01'

        return pd.to_datetime(data, dayfirst=True, errors='coerce').strftime('%Y-%m-%d')

    except Exception as e:
        return None

# Função principal para extrair informações dos PDFs
def extract_info_from_pages(pdf_document):

    info = {
        'cod_cnpj': None,
        'nome_cliente': None,
        'cod_perdcomp': None,
        'data_transmissao': None,
        'tipo_transacao': None,
        'tipo_credito': None,
        'tipo_perdcomp_retificacao': None,
        'cod_perdcomp_retificacao': None,
        'origem_credito_judicial': None,
        'nome_responsavel_preenchimento': None,
        'cod_cpf_preenchimento': None,
        #'regime_tributario': None,
        'cod_per_origem': None,
        'data_inicial_credito': None,
        'data_final_credito': None,
        'data_competencia':None,
        'valor_credito': None,
        'valor_credito_atualizado': None,
        'selic_acumulada': None,
        'valor_compensado_dcomp': None,
        'valor_credito_data_transmissao': None,
        'valor_saldo_original': None,
        'cod_perdcomp_cancelado': None,  #Eu posso cancelar tanto PER como DCOMP
        'codigos_receita':[], 
        'data_vencimento_tributo':[],
        'valor_principal_tributo': [],
        'valor_multa_tributo': [],
        'valor_juros_tributo': [],
        'valor_total_tributo': []

 }
    
    
    page_patterns = {
        0: {
            'cod_cnpj': r"CNPJ \s*([\d./-]+)",
            'cod_perdcomp': r"CNPJ \s*[\d./-]+\s*([\d.]+-\d+)", 
            'nome_cliente': r"Nome Empresarial\s*([A-Za-z\s]+?(?:LTDA|ME|EIRELI|SA)\b)",
            'data_transmissao': r"Data de Transmissão\s*([\d/]+)",
            'tipo_transacao': r"Tipo de Documento\s*([\w\s]+?)(?=\s*Tipo de Crédito)",  
            'tipo_credito': r"Tipo de Crédito\s*([\w\s]+)(?=\s*PER/DCOMP Retificador)",
            'tipo_perdcomp_retificacao': r"PER/DCOMP Retificador\s*([\w\s]+?)(?=\n|\.|$)",
             'cod_perdcomp_retificacao': r"N[º°] PER/DCOMP Retificado\s*([\d.]+-[\d.]+)",  # Ajustado para capturar o padrão exato
            'origem_credito_judicial': r"Crédito Oriundo de Ação Judicial\s*([\w\s]+?)(?=\n|\.|$)",
            'nome_responsavel_preenchimento': r"Nome\s+([\w\s]+)\s+CPF\s+(\d{3}\.\d{3}\.\d{3}-\d{2})",
            'cod_cpf_preenchimento': r"CPF \s*([\d./-]+)",
            'cod_perdcomp_cancelado': r"Número do PER/DCOMP a Cancelar\s*([\d./-]+)"

        },
        1: {
             'nome_responsavel_preenchimento': r"Nome\s+([\w\s]+)\s+CPF\s+(\d{3}\.\d{3}\.\d{3}-\d{2})",
             'cod_cpf_preenchimento': r"CPF \s*([\d./-]+)"

         },
        2: {
            # Ajuste no padrão de regex
            #'regime_tributario': r"(?:Forma de Tributação no Período |Forma de Tributação do Lucro)\s*([\w\s])",
            'cod_per_origem': r"N[º°] do PER/DCOMP Inicial\s*([\d.]+-[\d.]+)",
            'data_inicial_credito': r"Data Inicial do Período\s*([\d/]+)",
            'data_final_credito': r"Data Final do Período\s*([\d/]+)",
            'valor_credito': r"Valor do Saldo Negativo\s*([\d.,]+)",
            'valor_credito_atualizado': r"Crédito Atualizado\s*([\d.,]+)",
            'valor_saldo_original': r"Saldo do Crédito Original\s*([\d.,]+)",
            'selic_acumulada': r"Selic Acumulada\s*([\d.,]+)",
            'data_competencia': r"(?:1[º°]|2[º°]|3[º°]|4[º°])\s*Trimestre/\d{4}",  # Novo regex para capturar corretamente
            'valor_credito_data_transmissao': r"\s*([\d.,]+)Crédito Original na Data da Entrega"
        }}
    
    codigo_receita_pattern = r"Código da Receita/Denominação\s*(\d{4}-\d{2})"
    data_vencimento_tributo_pattern = r"Data de Vencimento do Tributo/Quota\s*([\d/]+)"
    valor_principal_tributo_pattern = r"Principal\s*([\d,.]+)"
    valor_multa_tributo_pattern = r"Multa\s*([\d.,]+)"
    valor_juros_tributo_pattern = r"Juros\s*([\d.,]+)"
    valor_total_tributo_pattern = r"Total\s*([\d.,]+)"
    valor_compensado_pattern = r"Total do Crédito Original Utilizado nesta DCOMP\s*([\d.,]+)"
    #total_debitos_pattern = r"Total dos débitos deste Documento\s*([\d.,]+)"
    valor_credito_transmissao_pattern = r"([\d.,]+)\sCrédito Original na Data da Entrega"

    for page_num, patterns in page_patterns.items():
        if page_num < pdf_document.page_count:
            page = pdf_document[page_num]
            page_text = page.get_text()
            #print(page_text)

            # Ajuste para 'data_competencia' se tipo_transacao for 'Pedido de Ressarcimento'
            if info.get('tipo_transacao') == 'Pedido de Ressarcimento' and page_num == 2:
                    ano_match = re.search(r"Ano\s*(\d{4})", page_text)
                    trimestre_match = re.search(r"(\d{1,2}[º])\s*Trimestre", page_text)

                    if ano_match and trimestre_match:
                        ano = ano_match.group(1)
                        trimestre = trimestre_match.group(1)
                        info['data_competencia'] = f'{trimestre}/{ano}'
                    else:
                        info['data_competencia'] = "---"
            
            for key, pattern in patterns.items():
                #print(patterns.items())
                matches = re.findall(pattern, page_text)
                if matches:
                    if key == 'nome_responsavel_preenchimento' and len(matches) > 1:
                        info['nome_responsavel_preenchimento'] = matches[1][0].strip()
                        info['cod_cpf_preenchimento'] = matches[1][1].strip()
                    elif key not in ['nome_responsavel_preenchimento', 'cod_cpf_preenchimento']:
                        info[key] = matches[0].strip()


            # Capturar valor_compensado_dcomp
            match_compensado = re.search(valor_compensado_pattern, page_text)
            if match_compensado:
                info['valor_compensado_dcomp'] = match_compensado.group(1)
                info['valor_compensado_dcomp'] = info['valor_compensado_dcomp'].replace('.', '').replace(',', '.') 

            # Capturar valor_credito_data_transmissao
            match_credito_transmissao = re.search(valor_credito_transmissao_pattern, page_text)
            if match_credito_transmissao:
                info['valor_credito_data_transmissao'] = match_credito_transmissao.group(1)
                

            if 'tipo_transacao' in info and info['tipo_transacao']:
                if info['tipo_transacao'] in ['Pedido de Restituição', 'Declaração de Compensação', 'Pedido de Ressarcimento']:
                    tipo_credito_pattern = r"Tipo de Crédito\s*([\w\s\-\/\.]+)(?=\s*PER/DCOMP Retificador)"
                    cod_per_origem_pattern = r"N[º°] do PER/DCOMP Inicial\s*([\d./-]+)"
                elif info['tipo_transacao'] == "Pedido de Cancelamento":
                    tipo_credito_pattern = r"Tipo de Crédito\s*([\w\s]+)(?=\s*Número do PER)"
                    cod_per_origem_pattern = r"Número do PER/DCOMP a Cancelar\s*([\d./-]+)"


                tipo_credito_match = re.search(tipo_credito_pattern, page_text)
                if tipo_credito_match:
                    info['tipo_credito'] = tipo_credito_match.group(1).strip()


                cod_per_origem_match = re.search(cod_per_origem_pattern, page_text)
                if cod_per_origem_match:
                    if info['tipo_transacao'] == "Pedido de Cancelamento":
                        info['cod_perdcomp_cancelado'] = cod_per_origem_match.group(1).strip()
                    else:
                        info['cod_per_origem'] = cod_per_origem_match.group(1).strip()


            # Mapeia os campos e seus respectivos padrões de regex
            patterns = {
                'codigos_receita': codigo_receita_pattern,
                'data_vencimento_tributo': data_vencimento_tributo_pattern,
                'valor_principal_tributo': valor_principal_tributo_pattern,
                'valor_multa_tributo': valor_multa_tributo_pattern,
                'valor_juros_tributo': valor_juros_tributo_pattern,
                'valor_total_tributo': valor_total_tributo_pattern,
            }

            # Itera sobre os padrões e extrai as informações
            for page_num in range(3, pdf_document.page_count):
                page = pdf_document[page_num]
                page_text = page.get_text()

                for key, pattern in patterns.items():
                    # Encontra as correspondências usando o padrão
                    matches = re.findall(pattern, page_text)
                    if matches:
                        if isinstance(info[key], list):  # Certifique-se de que ainda é uma lista
                            info[key].extend(matches)
                            info[key] = list(set(info[key]))

            for key, value in info.items():
                if isinstance(value, list):  # Verificar se o valor ainda é uma lista
                    info[key] = ";".join(value)
                    #info[key] = info[key].replace('.','').replace(',','.')
                    
    return info


def extrair_valor_numerico(texto, formatar_para_exibicao=False):
    """
    Converte o texto para um número float, com opção de retornar como string formatada.
    """
    if texto:
        texto = texto.strip()
        texto = texto.replace('.', '').replace(',', '.')  # Remover separadores de milhares e ajustar decimal
        try:
            valor = float(texto)
            if formatar_para_exibicao:
                return f"{valor:,.2f}".replace('.', ',')  # Formatar para exibição com `,` como separador decimal
            return valor
        except ValueError:
            print(f"[ERRO] Não foi possível converter o texto '{texto}' para float.")
    return 0



def process_pdfs_in_directory(directory_path):
    all_data = []
    for file_name in os.listdir(directory_path):
        if file_name.endswith(".pdf"):
            pdf_path = os.path.join(directory_path, file_name)
            with fitz.open(pdf_path) as pdf_document:
                info = extract_info_from_pages(pdf_document)
                info['Arquivo'] = file_name
                all_data.append(info)

    df = pd.DataFrame(all_data)

    colunas_numericas_texto = ['valor_compensado_dcomp', 'valor_credito_data_transmissao']
    for coluna in colunas_numericas_texto:
        df[coluna] = df[coluna].apply(extrair_valor_numerico)
        #print(df[coluna])

    # Outras colunas numéricas
    colunas1 = ['valor_credito', 'valor_credito_atualizado', 'valor_saldo_original', 'selic_acumulada']
    # for coluna in colunas1:
    #     df[coluna] = df[coluna].fillna('').apply(lambda x: x.replace('.', '').replace(',', '.') if isinstance(x, str) else x)
    #     df[coluna] = pd.to_numeric(df[coluna], errors='coerce')#.fillna(0)

    colunas2 = ['tipo_perdcomp_retificacao', 'cod_perdcomp_retificacao', 'tipo_credito', 'origem_credito_judicial', 'nome_responsavel_preenchimento', 'cod_cpf_preenchimento', 'cod_per_origem', 'cod_perdcomp_cancelado']
    for coluna in colunas2:
        df[coluna] = df[coluna].fillna('---')

    
    colunas3 = ['data_inicial_credito', 'data_final_credito', 'data_transmissao'] #, 
    for coluna in colunas3:
        df[coluna] = pd.to_datetime(df[coluna], format='%d/%m/%Y', errors='coerce')#.fillna('---')

    return df


directory_path = r"G:\Drives compartilhados\Clientes COM Faturamento\DIGITHOBRASIL SOLUCOES EM SOFTWARE LTDA\7-DECLARAÇÕES\PERDCOMP\20250210_PDF"
df_result = process_pdfs_in_directory(directory_path)


st.write("Informações extraídas dos PDFs:")
st.dataframe(df_result)
#print(df_result.dtypes)

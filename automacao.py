import os
from seleniumbase import SB
import pandas as pd
import shutil
from selenium.webdriver.common.keys import Keys

# Carregar e processar dados
arquivo = 'dados_1730724702442.csv'
data = pd.read_csv(arquivo, sep=';')
data.rename(columns={
    'Número do PER/DCOMP': 'cod_perdcomp',
    'Data de Transmissão': 'data_transmissao',
    'Tipo de Crédito': 'tipo_credito',
    'Tipo de Documento': 'tipo_documento',
    'Situação': 'tipo_status'}, inplace=True)

data['cod_perdcomp'] = data['cod_perdcomp'].astype(str)
data['cod_perdcomp_clean'] = data['cod_perdcomp'].str.replace(r'[^\w]', '', regex=True)
data['data_transmissao'] = pd.to_datetime(data['data_transmissao'], dayfirst=True)
data_filtrada = data.loc[data['data_transmissao'] > '2019-01-01'].reset_index()
data_filtrada = data_filtrada.iloc[73:]

# Caminho base
base_path = r"C:\Users\hailleen.gonzalez\Documents\LendoPDF\PDF_Extraidos"
if not os.path.exists(base_path):
    os.makedirs(base_path)

# Inicializar lista para armazenar os resultados
resultados = []

with SB(uc=True, test=True) as sb:
    ecac_site = 'https://cav.receita.fazenda.gov.br'
    login_certificado = '//*[@id="login-dados-certificado"]/p[2]/input'
    select_certificado = '//*[@id="login-certificate"]'
    consulta_perdcomp_pdf = 'https://www3.cav.receita.fazenda.gov.br/perdcomp-web/#/documento/identificacao-novo'

    # Realizar login
    while True:
        try:
            sb.maximize_window()
            sb.activate_cdp_mode(ecac_site)
            sb.sleep(2)
            sb.cdp.click_if_visible(login_certificado)
            sb.sleep(5)
            sb.cdp.mouse_click(select_certificado)
            sb.wait_for_element_visible('//*[@id="btnPerfil"]/span')
            break
        except Exception as e:
            print(f"Erro ao tentar login: {e}")

    sb.cdp.mouse_click('//*[@id="btnPerfil"]/span')
    sb.sleep(0.5)
    cnpj = '06029385000104'
    sb.cdp.send_keys('//*[@id="txtNIPapel2"]', cnpj)
    sb.sleep(0.5)
    sb.cdp.mouse_click('//*[@id="formPJ"]/input[4]')

    while True:
        try:
            sb.open(consulta_perdcomp_pdf)
            sb.sleep(0.5)
            print('Localizando o visualizar documentos')
            sb.wait_for_element_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
            print("Tentando clicar no botão Documentos.")
            sb.click_if_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
            sb.click('//*[@id="myTab"]/li[2]')
            break
        except Exception as e:
            print(e)

    for index, row in data_filtrada.iterrows():
        perdcomp = row['cod_perdcomp']
        perdcomp_clean = row['cod_perdcomp_clean']
        print(f"Processando PERDCOMP {perdcomp} (Índice: {index})")

        cnpj_path = os.path.join(base_path, cnpj)
        if not os.path.exists(cnpj_path):
            os.makedirs(cnpj_path)

        downloads_path = r"C:\Users\hailleen.gonzalez\Documents\LendoPDF\downloaded_files"
        downloaded_file = os.path.join(downloads_path, f"{perdcomp_clean}.pdf")
        destination = os.path.join(cnpj_path, f"{perdcomp_clean}.pdf")

        status_salvo = False

        try:
            while not os.path.exists(destination):
                sb.wait_for_element_visible('//*[@id="numeroPerdcomp"]')
                sb.cdp.mouse_click('//*[@id="numeroPerdcomp"]')
                sb.send_keys('//*[@id="numeroPerdcomp"]', perdcomp)
                sb.sleep(1)

                download_xpaths = [
                    '//*[@id="page-content-wrapper"]/perdcomp-tela-inicial/div/perdcomp-tabs/perdcomp-tab[2]/div/perdcomp-listar-docs-enviados/div/div[1]/simple-collapsible/div/div[2]/div[7]/div[3]/i',
                    '//*[@id="page-content-wrapper"]/perdcomp-tela-inicial/div/perdcomp-tabs/perdcomp-tab[2]/div/perdcomp-listar-docs-enviados/div/div[1]/simple-collapsible/div/div[2]/div[7]/div[2]/i'
                ]

                for xpath in download_xpaths:
                    try:
                        sb.wait_for_element_visible(xpath, timeout=5)
                        sb.click(xpath)
                        sb.sleep(5)

                        if os.path.exists(downloaded_file):
                            if os.path.exists(destination):
                                os.remove(destination)
                            shutil.move(downloaded_file, destination)
                            print(f"PDF salvo em: {destination}")
                            status_salvo = True
                            break
                    except Exception:
                        continue
                break

        except Exception as e:
            print(f"Erro ao processar PERDCOMP {perdcomp}: {e}")

        # Adicionar informações ao resultado
        resultados.append({
            "Índice": index,
            "PERDCOMP Clean": perdcomp_clean,
            "Salvo": status_salvo
        })

    # Criar DataFrame com os resultados e salvar em Excel
    df_resultados = pd.DataFrame(resultados)
    print(df_resultados)
    df_resultados.to_excel(r"C:\Users\hailleen.gonzalez\Documents\LendoPDF\Tabela_resultados.xlsx", index=False)

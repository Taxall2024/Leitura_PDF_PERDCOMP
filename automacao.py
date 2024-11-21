import os
from seleniumbase import SB
from seleniumbase import BaseCase
import time
import pandas as pd
import shutil


from selenium.webdriver.common.keys import Keys
#from dotenv import load_dotenv
#from certificate import CertificadoWindows

# load_dotenv()
# CERT_PASS = os.getenv('CERT_PASS')
# CERT_NAME = os.getenv('CERT_NAME')

arquivo = 'dados_1730724418707.csv'  #, 'dados_1730724418707.csv']
data= pd.read_csv(arquivo, sep=';')
data.rename(columns={'Número do PER/DCOMP': 'cod_perdcomp', 
                   'Data de Transmissão': 'data_transmissao', 
                   'Tipo de Crédito': 'tipo_credito', 
                   'Tipo de Documento': 'tipo_documento', 
                   'Situação': 'tipo_status'}, inplace=True)
data['cod_perdcomp'] = data['cod_perdcomp'].astype(str)    #.str.replace('.', '').str.replace('-', '')
data['cod_perdcomp_clean'] = data['cod_perdcomp'].str.replace(r'[^\w]', '', regex=True)  # Remover caracteres especiais
data['data_transmissao'] = pd.to_datetime(data['data_transmissao'], dayfirst=True)
data_filtrada = data.loc[data['data_transmissao'] > '2019-01-01'].reset_index()
#start_index = 73  # Começar a partir do índice desejado
data_filtrada = data_filtrada.iloc[-5:]

base_path = r"C:\Users\hailleen.gonzalez\Documents\LendoPDF\PDF_Extraidos"
if not os.path.exists(base_path):
    os.makedirs(base_path)

with SB(uc=True, test=True) as sb: # , disable_csp=True
    ecac_site = 'https://cav.receita.fazenda.gov.br' # Site oficial do eCAC
    login_certificado = '//*[@id="login-dados-certificado"]/p[2]/input' # Botão para acessar eCAC com certificado
    select_certificado = '//*[@id="login-certificate"]' # Botão para selecionar certificado
    ecac_homebutton = '//*[@id="linkHome"]'
    consulta_perdcomp_tabela = 'https://www3.cav.receita.fazenda.gov.br/consprocperdcomp/consulta/processamento'
    consulta_perdcomp_pdf = 'https://www3.cav.receita.fazenda.gov.br/perdcomp-web/#/documento/identificacao-novo'


    while True:
        try:    
            sb.maximize_window()
            sb.activate_cdp_mode(ecac_site)
            sb.sleep(2)
            # sb.cdp.clear_cookies()
            #sb.sleep(2)

            sb.cdp.click_if_visible(login_certificado)
            sb.sleep(5)
            sb.cdp.mouse_click(select_certificado)
            # windows = sb.window_handles
            # sb.switch_to.window(windows[1])
            #CW = CertificadoWindows()
            #CW.ChangeCurrentCertificate(CERT_NAME, CERT_PASS, 'https://certificado.sso.acesso.gov.br')
            sb.sleep(5)
            #sb.cdp.mouse_click(ecac_homebutton)
            sb.wait_for_element_visible('//*[@id="btnPerfil"]/span')
            break

        except Exception as e:
            print(f'Erro ao tentar login: {e}')
            # sb.cdp.reload(ignore_cache=True, script_to_evaluate_on_load=None)
        sb.sleep(0.5)

    sb.cdp.mouse_click('//*[@id="btnPerfil"]/span')
    sb.sleep(0.5)
    cnpj = '06926324000131' #, '06926324000131']
    sb.cdp.send_keys('//*[@id="txtNIPapel2"]', cnpj ) #Preciso iterar sobre todos os CNPJ dos clientes press_keys
    sb.sleep(0.5)
    sb.cdp.mouse_click('//*[@id="formPJ"]/input[4]')
    sb.sleep(0.5)
    # sb.cdp.mouse_click('//*[@id="btn263"]/input[4]')

    
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
        try:
            perdcomp = row['cod_perdcomp']
            perdcomp_clean = row['cod_perdcomp_clean']
            print(f"Processando PERDCOMP {perdcomp}")
            print(f"Índice: {index}")

            # Caminho onde o PDF deve ser salvo
            cnpj_path = os.path.join(base_path, cnpj)
            if not os.path.exists(cnpj_path):
                os.makedirs(cnpj_path)
            
            downloads_path = os.path.expanduser(r"C:\Users\hailleen.gonzalez\Documents\LendoPDF\downloaded_files")
            downloaded_file = os.path.join(downloads_path, f"{perdcomp_clean}.pdf")
            destination = os.path.join(cnpj_path, f"{perdcomp_clean}.pdf")

            # Continuar tentando até que o PDF seja salvo
            max_attempts = 3  # Definir o número máximo de tentativas
            attempts = 0

            while not os.path.exists(destination):
                attempts += 1  # Incrementa o contador de tentativas
                print(f"Tentativa {attempts}/{max_attempts} para baixar PERDCOMP {perdcomp}")

                try:
                    # Aguarda e insere o código PERDCOMP no campo de entrada
                    sb.wait_for_element_visible('//*[@id="numeroPerdcomp"]')
                    sb.cdp.mouse_click('//*[@id="numeroPerdcomp"]')
                    sb.send_keys('//*[@id="numeroPerdcomp"]', Keys.HOME)
                    sb.sleep(0.5)
                    sb.send_keys('//*[@id="numeroPerdcomp"]', perdcomp)
                    sb.sleep(0.5)

                    # Função para tentar clicar nos botões de download
                    def try_download_buttons(sb, download_xpaths):
                        for xpath in download_xpaths:
                            try:
                                sb.wait_for_element_visible(xpath, timeout=5)
                                sb.click(xpath)
                                sb.sleep(5)
                                return True  # Sucesso no clique
                            except Exception:
                                continue
                        return False  # Falha em todos os cliques

                    # XPaths dos botões de download
                    download_xpaths = [
                        '//*[@id="page-content-wrapper"]/perdcomp-tela-inicial/div/perdcomp-tabs/perdcomp-tab[2]/div/perdcomp-listar-docs-enviados/div/div[1]/simple-collapsible/div/div[2]/div[7]/div[3]/i',
                        '//*[@id="page-content-wrapper"]/perdcomp-tela-inicial/div/perdcomp-tabs/perdcomp-tab[2]/div/perdcomp-listar-docs-enviados/div/div[1]/simple-collapsible/div/div[2]/div[7]/div[2]/i'
                    ]

                    # Tentar clicar em um dos botões de download
                    if not try_download_buttons(sb, download_xpaths):
                        print(f"Botão de download não encontrado para PERDCOMP {perdcomp}. Tentando novamente...")
                        continue  # Volta ao início do loop

                    sb.sleep(15)  # Aguarda tempo suficiente para o download começar

                    # Verificar se o arquivo foi baixado e mover para a pasta de destino
                    if os.path.exists(downloaded_file):
                        if os.path.exists(destination):
                            os.remove(destination)  # Remove o arquivo existente para sobrescrever
                        shutil.move(downloaded_file, destination)
                        print(f"PDF salvo em: {destination}")
                    else:
                        print(f"PDF {perdcomp_clean}.pdf não encontrado na pasta de downloads. Tentando novamente...")
                        sb.refresh()
                        #sb.wait_for_element_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
                        sb.click_if_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
                        sb.click('//*[@id="myTab"]/li[2]')

                except Exception as e:
                    print(f"Erro na tentativa {attempts} para baixar PERDCOMP {perdcomp}: {e}")
                    sb.refresh()
                    sb.click_if_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
                    sb.click('//*[@id="myTab"]/li[2]')
                    continue

            if attempts >= max_attempts:
                print(f"Não foi possível processar PERDCOMP {perdcomp} após {max_attempts} tentativas. Pulando para o próximo.")

            # Atualiza a página após cada download
            sb.refresh()
            #sb.wait_for_element_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
            sb.click_if_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
            sb.click('//*[@id="myTab"]/li[2]')

        except Exception as e:
            print(f"Erro ao processar PERDCOMP {perdcomp}: {e}")
            sb.refresh()
            #sb.wait_for_element_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
            sb.click_if_visible('//*[@id="sidebar-wrapper"]/ul/li[3]/a/div/div[2]')
            sb.click('//*[@id="myTab"]/li[2]')
            continue
    
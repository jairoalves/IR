import pandas as pd

import time
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from msedge.selenium_tools import Edge, EdgeOptions

from util.utilitarios_gerais import gera_saida_em_excel

class BDR(dict):

    def __init__(self, nome, els, escriturador='-'):

        dados_bdr = {
            'nome':	nome,
            'Nome de Pregão': els[0].text.strip(),
            'Códigos de Negociação': els[1].text.strip().replace('Mais Códigos\n', ''),
            'CNPJ':	els[2].text.strip(),
            'Classificação Setorial': els[4].text.strip(),
            'Site': els[5].text.strip(),
            'Escriturador': escriturador,
            'Quantidade': '-',
            'Classe': 'BDR'
        }

        super(BDR, self).__init__(dados_bdr)

class AcessoBDRsB3(Edge):

    def config_inicial(self):
        # webdriver
        self.edge_webdriver = r'C:\Users\u495\GitHub\PDP\webdrivers\msedgedriver.exe'

        # URLs
        self.url_pag_inicial_BDRs = r"http://bvmf.bmfbovespa.com.br/cias-listadas/Mercado-Internacional/Mercado-Internacional.aspx?idioma=pt-br"

        # arquivos de saida
        self.end_arq_info_pags_bdrs = './dados/saida/info_pags_bdrs_b3.xlsx'
        self.end_arq_detalhes_bdrs = './dados/saida/detalhes_bdrs_b3.xlsx'

        # listas de controle interno
        self.info_pags_bdrs = []
        self.lista_detalhes_bdrs = []

        # filtros web
        self.xpath_botao_cookies = "/html/body/div[3]/div[3]/div/div/div[2]/div/button[2]"
        self.id_iframe = "ctl00_contentPlaceHolderConteudo_iframeCarregadorPaginaExterna"
        self.xpath_dados_bdr = "//td[2]"
        self.css_aba_contatos = 'body > div:nth-child(4) > div.large-8.columns > ul > li:nth-child(2)'
        self.xpath_escriturador = '//*[@id="divContatos"]//td[2]'
        self.css_escriturador = '#divContatos > div > table > tbody > tr > td:nth-child(2)'
        self.seletores_tab_bdr = '#tblBdrs > tbody > tr > td.razaoSocial > a'
        self.xpath_qtde = '//*[@id="divComposicaoCapitalSocial"]/div/table/tbody/tr/td[2]'

    def __init__(self):
        self.config_inicial()
        self.abre_edge()

    def abre_edge(self):
        self.options = EdgeOptions()
        self.options.use_chromium = True

        super().__init__(
            executable_path=self.edge_webdriver,
            options=self.options)

    def abre_pag_inicial_BDRs(self):
        self.get(self.url_pag_inicial_BDRs)
        time.sleep(3)

    def aceita_cookies(self):
        """Aguarda botão de aceitar cookies e clica nele"""

        WebDriverWait(self, 20).until(
            EC.visibility_of_element_located(
                (By.XPATH, self.xpath_botao_cookies)))

        botao_cookies = self.find_element_by_xpath(self.xpath_botao_cookies)
        botao_cookies.click()
    
    def gera_tabela_info_BDRs(self, max=0):
        """Max permite dar um limite máximo de itens a coletar"""

        els_bdrs = self.find_elements_by_css_selector(self.seletores_tab_bdr)

        if max == 0:
            max = len(els_bdrs)

        for el in els_bdrs[:max]:
            print(f'Coletando link de {el.text}..')
            info_pag_bdr = {'nome': el.text.strip(), 'link': el.get_attribute('href')}
            self.info_pags_bdrs.append(info_pag_bdr)
    
    def gera_detalhes_lista_BDRs(self, max=0):
        """Max permite dar um limite máximo de itens a coletar"""

        df_links = pd.read_excel(self.end_arq_info_pags_bdrs)

        if max == 0:
            max = len(df_links)

        # para cada BDR da lista geral
        for idx, info_bdr in df_links.iterrows():
            # obtem seus detalhes
            nome_bdr, link_bdr = info_bdr.tolist()
            print(f'\nFazendo a coleta dos dados do BDR: {nome_bdr}..')
            bdr_atual = self.obtem_detalhes_um_bdr(nome_bdr, link_bdr)
            
            print('\n\n')
            print(bdr_atual)
        
            # salva detalhes do BDR atual na lista completa
            self.lista_detalhes_bdrs.append(bdr_atual)

            # checa o parâmetro de máximo
            if idx >= (max - 1):
                break
    
    def obtem_detalhes_um_bdr(self, nome_bdr, url_bdr):
        self.get(url_bdr)
        
        # Seleciona o frame
        WebDriverWait(self, 20).until(
            EC.visibility_of_element_located(
                (By.ID, self.id_iframe)))
        el_iframe = self.find_element_by_id(self.id_iframe)
        self.switch_to.frame(el_iframe)

        # Pega dados da Companhia
        WebDriverWait(self, 20).until(
            EC.visibility_of_element_located(
                (By.XPATH, self.xpath_dados_bdr)))
        els_dados_cia = self.find_elements_by_xpath(self.xpath_dados_bdr)

        # carrega dados da aba "Dados da Companhia"
        bdr_atual = BDR(nome_bdr, els_dados_cia)

        # carrega dados da aba "Contatos"
        el_aba_contatos = self.find_element_by_css_selector(self.css_aba_contatos)
        el_aba_contatos.click()
        escriturador = WebDriverWait(self, 20).until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, self.css_escriturador))).text
        
        bdr_atual['Escriturador'] = escriturador
    
        # Pega dados de Quantidade
        WebDriverWait(self, 20).until(
            EC.visibility_of_element_located(
                (By.XPATH, self.xpath_qtde)))
        quantidade = self.find_element_by_xpath(self.xpath_qtde).text
        bdr_atual['Quantidade'] = quantidade

        return bdr_atual
    
    def salva_excel_info_pags_BDRs(self):
        df_links_bdrs = pd.DataFrame(self.info_pags_bdrs).set_index('nome')
        df_links_bdrs.to_excel(self.end_arq_info_pags_bdrs)
        gera_saida_em_excel(
            {'Lista_Links_BDRs': df_links_bdrs},
            self.end_arq_detalhes_bdrs
        )
        return df_links_bdrs
    
    def salva_excel_detalhes_BDRs(self):
        df_detalhes_bdrs = pd.DataFrame(self.lista_detalhes_bdrs).set_index('nome')
        gera_saida_em_excel(
            {'Detalhes_BDRs': df_detalhes_bdrs},
            self.end_arq_detalhes_bdrs
        )
        return df_detalhes_bdrs


###################################################################
#                            Execução                             #
###################################################################
max_bdrs_a_coletar = 0 # 0 = máximo

# Monta planilha de links
# with AcessoBDRsB3() as edge:
#     edge.abre_pag_inicial_BDRs()
#     edge.gera_tabela_info_BDRs(max=max_bdrs_a_coletar)
#     edge.salva_excel_info_pags_BDRs()

# Monta planilha de detalhes
with AcessoBDRsB3() as edge:
    edge.abre_pag_inicial_BDRs()
    edge.gera_detalhes_lista_BDRs(max=max_bdrs_a_coletar)
    edge.salva_excel_detalhes_BDRs()
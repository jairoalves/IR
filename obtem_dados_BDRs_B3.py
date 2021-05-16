import pandas as pd

import time
import selenium.webdriver.support.ui as ui
from selenium.webdriver import ActionChains
from msedge.selenium_tools import Edge, EdgeOptions

class BDR(dict):

    def __init__(self, els):

        dados_bdr = {
            'nome':	els[0].text.strip(),
            'Códigos de Negociação': els[1].text.strip().replace('Mais Códigos\n', ''),
            'CNPJ':	els[2].text.strip(),
            'Classificação Setorial': els[4].text.strip(),
            'Site': els[5].text.strip(),
            'Quantidade': els[6].text.strip(),
            'Escriturador': 'BANCO B3 S.A.',
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
        self.seletores_tab_bdr = '#tblBdrs > tbody > tr > td.razaoSocial > a'

    def __init__(self):
        self.config_inicial()
        self.abre_edge()

    def abre_edge(self):
        self.options = EdgeOptions()
        self.options.use_chromium = True

        super().__init__(executable_path=self.edge_webdriver)

    def abre_pag_inicial_BDRs(self):
        self.get(self.url_pag_inicial_BDRs)
        time.sleep(3)

    def aceita_cookies(self):
        while len(self.find_elements_by_xpath(self.xpath_botao_cookies)) != 1:
            print('esperando botão de cookies carregar..')
            time.sleep(1)
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
    
    def salva_excel_info_pags_BDRs(self):
        df_links_bdrs = pd.DataFrame(self.info_pags_bdrs).set_index('nome')
        df_links_bdrs.to_excel(self.end_arq_info_pags_bdrs)
        return df_links_bdrs
    
    def salva_excel_detalhes_BDRs(self):
        df_detalhes_bdrs = pd.DataFrame(self.lista_detalhes_bdrs).set_index('nome')
        df_detalhes_bdrs.to_excel(self.end_arq_detalhes_bdrs)
        return df_detalhes_bdrs
    
    def gera_detalhes_lista_BDRs(self, max=0):
        """Max permite dar um limite máximo de itens a coletar"""
        if max == 0:
            max = len(self.info_pags_bdrs)

        df_links = pd.read_excel(self.end_arq_info_pags_bdrs)

        # para cada BDR da lista geral
        for idx, info_bdr in self.info_pags_bdrs[:max]:
            # obtem seus detalhes
            nome_bdr, link_bdr = info_bdr.tolist()
            print(f'Fazendo a coleta dos dados do BDR: {nome_bdr}..')
            bdr_atual = self.obtem_detalhes_um_bdr(link_bdr)
            
            print('\n\n')
            print(bdr_atual)
        
            # salva detalhes do BDR atual na lista completa
            self.lista_detalhes_bdrs.append(bdr_atual)
    
    def obtem_detalhes_um_bdr(self, url_bdr):
        self.get(url_bdr)
        time.sleep(3)

        el_iframe = self.find_element_by_id(self.id_iframe)

        self.switch_to.frame(el_iframe)

        els = self.find_elements_by_xpath(self.xpath_dados_bdr)
        bdr_atual = BDR(els)

        return bdr_atual


###################################################################
#                            Execução                             #
###################################################################
max_bdrs_a_coletar = 0 # 0 = máximo

with AcessoBDRsB3() as edge:
    edge.abre_pag_inicial_BDRs()
    edge.gera_tabela_info_BDRs(max=max_bdrs_a_coletar)
    edge.salva_excel_info_pags_BDRs()

with AcessoBDRsB3() as edge:
    edge.abre_pag_inicial_BDRs()
    edge.gera_detalhes_lista_BDRs(max=max_bdrs_a_coletar)
    edge.salva_excel_detalhes_BDRs()
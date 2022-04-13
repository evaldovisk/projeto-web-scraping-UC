from subprocess import CREATE_NO_WINDOW
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service


class Browser:
    def __init__(self, site, login, senha, webd):
        self.site = site
        self.login = login
        self.senha = senha
        self.webdriver = webd

        self.service = Service(self.webdriver)
        self.service.creationflags = CREATE_NO_WINDOW
        self.options = Options()
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--disable-extensions")
        self.options.add_argument("--window-size=1100,900")
        self.options.add_argument("--headless")
        self.navegador = webdriver.Chrome(executable_path=self.webdriver, chrome_options=self.options, service=self.service)

    # abre o navegador

    def abre_chrome_e_faz_login(self):
        self.navegador.get(self.site)

        # entra no site abre a barra lateral
        self.navegador.find_element_by_xpath('//*[@id="txtUsuario"]').clear()
        self.navegador.find_element_by_xpath('//*[@id="txtUsuario"]').send_keys(self.login)
        self.navegador.find_element_by_xpath('//*[@id="txtSenha"]').clear()
        self.navegador.find_element_by_xpath('//*[@id="txtSenha"]').send_keys(self.senha)
        self.navegador.find_element_by_xpath('//*[@id="btnEntrar"]').click()
        self.navegador.find_element_by_xpath('//*[@id="sidebar-menu"]/div[1]/ul/li/a').click()

    # pesquisa o documento e captura o UC
    def pesquisar_UC_usando_documento(self, documento):
        self.navegador.find_element_by_xpath('//*[@id="sidebar-menu"]/div[1]/ul/li/ul/li[1]/a').click()
        self.navegador.find_element_by_xpath('//*[@id="txt_Documento"]').clear()
        self.navegador.find_element_by_xpath('//*[@id="txt_Documento"]').send_keys(documento)
        self.navegador.find_element_by_xpath('//*[@id="btnAvancarS1"]').click()

        # html do site
        navegador_html = BeautifulSoup(self.navegador.page_source, 'html.parser')
        # try:
        #     error_1 = navegador_html.find('span', attrs={'id': 'lblErro'})
        #     if error_1.text == '*Documento Inválido':
        #         return ['ERROR 1', '0']
        #
        # except  AttributeError:
        #     try:
        #         error_2 = navegador_html.find('tbody', attrs={'id': 'tb_ucsDist_tbody'})
        #         if 'A conexão subjacente estava fechada' in error_2.find('td').text:
        #             return ['ERROR 2', '0']
        #     except AttributeError:
        try:
            ok = navegador_html.find('tbody', attrs={'id': 'tb_ucsDist_tbody'})
            return ['OK', str(ok.find('a').text)]
        except AttributeError:
            return ['FALSE', '0']

    def fechar_chrome(self):
        self.navegador.close()

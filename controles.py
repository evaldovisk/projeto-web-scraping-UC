from excel import Excel
from navegador import Browser


def converte_coluna_em_numero(coluna):
    alfabeto = {'a': 0, 'b': 1, 'c': 2, 'd': 3, 'e': 4, 'f': 5, 'g': 6, 'h': 7, 'i': 8, 'j': 9, 'k': 10, 'l': 11,
                'm': 12, 'n': 13, 'o': 14, 'p': 15, 'q': 16, 'r': 17, 's': 18, 't': 19, 'u': 20, 'v': 21, 'w': 22,
                'x': 23, 'y': 24, 'z': 25}
    coluna = coluna.lower()
    if len(coluna) == 1:
        coluna = alfabeto[coluna]
        return coluna
    else:
        coluna_caract_um = coluna[0]
        coluna_caract_dois = coluna[1]
        coluna = (alfabeto[coluna_caract_um] + 1) * 27 + alfabeto[coluna_caract_dois]
        return int(coluna) - 1


class App:
    def __init__(self, path, sheet, linha, site, login, senha, webdriver):
        self.arquivo = Excel(path, sheet, int(linha))
        self.browser = Browser(site, login, senha, webdriver)

    def captura_documentos(self, documento, digito):
        # Captura todos os cpfs da planilha
        conversao_documento = converte_coluna_em_numero(documento)
        conversao_digito = converte_coluna_em_numero(digito)
        return self.arquivo.captura_todos_cpf(conversao_documento, conversao_digito)

    def captura_contratos(self, contrato):
        # Captura todos os contratos da planilha
        conversao_contrato = converte_coluna_em_numero(contrato)
        return self.arquivo.captura_todos_contratos(conversao_contrato)

    def teste(self, lista):
        lista.insert('', 'end', value=('a', 'a', 'a', 'a'))

    def scrapyng_uc(self, lista_contratos, lista_documentos, treeview):
        # Faz o web scrapyng para capturar as ucs
        lista_uc = list()

        self.browser.abre_chrome_e_faz_login()
        for contrato, documento in zip(lista_contratos, lista_documentos):
            sv = self.browser.pesquisar_UC_usando_documento(documento)
            status, value = sv
            treeview.insert('', 'end', value=(contrato, documento, status, value))
            lista_uc.append(value)

        self.browser.fechar_chrome()
        return lista_uc

    def salva_arquivo_xlsx(self, uc, lista_uc):
        conversao_uc = converte_coluna_em_numero(uc)
        self.arquivo.salvar_ucs_excel(conversao_uc, lista_uc)

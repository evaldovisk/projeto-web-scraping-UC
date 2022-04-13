from controles import App
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from navegador import Browser
from excel import Excel

window = Tk()


class Funcs(Excel, Browser):
    def funcao_botao_buscar(self):
        app = App(self.entry_explorador.get(), self.entry_frame_3.get(), self.entry_lin_inicial.get(),
                  self.entry_site.get(), self.entry_login.get(), self.entry_senha.get(), self.entry_webdriver.get())

        lista_documento = app.captura_documentos(self.entry_col_documento.get(), self.entry_col_digito.get())
        lista_contrato = app.captura_contratos(self.entry_col_contrato.get())

        inserir_treeview = self.lista
        lista_uc = app.scrapyng_uc(lista_contrato, lista_documento, inserir_treeview)

        app.salva_arquivo_xlsx(self.entry_col_uc.get(), lista_uc)




class Desktop(Funcs):
    def __init__(self):
        self.window = window
        self.config_tela()
        self.frame()
        self.frame3()
        self.frame2()
        self.frame1()
        window.mainloop()

    # Configuração da janela de interface
    def browse_files(self):
        path = filedialog.askopenfilename(initialdir="/", title="Select a File")
        self.entry_explorador.delete(0, 'end')
        self.entry_explorador.insert(0, path)

    def webdriver_file(self):
        path = filedialog.askopenfilename(initialdir="/", title="Select a File")
        self.entry_webdriver.delete(0, 'end')
        self.entry_webdriver.insert(0, path)

    def config_tela(self):
        self.window.title('UC AUTO')
        self.window.configure(background='#fedccc')
        largura, altura = 720, 450
        largura_screen, altura_screen = window.winfo_screenwidth(), window.winfo_screenheight()

        positionX = int(largura_screen / 2 - largura / 2)
        positionY = int((altura_screen / 2 - altura / 2) - 80)


        window.geometry('{}x{}+{}+{}'.format(largura, altura, positionX, positionY))

        window.minsize(width=720, height=450)
        window.maxsize(width=720, height=900)

    def frame(self):
        # Frame 1 -> TreeView
        self.frame_1 = Frame(self.window)
        self.frame_1.place(relx=0.01, rely=0.285, relwidth=0.69, relheight=0.7)

        # Frame 2 -> Dados
        self.frame_2 = Frame(self.window)
        self.frame_2.place(relx=0.705, rely=0.01, relwidth=0.284, relheight=0.975)

        # Frame 3 -> Arquivo Excel
        self.frame_3 = Frame(self.window)
        self.frame_3.place(relx=0.01, rely=0.01, relwidth=0.69, relheight=0.265)

    def frame3(self):
        # Texto indicador do input do Explorador
        self.texto2_frame_3 = Label(self.frame_3, text='Selecione o arquivo XLSX', font='arial 10')
        self.texto2_frame_3.place(relx=0.05, rely=0.05)

        # Botão explorar "..."
        self.btn_explorador = Button(self.frame_3, text='...', justify='center', command=lambda: self.browse_files())
        self.btn_explorador.place(relx=0.80, rely=0.25, relwidth=0.07, relheight=0.2)

        # Caminha do Explorador
        self.entry_explorador = Entry(self.frame_3, font='arial 10', borderwidth=1, relief='solid'
                                      , background='white', justify='center')
        self.entry_explorador.place(relx=0.056, rely=0.26, relwidth=0.74, relheight=0.19)
        self.entry_explorador.insert(0, 'Path do arquivo')

        # Texto Indicador do input da sheet
        self.texto3_frame_3 = Label(self.frame_3, text='Nome da Sheet', font='arial 10', justify='center')
        self.texto3_frame_3.place(relx=0.05, rely=0.55)

        # Input do nome da sheet
        self.entry_frame_3 = Entry(self.frame_3, borderwidth=0.5, relief='solid', font='Arial 11', justify='center')
        self.entry_frame_3.place(relx=0.056, rely=0.75, relwidth=0.3, relheight=0.19)

        # Texto indicador do input do webdriver
        self.texto_webdriver = Label(self.frame_3, text='Selecione o webdriver', font='arial 10')
        self.texto_webdriver.place(relx=0.392, rely=0.55)

        # Botão do webdriver
        self.btn_webdriver = Button(self.frame_3, text='...', justify='center', command=lambda: self.webdriver_file())
        self.btn_webdriver.place(relx=0.80, rely=0.74, relwidth=0.07, relheight=0.2)

        # Caminha do webdriver
        self.entry_webdriver = Entry(self.frame_3, font='arial 10', borderwidth=1, relief='solid'
                                      , background='white', justify='center')
        self.entry_webdriver.place(relx=0.4, rely=0.75, relwidth=0.395, relheight=0.19)
        self.entry_webdriver.insert(0, 'Path do webdriver')

    def frame2(self):
        # Texto indicador da linha inicial dos dados
        self.text_lin_inicial = Label(self.frame_2, text='Linha inicial dos dados', font='arial 10')
        self.text_lin_inicial.place(relx=0.06, rely=0.01)

        # Input da linha inicial dos dados
        self.entry_lin_inicial = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10', justify='center')
        self.entry_lin_inicial.place(relx=0.075, rely=0.06, relwidth=0.25)

        # Linha e indicador da Coluna inicial
        self.text_coluna = Label(self.frame_2, text='---- Coluna ----------------------------', font='arial 10')
        self.text_coluna.place(relx=0.06, rely=0.15)

        # Texto indicador do input da coluna do contrato
        self.text_contrato = Label(self.frame_2, text='Contrato', font='arial 10')
        self.text_contrato.place(relx=0.121, rely=0.215)

        # Input da coluna do contrato
        self.entry_col_contrato = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10'
                                        , justify='center')
        self.entry_col_contrato.place(relx=0.14, rely=0.265, relwidth=0.3)

        # Texto indicador do input da coluna do UC
        self.text_uc = Label(self.frame_2, text='UC', font='arial 10')
        self.text_uc.place(relx=0.52, rely=0.215)

        # Input da coluna do UC
        self.entry_col_uc = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10'
                                  , justify='center')
        self.entry_col_uc.place(relx=0.54, rely=0.265, relwidth=0.3)

        # Texto indicador do input da coluna do documento
        self.text_documento = Label(self.frame_2, text='Doc', font='arial 10')
        self.text_documento.place(relx=0.121, rely=0.35)

        # Input da coluna do documento
        self.entry_col_documento = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10'
                                         , justify='center')
        self.entry_col_documento.place(relx=0.14, rely=0.4, relwidth=0.3)

        # Texto indicador do input da coluna do digito
        self.text_digito = Label(self.frame_2, text='Digito', font='arial 10')
        self.text_digito.place(relx=0.52, rely=0.35)

        # Input da coluna do digito
        self.entry_col_digito = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10', justify='center')
        self.entry_col_digito.place(relx=0.54, rely=0.4, relwidth=0.3)

        # Linha Coluna final
        self.text_coluna = Label(self.frame_2, text='--------------------------------------------', font='arial 10')
        self.text_coluna.place(relx=0.06, rely=0.48)

        # Texto do input do site
        self.text_site = Label(self.frame_2, text='Site', font='arial 10')
        self.text_site.place(relx=0.06, rely=0.55)

        # Input do site
        self.entry_site = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10', justify='center')
        self.entry_site.place(relx=0.075, rely=0.6, relwidth=0.85)

        # Texto indicador do input do login
        self.text_login = Label(self.frame_2, text='Login', font='arial 10')
        self.text_login.place(relx=0.06, rely=0.65)

        # Input do login
        self.entry_login = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10', justify='center')
        self.entry_login.place(relx=0.075, rely=0.7, relwidth=0.85)

        # Texto indicador do input da senha
        self.text_senha = Label(self.frame_2, text='Senha', font='arial 10')
        self.text_senha.place(relx=0.06, rely=0.75)

        # Input da senha do site
        self.entry_senha = Entry(self.frame_2, borderwidth=0.5, relief='solid', font='Arial 10', show='*'
                                 , justify='center')
        self.entry_senha.place(relx=0.075, rely=0.8, relwidth=0.85)

        # Botão de Buscar/start os comandos
        self.btn_buscar = Button(self.frame_2, text='Buscar', command=lambda: self.funcao_botao_buscar())
        self.btn_buscar.place(relx=0.3, rely=0.89, relwidth=0.4, relheight=0.1)

    def frame1(self):
        self.scrool_lista = Scrollbar(self.frame_1, orient='vertical')

        self.lista = ttk.Treeview(self.frame_1, columns=('oc', 'cpf', 'states', 'uc'), show='headings'
                                  , yscrollcommand=self.scrool_lista.set, style='mystyle.Treeview')
        self.lista.column('oc', width=60, anchor='center')
        self.lista.column('cpf', width=145, anchor='center')
        self.lista.column('states', width=60, anchor='center')
        self.lista.column('uc', width=125, anchor='center')
        self.lista.heading('oc', text='OC')
        self.lista.heading('cpf', text='CPF')
        self.lista.heading('states', text='State')
        self.lista.heading('uc', text='UC')
        self.lista.place(relx=0.01, rely=0.01, relwidth=0.95, relheight=0.98)

        style = ttk.Style()
        style.configure('mystyle.Treeview', font=('Arial', 10), anchor='center')

        self.scrool_lista.configure(command=self.lista.yview)
        self.scrool_lista.place(relx=0.96, rely=0.01, relwidth=0.03, relheight=0.98)


Desktop()


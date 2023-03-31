from banco import *
from etiqueta import GerarEtiqueta
from os import path, makedirs
from pathlib import Path
from playsound import playsound
from selenium import webdriver
from selenium.webdriver.common.by import By
from tkcalendar import DateEntry
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
import pandas as pd
import psutil
import pyperclip
import socket
import time
import webbrowser


# Apagando a interface de login e abrindo a interface do menu principal
def call_menu(user, perfil, lote):
    root.withdraw()
    for child in root.winfo_children():
        child.destroy()
    root.deiconify()
    MainMenu(root, user, perfil, lote)


# Voltar ao menu principal
def voltar_menu(user, perfil, lote):
    root.withdraw()
    for child in root.winfo_children():
        child.destroy()
    root.deiconify()
    MainMenu(root, user, perfil, lote)


# Sair do menu principal e retornar à tela de login
def logoff():
    root.withdraw()
    for child in root.winfo_children():
        child.destroy()
    root.deiconify()
    Application(root)


# Alteração de perfil de usuário
def alterar_perfil():
    AlterarPerfil(root)


# Alteração de senha do usuário
def alteracao_senha(usuario):
    SenhaUsuario(root, usuario)


# Atualização de usuário e senha do Atlas
def alteracao_atlas():
    AtualizacaoAtlas(root)


# Alteração de cadastro de materiais
def parametros_materiais():
    CadastroMateriais(root)


# Cadastro de fornecedor de REVERSA (UREV)
def cadastrar_fornecedor():
    CadastroFornecedor(root)


# Seleção de fornecedor para salvar packlist ou iniciar leitura
def selecionar_fornecedor():
    fornecedor = SelecionarFornecedor(root).fornecedor
    return fornecedor


# Atualizar tabela de materiais, utilizada ao iniciar a leitura
def atualizar_materiais():
    # Apagando a tabela de materiais local
    param_rede = BancoParametrosRede()
    param_local = BancoParametrosLocal()

    try:
        param_local.drop_materiais()
        param_local.create_materiais()

        # Copiando os dados da tabela da rede
        param_rede.cursor.execute(f"SELECT * FROM materiais")
        data_param = param_rede.cursor.fetchall()

        # Inserindo os dados na tabela local
        param_local.cursor.executemany(f"INSERT INTO materiais (id, codigo, descricao, qtd, card, giro) "
                                       f"VALUES (?, ?, ?, ?, ?, ?)", data_param)
        param_local.connect.commit()

    except:
        pass

    # Fechando as conexões
    param_rede.connect.close()
    param_local.connect.close()


# Selecionar o material e número de caixa, utilizado para gerar QCODE e/ou apagar caixa da leitura
def get_material(tipo_leitura, materiais):
    info = SelecionarMaterial(root, tipo_leitura, materiais).info
    return info


# Encerrar o Chromedriver
def encerrar_chromedriver():
    # Nome do processo
    process_name = 'chromedriver.exe'

    # Buscar pelos processos em execução
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            # Verificar se o nome do processo é igual ao nome especificado
            if proc.name() == process_name:
                # Encerra o processo
                os.kill(proc.info['pid'], 9)

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


# Exportação da produtividade das bancadas
def export_produtividade(perfil, usuario):
    Produtividade(root, perfil, usuario)


# Calcular e definir posição centralizada da janela
def coordenadas(largura, altura):
    # Obter a largura e altura da tela do usuário
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Definir largura e altura da janela da aplicação
    window_width = largura
    window_height = altura

    # Calcular as coordenadas para centralizar a janela na tela
    pos_x = (screen_width / 2) - (window_width / 2)
    pos_y = ((screen_height / 2) - (window_height / 2)) - 60

    window_info = (window_width, window_height, pos_x, pos_y)
    return window_info


# Abrir o linkedin ao clicar no botão dos créditos
def open_credits():
    webbrowser.open('https://www.linkedin.com/in/wanderson-carelli/')


# Criar a pasta de leituras local para salvar os pallets
def pasta_leitura(tipo_leitura):
    # Definir a data atual
    hoje = datetime.now()
    hoje = hoje.strftime('%d-%m-%Y')

    # Definir o caminho o desktop
    desktop = path.join(path.expanduser('~'), 'Desktop')
    desktop = Path(desktop)

    # Definir a pasta "Leituras"
    pasta_leituras = path.join(desktop, 'Leituras')
    pasta_leituras = Path(pasta_leituras)

    # Criar a pasta "Leituras" caso não exista
    if not path.isdir(pasta_leituras):
        makedirs(pasta_leituras)

    # Definir a pasta do tipo de leitura
    pasta_tipo = path.join(pasta_leituras, f'{tipo_leitura}')
    pasta_tipo = Path(pasta_tipo)

    # Criar a pasta do tipo de leitura caso não exista
    if not path.isdir(pasta_tipo):
        makedirs(pasta_tipo)

    # Definir a pasta com a data atual
    pasta_data = path.join(pasta_tipo, f'{hoje}')
    pasta_data = Path(pasta_data)

    # Criar a pasta do tipo de leitura caso não exista
    if not path.isdir(pasta_data):
        makedirs(pasta_data)

    return pasta_data


# Início da aplicação / Página de login
class Application:
    def __init__(self, master):
        self.fonte = ("Calibri", 15)  # Fonte padrão

        # Importar imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-login.png')
        self.img_cadastrar = PhotoImage(file=r'assets/images/btn-cadastrar.png')
        self.img_login = PhotoImage(file=r'assets/images/btn-login.png')
        self.img_show = PhotoImage(file=r'assets/images/btn-show.png')
        self.img_hide = PhotoImage(file=r'assets/images/btn-hide.png')

        # Conectando ao banco de dados de usuários
        self.bd_user = BancoUsuarios()

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 600)

        # Configurações principais da janela
        self.window = master
        self.window.title("TPC PLH | Framework IN - Login")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Criação da caixa de entrada do usuário
        self.txt_usuario = Entry(self.window, font=self.fonte, justify=LEFT, background="#F2F2F2", border=False)
        self.txt_usuario.place(width=190, height=35, x=120, y=364)
        self.txt_usuario.focus_set()

        # Criação da caixa de entrada da senha
        self.txt_senha = Entry(self.window, font=self.fonte, justify=LEFT, background="#F2F2F2", border=False, show="•")
        self.txt_senha.bind("<Return>", self.entrar)
        self.txt_senha.place(width=190, height=35, x=120, y=437)

        # Criação do botão de ocultar a senha
        self.btn_hide = Button(self.window, image=self.img_hide, border=False, command=self.ocultar_senha,
                               activebackground='#F2F2F2')

        # Criação do botão de mostrar a senha
        self.btn_show = Button(self.window, image=self.img_show, border=False, command=self.mostrar_senha,
                               activebackground='#F2F2F2')
        self.btn_show.image = self.img_show
        self.btn_show.place(width=30, height=30, x=285, y=439)

        # Criação do botão de cadastro
        self.btn_cadastrar = Button(self.window, image=self.img_cadastrar, border=False, command=self.cadastrar,
                                    activebackground='#FFFFFF')
        self.btn_cadastrar.image = self.img_cadastrar
        self.btn_cadastrar.place(width=150, height=35, x=35, y=520)

        # Criação do botão de login
        self.btn_login = Button(self.window, image=self.img_login, border=False, command=self.entrar,
                                activebackground='#FFFFFF')
        self.btn_login.image = self.img_login
        self.btn_login.place(width=150, height=35, x=210, y=520)

        # Mensagem de aviso de cadastro de usuário
        self.lbl_mensagem = Label(self.window, font=("Calibri", "12", "italic"), background="#FFFFFF", justify=CENTER)
        self.lbl_mensagem.place(width=300, height=25, x=50, y=480)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=575)

    def mostrar_senha(self):
        self.txt_senha["show"] = ""
        self.btn_show.destroy()

        # Criação do botão de ocultar a senha
        self.btn_hide = Button(self.window, image=self.img_hide, border=False, command=self.ocultar_senha,
                               activebackground='#F2F2F2')
        self.btn_hide.image = self.img_hide
        self.btn_hide.place(width=30, height=35, x=285, y=437)

    def ocultar_senha(self):
        self.txt_senha["show"] = "•"
        self.btn_hide.destroy()

        self.btn_show = Button(self.window, image=self.img_show, border=False, command=self.mostrar_senha,
                               activebackground='#F2F2F2')
        self.btn_show.image = self.img_show
        self.btn_show.place(width=30, height=30, x=285, y=439)

    def cadastrar(self):
        # Definir caracteres válidos para nome de usuário
        good_chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'

        # Atribuir usuário e senha em variáveis
        usuario = self.txt_usuario.get().capitalize().strip()  # Deixa o nome apenas com a primeira letra em maíuscula
        senha = self.txt_senha.get().strip()

        # Validação de digitação do nome de usuário e senha
        if len(usuario) == 0:
            self.lbl_mensagem["text"] = "Digite o usuário."
            return

        elif len(usuario) > 0:
            for i in range(0, len(usuario)):  # Validação de caracteres no nome de usuário
                if usuario[i] not in good_chars:
                    self.lbl_mensagem["text"] = 'Usuário inválido. Digite apenas letras.'
                    return

        elif len(senha) == 0:
            self.lbl_mensagem["text"] = "Digite a senha."
            return

        elif len(senha) < 6:
            self.lbl_mensagem["text"] = "A senha deve conter no mínimo 6 caracteres."
            return

        # Verificação de usuário já existente
        self.bd_user.cursor.execute(f"SELECT * FROM usuarios WHERE nome = '{usuario}' ")
        if self.bd_user.cursor.fetchone() is not None:
            self.lbl_mensagem["text"] = 'Usuário já cadastrado.'
            return

        try:  # Inserindo cadastro no banco de usuários
            self.bd_user.cursor.execute("INSERT INTO usuarios (nome, senha, perfil) VALUES (?, ?, ?)",
                                        (usuario, senha, 'Default'))
            self.bd_user.connect.commit()
            self.lbl_mensagem["text"] = 'Usuário cadastrado com sucesso!'

        except:  # Em caso de erro
            self.lbl_mensagem["text"] = 'Erro ao cadastrar o usuário.'

    def entrar(self, event=None):
        # Atribuir usuário e senha em variáveis
        usuario = self.txt_usuario.get().capitalize()  # Converte o nome para a apenas a primeira letra em maíuscula
        senha = self.txt_senha.get()

        if len(usuario) == 0:
            self.lbl_mensagem["text"] = 'Digite o usuário.'
            return
        elif len(senha) == 0:
            self.lbl_mensagem["text"] = 'Digite a senha.'
            return

        # Buscando pelo nome de usuário no banco de dados
        self.bd_user.cursor.execute(f"SELECT * FROM usuarios WHERE nome = '{usuario}'")
        data = self.bd_user.cursor.fetchone()

        # Caso o nome de usuário não esteja cadastrado
        if data is None:
            self.lbl_mensagem["text"] = 'Usuário não cadastrado.'
            return

        if senha == data[2]:
            perfil = data[3]
            lote = data[4]
            call_menu(usuario, perfil, lote)
        else:
            self.lbl_mensagem["text"] = 'Senha inválida.'


# Menu principal
class MainMenu:
    def __init__(self, master, user, perfil, lote):
        # Obter IP do usuário atual
        self.hostname = socket.gethostname()
        self.ip_address = socket.gethostbyname(self.hostname)

        # Configurações de parâmetros
        self.titulo = 'TPC PLH | Framework IN'
        self.user = user
        self.lote = lote
        self.bancada = '00'
        self.perfil = perfil
        self.menu_bar(master, perfil)

        # Definir fonte padrão
        self.fonte = ("Calibri", 15)
        self.fonte_welcome = ("Calibri", 20, "italic")

        # Importar imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-menu.png')
        self.img_transferencia = PhotoImage(file=r'assets/images/btn-transferencia.png')
        self.img_reversa = PhotoImage(file=r'assets/images/btn-reversa.png')
        self.img_laboratorio = PhotoImage(file=r'assets/images/btn-laboratorio.png')
        self.img_assinantes = PhotoImage(file=r'assets/images/btn-assinantes.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(750, 450)

        # Configurações principais da janela
        self.window = master
        self.window.title("TPC PLH | Framework IN - Menu Principal")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')
        self.frame = Frame(master)
        self.frame.pack()

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.frame, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Definição da hora atual
        agora = datetime.now()
        agora = agora.strftime('%H')
        agora = int(agora)

        # Definição da mensagem de saudação
        if 6 <= agora < 12:
            msg = f'Bom dia, {self.user}!'
        elif 12 <= agora < 18:
            msg = f'Boa tarde, {self.user}!'
        else:
            msg = f'Boa noite, {self.user}!'

        # Criação da label de mensagem de saudação
        self.lbl_welcome = Label(self.frame, text=msg, font=self.fonte_welcome, background='white')
        self.lbl_welcome.place(width=350, height=50, x=10, y=100)

        # Caixa de entrada do número da nota fiscal
        self.txt_nf = Entry(font=self.fonte, justify=CENTER, background="#F2F2F2", border=False)
        self.txt_nf.place(width=120, height=35, x=110, y=258)
        self.txt_nf.focus_set()

        # Botão para leitura de TRANSFERÊNCIA (USAD)
        self.btn_transferencia = Button(image=self.img_transferencia, border=False, command=self.leitura_usad,
                                        activebackground='#FFFFFF')
        self.btn_transferencia.image = self.img_transferencia
        self.btn_transferencia.place(width=150, height=35, x=27, y=356)

        # Botão para leitura de REVERSA (UREV)
        self.btn_reversa = Button(image=self.img_reversa, border=False, command=self.leitura_urev,
                                  activebackground='#FFFFFF')
        self.btn_reversa.image = self.img_reversa
        self.btn_reversa.place(width=150, height=35, x=209, y=356)

        # Botão para leitura de LABORATÓRIO (USUS)
        self.btn_laboratorio = Button(image=self.img_laboratorio, border=False, command=self.leitura_usus,
                                      activebackground='#FFFFFF')
        self.btn_laboratorio.image = self.img_laboratorio
        self.btn_laboratorio.place(width=150, height=35, x=390, y=356)

        # Botão para leitura de LABORATÓRIO (USUS)
        self.btn_assinantes = Button(image=self.img_assinantes, border=False, command=self.leitura_uass,
                                     activebackground='#FFFFFF')
        self.btn_assinantes.image = self.img_assinantes
        self.btn_assinantes.place(width=150, height=35, x=571, y=356)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=580, y=425)

        self.frame.mainloop()

    # Criação da barra de menus horizontal no topo da interface
    def menu_bar(self, master, perfil):
        menubar = Menu()

        # Menu Configurações
        configs = Menu(menubar, tearoff=False)
        if perfil == 'Owner':
            configs.add_command(label='Alterar perfil de usuário', command=alterar_perfil)
        configs.add_command(label='Alterar senha de usuário', command=self.senha_usuario)
        menubar.add_cascade(label='Configurações', menu=configs)

        # Ajustando configurações para os tipos de perfis
        if perfil != 'Default':
            # Menu Configurações
            configs.add_command(label='Alterar senha do Atlas', command=alteracao_atlas)
            configs.add_command(label='Cadastro de materiais', command=parametros_materiais)

            # Menu TRANSFERÊNCIA (USAD)
            producao_usad = Menu(menubar, tearoff=False)
            producao_usad.add_command(label='Exportar leitura', command=self.export_usad)
            producao_usad.add_command(label='Importar packlist', command=self.import_usad)

            # Menu REVERSA (UREV)
            producao_urev = Menu(menubar, tearoff=False)
            producao_urev.add_command(label='Cadastro de fornecedor', command=cadastrar_fornecedor)
            producao_urev.add_command(label='Exportar leitura', command=self.export_urev)
            producao_urev.add_command(label='Importar packlist', command=self.import_urev)

            # Menu LABORATÓRIO (USUS)
            producao_usus = Menu(menubar, tearoff=False)
            producao_usus.add_command(label='Exportar lote', command=self.export_usus)

            # Menu ASSINANTES (UASS)
            producao_uass = Menu(menubar, tearoff=False)
            producao_uass.add_command(label='Exportar seriais', command=self.export_assinantes)

            menubar.add_cascade(label='Transferência', menu=producao_usad)
            menubar.add_cascade(label='Reversa', menu=producao_urev)
            menubar.add_cascade(label='Laboratório', menu=producao_usus)
            menubar.add_cascade(label='Assinantes', menu=producao_uass)

        # Menu Produtividade
        produtividade = Menu(menubar, tearoff=False)
        produtividade.add_command(label='Consultar produtividade', command=self.call_produtividade)
        menubar.add_cascade(label='Produtividade', menu=produtividade)

        menubar.add_cascade(label='Sair', command=logoff)

        master.config(menu=menubar)

    # Chamando função de alteração de senha
    def senha_usuario(self):
        alteracao_senha(self.user)

    # Exportação da leitura de TRANSFERÊNCIA (USAD)
    def export_usad(self):
        nf = f'nf_{self.txt_nf.get()}'

        # Verificar se a nota fiscal foi digitada
        if len(self.txt_nf.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o número da nota fiscal.')
            return self.txt_nf.focus_set()

        elif len(self.txt_nf.get()) > 0:
            try:
                int(self.txt_nf.get())
            except:
                messagebox.showerror(self.titulo, 'Número de nota fiscal inválido.')
                return self.txt_nf.delete(0, END)

        # Directório e conexão com o banco de dados da rede
        db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\usad.db'
        bd_rede = BancoUsadRede()

        # Verificando se existe a nota fiscal no banco de dados
        bd_rede.cursor.execute(f"SELECT * FROM {nf}")

        if not bd_rede.cursor.fetchone():
            bd_rede.connect.close()
            messagebox.showerror(self.titulo, f'Leitura da nota fiscal {self.txt_nf.get()} não encontrada.')
            return self.txt_nf.delete(0, END)

        bd_rede.connect.close()

        # Conectando no banco de dados
        bd_export = sqlite3.connect(db_path)

        try:
            # Realizando a leitura do banco de dados
            df = pd.read_sql_query(f"SELECT * FROM {nf}", bd_export).drop(columns=['id'])

            # Caixa de diálogo solicitando onde salvar a planilha
            filename = filedialog.asksaveasfilename(filetypes=[('Excel files (.xlsx)', '.xlsx')],
                                                    defaultextension='xlsx',
                                                    initialfile=f'Leitura NF {self.txt_nf.get()}')

            # Se cancelar a caixa de diálogo
            if len(filename) < 1:
                self.txt_nf.delete(0, END)
                return messagebox.showwarning(self.titulo, 'Processo cancelado.')

            # Exportando do banco de dados para o excel
            df.to_excel(filename, index=False)
            bd_export.close()
            messagebox.showinfo(self.titulo, f'Leitura da nota fiscal {self.txt_nf.get()} exportado com sucesso!')

        except:
            bd_export.close()
            messagebox.showerror(self.titulo, f'Falha ao exportar a leitura da nota fiscal {self.txt_nf.get()}.')
        self.txt_nf.delete(0, END)

    # Importação do packlist de terminais de TRANSFERÊNCIA (USAD)
    def import_usad(self):
        nf = f'nf_{self.txt_nf.get()}'

        # Verificar se a nota fiscal foi digitada
        if len(self.txt_nf.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o número da nota fiscal.')
            return self.txt_nf.focus_set()

        elif len(self.txt_nf.get()) > 0:
            # Verificar a nota fiscal digitada contém apenas números
            try:
                int(self.txt_nf.get())
            except:
                messagebox.showerror(self.titulo, 'Número de nota fiscal inválido.')
                return self.txt_nf.delete(0, END)

            # Verificar se já existe packlist salvo com a nota fiscal informada
            bd_rede = BancoUsadRede()
            bd_rede.cursor.execute(f"SELECT name FROM sqlite_master WHERE type = 'table' AND name='{nf}'")

            if bd_rede.cursor.fetchone():
                bd_rede.connect.close()
                messagebox.showerror(self.titulo, f'Já existe um packlist da nota fiscal {self.txt_nf.get()}.')
                return self.txt_nf.delete(0, END)

            bd_rede.connect.close()

        # Definindo o usuário e senha de acesso ao Atlas
        bd_user = BancoUsuarios()
        bd_user.cursor.execute(f"SELECT nome, senha FROM usuarios WHERE perfil = 'Atlas'")
        data = bd_user.cursor.fetchone()
        atlas_user = data[0]
        atlas_password = data[1]
        bd_user.connect.close()

        # Caixa de dialogo solicitando a planilha do packlist
        filename = filedialog.askopenfilename(title='Selecione o arquivo do packlist',
                                              filetypes=[('Excel files (.xlsx, .xls)', '.xlsx .xls')])

        # Se cancelar a caixa de diálogo
        if len(filename) < 1:
            self.txt_nf.delete(0, END)
            return messagebox.showwarning(self.titulo, 'Processo cancelado.')

        # Conectando ao banco de dados
        bd_local = BancoUsadLocal()
        bd_rede = ''  # Variavel será usada para abrir a conexão com o banco de dados

        try:
            # Inicia o Chrome em modo oculto
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            chrome = webdriver.Chrome(options=options)

            # Abrindo o atlas e fazendo login na operação Palhoça
            chrome.get('http://atlas/nethome/index.do')
            time.sleep(1)  # Aguardar 1 segundo afim de evitar erro no carregamento da página
            chrome.find_element(By.XPATH, '//*[@id="usuario"]').send_keys(atlas_user)
            chrome.find_element(By.XPATH, '//*[@id="senha"]').send_keys(atlas_password)
            chrome.find_element(By.XPATH, '//*[@id="base"]').send_keys('PALHOCA')
            chrome.find_element(By.XPATH, '//*[@id="tableLogin"]/tbody/tr[3]/td/table/tbody/tr/td/a[1]').click()

            # Abrindo a planilha do packlist e selecionando os seriais à serem consultados
            df = pd.read_excel(filename, usecols=[3], dtype={0: str})
            cont_linhas = df.tail(1).index[-1]
            cont_linhas += 1
            first_row = 0

            if cont_linhas > 1000:
                last_row = 1000
            else:
                last_row = cont_linhas

            chrome.get('http://atlas/nethome/equipamento/relatorioEquipamentos.do?acao=prepareSearch')
            chrome.find_element(By.XPATH, '//*[@id="ui-id-1"]').click()

            # Criar tabela da nota fiscal no banco de dados local
            bd_local.create_table(nf)

            while cont_linhas >= last_row:
                rng = df.iloc[first_row:last_row]
                rng = rng.to_string(index=False, header=False)
                pyperclip.copy(rng)
                lista = pyperclip.paste()

                chrome.find_element(By.XPATH,
                                    '//*[@id="tabs-2"]/table/tbody/tr[2]/td[2]/textarea').send_keys('', lista)
                chrome.find_element(By.XPATH, '//*[@id="btnConsultar"]').click()

                # Variáveis para rodar a lista de seriais no Atlas
                stop = False
                i = 1

                while not stop:
                    try:
                        material = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[5]').text
                        enderecavel = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[2]').text
                        serial_num = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[1]').text
                        tecnologia = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[4]').text
                        estado = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[7]').text
                        local = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[9]').text
                        packlist = 'SIM'

                        if 'PBL' not in material:
                            material = material[10:]

                        if '/' in enderecavel:
                            split = enderecavel.find(' / ')
                            serial_primario = enderecavel[:split]
                            serial_secundario = enderecavel[split + 3:]
                        else:
                            serial_primario = serial_secundario = enderecavel

                        # Enviando as informações para o banco de dados
                        bd_local.cursor.execute(f"INSERT INTO {nf} (material, serial_primario, serial_secundario, "
                                                f"serial_number, tecnologia, estado, local, packlist) "
                                                f"VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                                (material, serial_primario, serial_secundario, serial_num, tecnologia,
                                                 estado, local, packlist))

                        i += 1

                    except:
                        stop = True

                chrome.find_element(By.XPATH, '//*[@id="btnClear"]').click()

                first_row = last_row
                if first_row + 1000 < cont_linhas:
                    last_row += 1000
                else:
                    if last_row == cont_linhas:
                        break
                    else:
                        last_row = cont_linhas

            # Fechando o Chrome
            chrome.close()

            # Encerrar o processo do Chromedriver
            encerrar_chromedriver()

            # Salvando as informações no banco de dados local
            bd_local.connect.commit()

            # Copiar os dados do banco de dados local
            bd_local.cursor.execute(f"SELECT * FROM {nf}")
            data = bd_local.cursor.fetchall()

            # Conectar no banco de dados da rede
            bd_rede = BancoUsadRede()

            # Criar tabela no banco de dados da rede
            bd_rede.create_table(nf)

            # Adicionar os dados no banco de dados da rede
            bd_rede.cursor.executemany(f"INSERT INTO {nf} (id, material, serial_primario, serial_secundario, "
                                       f"serial_number, tecnologia, estado, local, packlist, lote, caixa, bancada, "
                                       f"usuario, hora) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)

            # Salvar as alterações
            bd_rede.connect.commit()

            # Fechar conexão com o banco de dados
            bd_local.connect.close()
            bd_rede.connect.close()

            # Confirmação de importação
            messagebox.showinfo(self.titulo, f'Packlist da nota fiscal {self.txt_nf.get()} salvo com sucesso!')

        except:
            # Fechar conexão com o banco de dados
            bd_local.connect.close()
            bd_rede.connect.close()

            messagebox.showerror(self.titulo, 'Falha ao importar o packlist.')

        self.txt_nf.delete(0, END)

    # Exportação da leitura de REVERSA (UREV)
    def export_urev(self):
        # Definir número da nota fiscal
        nf = self.txt_nf.get()

        # Verificar se a nota fiscal foi digitada
        if len(nf) == 0:
            messagebox.showerror(self.titulo, 'Informe o número da nota fiscal.')
            return self.txt_nf.focus_set()

        # Verificando se a nota fiscal informada possui apenas números
        elif len(nf) > 0:
            try:
                int(nf)
            except:
                messagebox.showerror(self.titulo, 'Número de nota fiscal inválido.')
                return self.txt_nf.delete(0, END)

        # Definindo o nome do fornecedor
        fornecedor = selecionar_fornecedor()
        if len(fornecedor) == 0:
            return

        # Ajustando o nome do fornecedor apenas com a primeira letra maiúscula
        nome_fornecedor = fornecedor.title()

        # Transformando o nome em letras minúsculas e retirando caracteres como espaços, pontos e parenteses
        fornecedor = fornecedor.lower()
        fornecedor = fornecedor.replace(' ', '_').replace('.', '')
        fornecedor = fornecedor.replace('(', '').replace(')', '')

        # Conectando ao banco de dados da rede
        bd_rede = BancoUrevRede()
        bd_rede.cursor.execute(f"SELECT * FROM {fornecedor} WHERE nota_fiscal = {nf}")
        if not bd_rede.cursor.fetchone():
            bd_rede.connect.close()
            return messagebox.showerror(self.titulo,
                                        f'Leitura da nota fiscal {nf} do fornecedor {nome_fornecedor} não encontrada.')

        bd_rede.connect.close()

        # Definir directório do arquivo do banco de dados na rede
        db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\urev.db'

        # Conectando no banco de dados
        bd_export = sqlite3.connect(db_path)

        try:
            # Realizando a leitura do banco de dados
            df = pd.read_sql_query(f"SELECT * FROM {fornecedor} WHERE nota_fiscal = '{nf}'",
                                   bd_export).drop(columns=['id'])

            filename = filedialog.asksaveasfilename(filetypes=[('Excel files (.xlsx)', '.xlsx')],
                                                    defaultextension='xlsx',
                                                    initialfile=f'{nome_fornecedor} - Leitura NF {nf}')

            if len(filename) < 1:
                self.txt_nf.delete(0, END)
                return messagebox.showwarning(self.titulo, 'Processo cancelado.')

            df.to_excel(filename, index=False)
            bd_export.close()

            messagebox.showinfo(self.titulo,
                                f'Leitura da nota fiscal {nf} do fornecedor {nome_fornecedor} exportada com sucesso!')

        except:
            bd_export.close()
            messagebox.showerror(self.titulo,
                                 f'Falha ao exportar a leitura da nota fiscal {nf} do fornecedor {nome_fornecedor}.')

    # Importação do packlist de REVERSA (UREV)
    def import_urev(self):
        # Definir número da nota fiscal
        nf = self.txt_nf.get()

        # Verificar se a nota fiscal foi digitada
        if len(nf) == 0:
            messagebox.showerror(self.titulo, 'Informe o número da nota fiscal.')
            return self.txt_nf.focus_set()

        # Verificando se a nota fiscal informada possui apenas números
        elif len(nf) > 0:
            try:
                int(nf)
            except:
                messagebox.showerror(self.titulo, 'Número de nota fiscal inválido.')
                return self.txt_nf.delete(0, END)

        # Definindo o nome do fornecedor
        fornecedor = selecionar_fornecedor()
        if len(fornecedor) == 0:
            return

        # Ajustando o nome do fornecedor apenas com a primeira letra maiúscula
        nome_fornecedor = fornecedor.title()

        # Transformando o nome em letras minúsculas e retirando caracteres como espaços, pontos e parenteses
        fornecedor = fornecedor.lower()
        fornecedor = fornecedor.replace(' ', '_').replace('.', '')
        fornecedor = fornecedor.replace('(', '').replace(')', '')

        # Conectar no banco de dados
        bd_rede = BancoUrevRede()
        bd_local = BancoUrevLocal()

        # Verificar se já existe o packlist da mesma NF e fornecedor
        bd_rede.cursor.execute(f"SELECT nota_fiscal FROM '{fornecedor}' WHERE nota_fiscal = '{nf}'")

        if bd_rede.cursor.fetchall():
            bd_rede.connect.close()
            return messagebox.showerror(self.titulo,
                                        f'Já existe o packlist da nota fiscal {nf} do fornecedor {nome_fornecedor}.')

        # Caixa de dialogo solicitando a planilha do packlist
        filename = filedialog.askopenfilename(title='Selecione o arquivo do packlist',
                                              filetypes=[('Excel files (.xlsx, .xls)', '.xlsx .xls')])

        if len(filename) < 1:
            self.txt_nf.delete(0, END)
            return messagebox.showwarning(self.titulo, 'Processo cancelado.')

        # Definindo o usuário e senha de acesso ao Atlas
        bd_user = BancoUsuarios()
        bd_user.cursor.execute(f"SELECT nome, senha FROM usuarios WHERE perfil = 'Atlas'")
        data = bd_user.cursor.fetchone()
        bd_user.connect.close()
        atlas_user = data[0]
        atlas_password = data[1]

        try:
            # Inicia o Chrome em modo oculto
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            chrome = webdriver.Chrome(options=options)

            # Abrindo o atlas e fazendo login na operação Palhoça
            chrome.get('http://atlas/nethome/index.do')
            time.sleep(1)  # Aguardar 1 segundo afim de evitar erro no carregamento da página
            chrome.find_element(By.XPATH, '//*[@id="usuario"]').send_keys(atlas_user)
            chrome.find_element(By.XPATH, '//*[@id="senha"]').send_keys(atlas_password)
            chrome.find_element(By.XPATH, '//*[@id="base"]').send_keys('PALHOCA')
            chrome.find_element(By.XPATH, '//*[@id="tableLogin"]/tbody/tr[3]/td/table/tbody/tr/td/a[1]').click()

            # Abrindo a planilha do packlist e selecionando os seriais à serem consultados
            df = pd.read_excel(filename, usecols=[0], dtype={0: str})
            cont_linhas = df.tail(1).index[-1]
            cont_linhas += 1
            first_row = 0
            if cont_linhas > 1000:
                last_row = 1000
            else:
                last_row = cont_linhas

            chrome.get('http://atlas/nethome/equipamento/relatorioEquipamentos.do?acao=prepareSearch')
            chrome.find_element(By.XPATH, '//*[@id="ui-id-1"]').click()

            # Loop para consultar seriais enquanto a contagem não atingir a quantidade total de linhas
            while cont_linhas >= last_row:
                rng = df.iloc[first_row:last_row]
                rng = rng.to_string(index=False, header=False)
                pyperclip.copy(rng)
                lista = pyperclip.paste()

                chrome.find_element(By.XPATH,
                                    '//*[@id="tabs-2"]/table/tbody/tr[2]/td[2]/textarea').send_keys('', lista)
                chrome.find_element(By.XPATH, '//*[@id="btnConsultar"]').click()

                # Variáveis para rodar a lista de seriais no Atlas
                stop = False
                i = 1

                while not stop:
                    try:
                        material = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[5]').text
                        enderecavel = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[2]').text
                        serial_num = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[1]').text
                        tecnologia = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[4]').text
                        estado = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[7]').text
                        local = chrome.find_element(
                            By.XPATH, f'//*[@id="tabPorEquipamentoResult"]/table/tbody/tr[{i}]/td[9]').text
                        packlist = 'SIM'

                        if 'PBL' not in material:
                            material = material[10:]

                        if '/' in enderecavel:
                            split = enderecavel.find(' / ')
                            serial_primario = enderecavel[:split]
                            serial_secundario = enderecavel[split + 3:]
                        else:
                            serial_primario = serial_secundario = enderecavel

                        # Enviando as informações para o banco de dados
                        bd_local.cursor.execute(f"INSERT INTO {fornecedor} (nota_fiscal, material, serial_primario, "
                                                f"serial_secundario, serial_number, tecnologia, estado, local, "
                                                f"packlist) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                                (nf, material, serial_primario, serial_secundario, serial_num,
                                                 tecnologia, estado, local, packlist))

                        i += 1

                    except:
                        stop = True

                chrome.find_element(By.XPATH, '//*[@id="btnClear"]').click()

                first_row = last_row
                if first_row + 1000 < cont_linhas:
                    last_row += 1000
                else:
                    if last_row == cont_linhas:
                        break
                    else:
                        last_row = cont_linhas

            # Fechando o chrome
            chrome.close()

            # Encerrar o processo do Chromedriver
            encerrar_chromedriver()

            # Salvando as informações no banco de dados local
            bd_local.connect.commit()

            # Copiar os dados do banco de dados local
            bd_local.cursor.execute(f"SELECT nota_fiscal, material, serial_primario, serial_secundario, serial_number, "
                                    f"tecnologia, estado, local, packlist, lote, caixa, bancada, usuario, hora "
                                    f"FROM {fornecedor} WHERE nota_fiscal = {nf}")
            data = bd_local.cursor.fetchall()

            # Adicionar os dados no banco de dados da rede
            bd_rede.cursor.executemany(f"INSERT INTO {fornecedor} (nota_fiscal, material, serial_primario, "
                                       f"serial_secundario, serial_number, tecnologia, estado, local, packlist, lote, "
                                       f"caixa, bancada, usuario, hora) "
                                       f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)

            # Salvar as alterações
            bd_rede.connect.commit()

            # Fechar conexão com o banco de dados
            bd_local.connect.close()
            bd_rede.connect.close()

            # Mensagem de aviso confirmando a importação do packlist
            messagebox.showinfo(self.titulo,
                                f'Packlist da nota fiscal {nf} do fornecedor {nome_fornecedor} salvo com sucesso!')

        except:
            # Fechar conexão com o banco de dados
            bd_local.connect.close()
            bd_rede.connect.close()

            # Mensagem de erro de importação
            messagebox.showerror(self.titulo, f'Falha ao importar packlist.')

        self.txt_nf.delete(0, END)

    # Exportação da leitura de USUS
    def export_usus(self):
        cod_material = self.txt_nf.get()

        # Verificar se o código do material foi digitado
        if len(cod_material) == 0:
            messagebox.showerror(self.titulo, 'Informe o código do material.')
            return self.txt_nf.focus_set()

        bd_param = BancoParametrosRede()
        bd_param.cursor.execute(f"SELECT codigo FROM materiais WHERE codigo = '{cod_material}'")
        data = bd_param.cursor.fetchone()
        bd_param.connect.close()

        if not data:
            messagebox.showerror(self.titulo, 'Código de material inválido.')
            self.txt_nf.delete(0, END)
            return self.txt_nf.focus_set()

        export = ExportarLote(root, cod_material)
        if export == 'sucesso':
            self.txt_nf.delete(0, END)

    # Exportação da leitura de ASSINANTES (UREV)
    def export_assinantes(self):
        # Verificando se existe leitura para exportar
        bd_rede = BancoUassRede()
        data = bd_rede.cursor.execute(f"SELECT * FROM assinantes").fetchone()
        bd_rede.connect.close()

        if not data:
            return messagebox.showerror(self.titulo, f'Não existe leitura de materiais em assinante.')

        # Janela solicitando o directório para salvar a planilha
        filename = filedialog.asksaveasfilename(filetypes=[('Excel files', '.xlsx')], defaultextension='xlsx',
                                                initialfile=f'Leitura ASSINANTES')

        if len(filename) < 1:
            return messagebox.showwarning(self.titulo, 'Processo cancelado.')

        # Definir directório do arquivo do banco de dados na rede
        db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\uass.db'

        # Conectando no banco de dados
        bd_export = sqlite3.connect(db_path)

        try:
            df = pd.read_sql_query(f"SELECT * FROM assinantes", bd_export).drop(columns=['id'])
            df.to_excel(filename, index=False)
            bd_export.close()
            messagebox.showinfo(self.titulo, f'Leitura de assinantes exportada com sucesso!')

        except:
            bd_export.close()
            messagebox.showerror(self.titulo, f'Falha na exportação da leitura de assinantes.')

    # Iniciar a leitura de TRANSFERÊNCIA (USAD)
    def leitura_usad(self):
        nf = self.txt_nf.get().strip()

        if len(nf) == 0:
            messagebox.showerror(self.titulo, 'Informe o número da nota fiscal.')
            return self.txt_nf.focus_set()

        # Definindo o número da NF e conectando no banco de dados
        nf = f'nf_{self.txt_nf.get()}'
        tipo_leitura = 'TRANSFERENCIA'

        # Definir número da bancada
        self.bancada = self.get_bancada()

        if not self.bancada or self.bancada is None:
            self.bancada = SelecionarBancada(root, self.bancada).bancada

        if not self.bancada or self.bancada is None:
            return

        bd_rede = BancoUsadRede()

        # Efetuar a contagem de total de seriais da nota fiscal
        try:
            bd_rede.cursor.execute(f"SELECT COUNT(*) FROM {nf} WHERE packlist LIKE 'SIM%'")
            total_seriais = bd_rede.cursor.fetchone()
            total_seriais = total_seriais[0]

        except:
            bd_rede.connect.close()
            messagebox.showerror(self.titulo, f'Packlist da nota fiscal {self.txt_nf.get()} não encontrado.')
            return self.txt_nf.focus_set()

        # Conectar no banco de dados local
        bd_local = BancoUsadLocal()

        try:
            # Verificar se já existe a tabela da nota fiscal
            bd_local.cursor.execute(f"SELECT name FROM sqlite_master WHERE type = 'table' AND name='{nf}'")

            if not bd_local.cursor.fetchone():
                # Copiar a tabela do banco de dados da rede
                bd_rede.cursor.execute(f"SELECT * FROM {nf}")
                data = bd_rede.cursor.fetchall()

                # Conectar e criar tabela no banco de dados local
                bd_local.create_table(nf)

                # Adicionar os dados no banco de dados local
                bd_local.cursor.executemany(f"INSERT INTO {nf} (id, material, serial_primario, serial_secundario, "
                                            f"serial_number, tecnologia, estado, local, packlist, lote, caixa, "
                                            f"bancada, usuario, hora) "
                                            f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)

                # Salvo as informações no banco de dados local
                bd_local.connect.commit()

            else:
                # Copiar as informações atualizadas da tabela da rede
                bd_rede.cursor.execute(f"SELECT serial_primario, packlist, lote, caixa, bancada, usuario, hora "
                                       f"FROM {nf} WHERE lote IS NOT NULL")

                dados_leitura = bd_rede.cursor.fetchall()

                for i in range(0, len(dados_leitura)):
                    serial = dados_leitura[i][0]
                    packlist = dados_leitura[i][1]
                    lote = dados_leitura[i][2]
                    caixa = dados_leitura[i][3]
                    bancada = dados_leitura[i][4]
                    usuario = dados_leitura[i][5]
                    hora = dados_leitura[i][6]

                    # Inserir informações no banco de dados local
                    bd_local.cursor.execute(f"UPDATE {nf} SET packlist = '{packlist}', lote = '{lote}', "
                                            f"caixa = '{caixa}', bancada = '{bancada}', usuario = '{usuario}', "
                                            f"hora = '{hora}' WHERE serial_primario = '{serial}'")

                # Salvar as informações do banco de dados local
                bd_local.connect.commit()

                # Fechar conexão com o banco de dados
                bd_rede.connect.close()
                bd_local.connect.close()

        except:
            messagebox.showerror(self.titulo, 'Falha ao iniciar a leitura. Tente novamente.')
            # Fechar conexão com o banco de dados
            bd_rede.connect.close()
            bd_local.connect.close()
            return

        # Definir número de lote:
        bd_user = BancoUsuarios()
        bd_user.cursor.execute(f"SELECT lote FROM usuarios WHERE nome = '{self.user}'")
        data = bd_user.cursor.fetchone()

        # Conectando no banco de dados na rede
        bd_param = BancoParametrosRede()

        if data[0] is None:
            self.lote = bd_param.gerar_lote(self.user, self.bancada, 'TRANSFERENCIA', self.txt_nf.get())
            bd_param.connect.close()

        if self.lote is None:
            bd_param.connect.close()
            return messagebox.showerror(self.titulo, 'Falha ao gerar número de lote para iniciar leitura.')

        # Atualizar tabela local de materiais
        atualizar_materiais()

        # Comandos para apagar limpar a interface atual, antes de abrir a nova interface
        self.window.withdraw()
        for child in self.window.winfo_children():
            child.destroy()
        time.sleep(0.5)
        self.window.deiconify()

        Leitura(self.window, nf, total_seriais, tipo_leitura, self.bancada, self.user, self.lote, self.perfil)

    # Iniciar a leitura de REVERSA (UREV)
    def leitura_urev(self):
        # Definir número da nota fiscal
        nf = self.txt_nf.get().strip()

        # Verificar se a nota fiscal foi digitada
        if len(nf) == 0:
            messagebox.showerror(self.titulo, 'Informe o número da nota fiscal.')
            return self.txt_nf.focus_set()

        # Verificando se a nota fiscal informada possui apenas números
        elif len(nf) > 0:
            try:
                int(nf)
            except:
                messagebox.showerror(self.titulo, 'Número de nota fiscal inválido.')
                return self.txt_nf.delete(0, END)

        # Definir número da bancada
        self.bancada = self.get_bancada()

        if not self.bancada or self.bancada is None:
            self.bancada = SelecionarBancada(root, self.bancada).bancada

        if not self.bancada or self.bancada is None:
            return

        # Definindo o nome do fornecedor
        fornecedor = selecionar_fornecedor()
        if len(fornecedor) == 0:
            return

        # Ajustando o nome do fornecedor apenas com a primeira letra maiúscula
        nome_fornecedor = fornecedor.title()

        # Transformando o nome em letras minúsculas e retirando caracteres como espaços, pontos e parenteses
        fornecedor = fornecedor.lower()
        fornecedor = fornecedor.replace(' ', '_').replace('.', '')
        fornecedor = fornecedor.replace('(', '').replace(')', '')

        # Conectar no banco de dados da rede
        bd_rede = BancoUrevRede()

        try:
            # Efetuar a contagem de total de seriais da nota fiscal
            bd_rede.cursor.execute(f"SELECT COUNT(*) FROM {fornecedor} "
                                   f"WHERE nota_fiscal = '{nf}' AND packlist LIKE 'SIM%'")
            total_seriais = bd_rede.cursor.fetchone()
            total_seriais = total_seriais[0]
        except:
            bd_rede.connect.close()
            messagebox.showerror(self.titulo,
                                 f'Packlist da nota fiscal {nf} do fornecedor {nome_fornecedor} não encontrado.')
            return

        # Conectar no banco de dados local
        bd_local = BancoUrevLocal()

        try:
            # Verificar se já existe seriais bipados da nota fiscal
            bd_local.cursor.execute(f"SELECT COUNT(*) FROM {fornecedor} WHERE nota_fiscal = '{nf}'")
            cont = bd_local.cursor.fetchone()

            # Se não existir seriais bipados, copiar toda a tabela da rede para a tabela local
            if cont[0] == 0:
                # Copiar a tabela do banco de dados da rede
                bd_rede.cursor.execute(f"SELECT nota_fiscal, material, serial_primario, serial_secundario, "
                                       f"serial_number, tecnologia, estado, local, packlist, lote, caixa, bancada, "
                                       f"usuario, hora FROM {fornecedor} WHERE nota_fiscal = {nf}")
                data = bd_rede.cursor.fetchall()

                # Adicionar os dados no banco de dados local
                bd_local.cursor.executemany(f"INSERT INTO {fornecedor} (nota_fiscal, material, serial_primario, "
                                            f"serial_secundario, serial_number, tecnologia, estado, local, packlist, "
                                            f"lote, caixa, bancada, usuario, hora) "
                                            f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)

            # Se já existir seriais bipados, atualizar a tabela local
            else:
                # Copiar as informações atualizadas da tabela da rede
                bd_rede.cursor.execute(f"SELECT nota_fiscal, serial_primario, packlist, lote, caixa, bancada, usuario, "
                                       f"hora FROM {fornecedor} WHERE nota_fiscal = '{nf}' AND lote IS NOT NULL")
                dados_leitura = bd_rede.cursor.fetchall()

                for i in range(0, len(dados_leitura)):
                    nota_fiscal = dados_leitura[i][0]
                    serial_primario = dados_leitura[i][1]
                    packlist = dados_leitura[i][2]
                    lote = dados_leitura[i][3]
                    caixa = dados_leitura[i][4]
                    bancada = dados_leitura[i][5]
                    usuario = dados_leitura[i][6]
                    hora = dados_leitura[i][7]

                    # Inserir informações no banco de dados local
                    bd_local.cursor.execute(f"UPDATE {fornecedor} SET nota_fiscal = '{nota_fiscal}', "
                                            f"packlist = '{packlist}', lote = '{lote}', caixa = '{caixa}', "
                                            f"bancada = '{bancada}', usuario = '{usuario}', hora = '{hora}' "
                                            f"WHERE serial_primario = '{serial_primario}'")

            # Salvar as informações do banco de dados local
            bd_local.connect.commit()

            # Fechar conexão com o banco de dados
            bd_rede.connect.close()
            bd_local.connect.close()

        except:
            messagebox.showerror(self.titulo, 'Falha ao iniciar a leitura. Tente novamente.')
            # Fechar conexão com o banco de dados
            bd_rede.connect.close()
            bd_local.connect.close()
            return

        # Definir número de lote:
        bd_user = BancoUsuarios()
        bd_user.cursor.execute(f"SELECT lote FROM usuarios WHERE nome = '{self.user}'")
        data = bd_user.cursor.fetchone()

        # Conectando no banco de dados na rede
        bd_param = BancoParametrosRede()

        if data[0] is None:
            self.lote = bd_param.gerar_lote(self.user, self.bancada, 'REVERSA', nf)
            bd_param.connect.close()

        if self.lote is None:
            bd_param.connect.close()
            return messagebox.showerror(self.titulo, 'Falha ao gerar número de lote para iniciar leitura.')

        # Atualizar tabela local de materiais
        atualizar_materiais()

        # Comandos para apagar limpar a interface atual, antes de abrir a nova interface
        self.window.withdraw()
        for child in self.window.winfo_children():
            child.destroy()
        time.sleep(0.5)
        self.window.deiconify()

        Leitura(self.window, nf, total_seriais, 'REVERSA', self.bancada, self.user, self.lote, self.perfil, fornecedor)

    # Iniciar a leitura de USUS
    def leitura_usus(self):
        cod_material = self.txt_nf.get().strip()

        if len(cod_material) == 0:
            messagebox.showerror(self.titulo, 'Informe o código do material.')
            return self.txt_nf.focus_set()

        else:
            bd_param = BancoParametrosRede()
            bd_param.cursor.execute(f"SELECT codigo FROM materiais WHERE codigo = '{cod_material}'")
            data = bd_param.cursor.fetchone()
            bd_param.connect.close()

            if not data:
                messagebox.showerror(self.titulo, 'Código de material inválido.')
                self.txt_nf.delete(0, END)
                return self.txt_nf.focus_set()

        nf = f'nf_{self.txt_nf.get()}'
        bd_local = BancoUsusLocal()
        bd_local.create_table(nf)
        bd_local.connect.close()
        tipo_leitura = 'LABORATORIO'
        total_seriais = '-'

        # Definir número da bancada
        self.bancada = self.get_bancada()

        if not self.bancada or self.bancada is None:
            self.bancada = SelecionarBancada(root, self.bancada).bancada

        if not self.bancada or self.bancada is None:
            return

        # Definir número de lote:
        bd_user = BancoUsuarios()
        bd_user.cursor.execute(f"SELECT lote FROM usuarios WHERE nome = '{self.user}'")
        data = bd_user.cursor.fetchone()

        # Conectando no banco de dados na rede
        bd_param = BancoParametrosRede()

        if data[0] is None:
            self.lote = bd_param.gerar_lote(self.user, self.bancada, 'LABORATORIO', self.txt_nf.get())
            bd_param.connect.close()

        if self.lote is None:
            bd_param.connect.close()
            return messagebox.showerror(self.titulo, 'Falha ao gerar número de lote para iniciar leitura.')

        # Atualizar tabela local de materiais
        atualizar_materiais()

        # Comandos para apagar limpar a interface atual, antes de abrir a nova interface
        self.window.withdraw()
        for child in self.window.winfo_children():
            child.destroy()
        time.sleep(0.5)
        self.window.deiconify()

        Leitura(self.window, nf, total_seriais, tipo_leitura, self.bancada, self.user, self.lote, self.perfil)

    # Iniciar a leitura de ASSINANTES (UASS)
    def leitura_uass(self):
        # Definir número da bancada
        self.bancada = self.get_bancada()

        if not self.bancada or self.bancada is None:
            self.bancada = SelecionarBancada(root, self.bancada).bancada

        if not self.bancada or self.bancada is None:
            return

        # Definir número de lote:
        bd_user = BancoUsuarios()
        bd_user.cursor.execute(f"SELECT lote FROM usuarios WHERE nome = '{self.user}'")
        data = bd_user.cursor.fetchone()

        # Conectando no banco de dados na rede
        bd_param = BancoParametrosRede()

        if data[0] is None:
            self.lote = bd_param.gerar_lote(self.user, self.bancada, 'ASSINANTES', "ASSIN.")
            bd_param.connect.close()

        if self.lote is None:
            bd_param.connect.close()
            return messagebox.showerror(self.titulo, 'Falha ao gerar número de lote para iniciar leitura.')

        # Atualizar tabela local de materiais
        atualizar_materiais()

        # Comandos para apagar limpar a interface atual, antes de abrir a nova interface
        self.window.withdraw()
        for child in self.window.winfo_children():
            child.destroy()
        time.sleep(0.5)
        self.window.deiconify()

        Leitura(self.window, '-', '-', 'ASSINANTES', self.bancada, self.user, self.lote, self.perfil)

    # Obter o número da bancada
    def get_bancada(self):
        bancada = None

        # Definindo o número da bancada com base no IP
        if self.ip_address == '10.69.78.169':
            bancada = '01'
        elif self.ip_address == '10.69.78.157':
            bancada = '02'
        elif self.ip_address == '10.69.78.168':
            bancada = '03'
        elif self.ip_address == '10.69.78.151':
            bancada = '04'
        elif self.ip_address == '10.69.78.155':
            bancada = '05'

        return bancada

    # Chamar a interface de produtividade
    def call_produtividade(self):
        export_produtividade(self.perfil, self.user)


# Alteração de tipo de perfil de usuário
class AlterarPerfil:
    def __init__(self, master):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-alterar_perfil.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_salvar = PhotoImage(file=r'assets/images/btn-salvar.png')

        # Conectando ao banco de dados de usuários
        self.bd_user = BancoUsuarios()
        self.bd_user.cursor.execute(f"SELECT nome FROM usuarios "
                                    f"WHERE perfil != 'Atlas' AND perfil != 'Owner' ORDER BY nome ASC")
        self.data = self.bd_user.cursor.fetchall()

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title('TPC PLH | Framework IN - Alterar Perfil')
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        # Combobox com os nomes de usuário
        self.cb_usuarios = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(),
                                        justify=CENTER)
        self.cb_usuarios.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_usuarios['values'] = self.data
        self.cb_usuarios['state'] = 'readonly'
        self.cb_usuarios.place(width=160, height=35, x=135, y=265)

        # Combobox com os tipos de perfil
        self.cb_perfis = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(), justify=CENTER)
        self.cb_perfis.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_perfis['values'] = ('ADM', 'PADRÃO')
        self.cb_perfis['state'] = 'readonly'
        self.cb_perfis.place(width=160, height=35, x=135, y=328)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para salvar
        self.btn_salvar = Button(self.window, image=self.img_salvar, border=False, command=self.salvar,
                                 activebackground='#FFFFFF')
        self.btn_salvar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def salvar(self):
        # Verificando a seleção de usuário e tipo de perfil
        if len(self.cb_usuarios.get()) == 0:
            return messagebox.showerror(self.titulo, 'Selecione o usuário a ser alterado.')
        elif len(self.cb_perfis.get()) == 0:
            return messagebox.showerror(self.titulo, 'Selecione o novo perfil de usuário.')

        # Atribuindo a nomeclatura do tipo de perfil usado no banco de dados
        if self.cb_perfis.get() == 'PADRÃO':
            permissao = 'Default'
        else:
            permissao = 'ADM'

        try:  # Atualizando o banco de dados
            username = self.cb_usuarios.get().capitalize()
            self.bd_user.cursor.execute(f"UPDATE usuarios SET perfil = '{permissao}' WHERE nome = '{username}'")
            self.bd_user.connect.commit()
            messagebox.showinfo(self.titulo, 'Perfil de usuário alterado com sucesso!')

        except:  # Em caso de erro
            messagebox.showerror(self.titulo, 'Falha na alteração do perfil.')

        # Fechar a janela de alteração
        self.cancelar()

    def cancelar(self):
        self.bd_user.connect.close()
        self.window.destroy()


# Alteração de senha do usuário
class SenhaUsuario:
    def __init__(self, master, user):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão
        self.user = user  # Usuário atual

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-alterar_senha.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_salvar = PhotoImage(file=r'assets/images/btn-salvar.png')
        self.img_show = PhotoImage(file=r'assets/images/btn-show.png')
        self.img_hide = PhotoImage(file=r'assets/images/btn-hide.png')

        # Conectando ao banco de dados de usuários
        self.bd_user = BancoUsuarios()

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title('TPC PLH | Framework IN - Alterar Senha')
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Exibir o nome do usuário atual
        self.lbl_usuario = Label(self.window, text=self.user, font=self.fonte, background="#F2F2F2", justify=CENTER)
        self.lbl_usuario.place(width=146, height=35, x=127, y=264)

        # Entrybox para digitar a nova senha
        self.txt_senha = Entry(self.window, font=self.fonte, justify=LEFT, background="#F2F2F2", border=False, show="•")
        self.txt_senha.bind("<Return>", self.salvar)
        self.txt_senha.place(width=140, height=35, x=130, y=326)

        # Criação do botão de ocultar a senha
        self.btn_hide = Button(self.window, image=self.img_hide, border=False, command=self.ocultar_senha,
                               activebackground='#F2F2F2')

        # Criação do botão de mostrar a senha
        self.btn_show = Button(self.window, image=self.img_show, border=False, command=self.mostrar_senha,
                               activebackground='#F2F2F2')
        self.btn_show.image = self.img_show
        self.btn_show.place(width=30, height=30, x=273, y=329)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para salvar
        self.btn_salvar = Button(self.window, image=self.img_salvar, border=False, command=self.salvar,
                                 activebackground='#FFFFFF')
        self.btn_salvar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Entrando na caixa de texto de senha
        self.txt_senha.focus_set()

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def salvar(self, event=None):
        # Verificação de digitação da senha (no mínimo 6 caracteres)
        if len(self.txt_senha.get()) == 0:
            messagebox.showerror(self.titulo, 'Digite a nova senha.')
            return self.txt_senha.focus_set()
        elif len(self.txt_senha.get()) < 6:
            messagebox.showerror(self.titulo, 'A senha deve conter no mínimo 6 caracteres.')
            return self.txt_senha.delete(0, END)

        # Atribuindo o usuário e senha em variáveis
        usuario = self.user
        senha = self.txt_senha.get()

        try:  # Atualizando o banco de dados
            self.bd_user.cursor.execute(f"UPDATE usuarios SET senha = '{senha}' WHERE nome = '{usuario}'")
            self.bd_user.connect.commit()
            messagebox.showinfo(self.titulo, 'Senha alterada com sucesso!')

        except:  # Em caso de erro
            messagebox.showerror(self.titulo, 'Falha na alteração da senha.')

        # Fechar a janela
        self.cancelar()

    def cancelar(self):
        self.bd_user.connect.close()
        self.window.destroy()

    def mostrar_senha(self):
        self.txt_senha["show"] = ""
        self.btn_show.destroy()

        # Criação do botão de ocultar a senha
        self.btn_hide = Button(self.window, image=self.img_hide, border=False, command=self.ocultar_senha,
                               activebackground='#F2F2F2')
        self.btn_hide.image = self.img_hide
        self.btn_hide.place(width=30, height=30, x=273, y=327)

    def ocultar_senha(self):
        self.txt_senha["show"] = "•"
        self.btn_hide.destroy()

        self.btn_show = Button(self.window, image=self.img_show, border=False, command=self.mostrar_senha,
                               activebackground='#F2F2F2')
        self.btn_show.image = self.img_show
        self.btn_show.place(width=30, height=30, x=273, y=329)


# Alteração de usuário e senha do Atlas
class AtualizacaoAtlas:
    def __init__(self, master):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-atualizar_atlas.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_salvar = PhotoImage(file=r'assets/images/btn-salvar.png')
        self.img_show = PhotoImage(file=r'assets/images/btn-show.png')
        self.img_hide = PhotoImage(file=r'assets/images/btn-hide.png')

        # Conectando ao banco de dados de usuários
        self.bd_user = BancoUsuarios()

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title("TPC PLH | Framework IN - Atualização Atlas")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Entrybox para inserir o usuário
        self.txt_usuario = Entry(self.window, font=self.fonte, justify=LEFT, background="#F2F2F2", border=False)
        self.txt_usuario.place(width=150, height=35, x=130, y=264)

        # Entrybox para inserir a senha
        self.txt_senha = Entry(self.window, font=self.fonte, justify=LEFT, background="#F2F2F2", border=False, show="•")
        self.txt_senha.bind("<Return>", self.salvar)
        self.txt_senha.place(width=140, height=35, x=130, y=326)

        # Botão de ocultar a senha
        self.btn_hide = Button(self.window, image=self.img_hide, border=False, command=self.ocultar_senha,
                               activebackground='#F2F2F2')

        # Botão de mostrar a senha
        self.btn_show = Button(self.window, image=self.img_show, border=False, command=self.mostrar_senha,
                               activebackground='#F2F2F2')
        self.btn_show.image = self.img_show
        self.btn_show.place(width=30, height=30, x=273, y=329)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para salvar
        self.btn_salvar = Button(self.window, image=self.img_salvar, border=False, command=self.salvar,
                                 activebackground='#FFFFFF')
        self.btn_salvar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Entrando na caixa de texto de usuário
        self.txt_usuario.focus_set()

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def salvar(self, event=None):
        # Verificando se foi digitado usuário e senha
        if len(self.txt_usuario.get()) == 0:
            messagebox.showerror(self.titulo, 'Digite o nome do usuário.')
            return self.txt_usuario.focus_set()
        elif len(self.txt_senha.get()) == 0:
            messagebox.showerror(self.titulo, 'Digite a senha do usuário.')
            return self.txt_senha.focus_set()

        # Atribuindo o usuário e senha em variáveis
        usuario = self.txt_usuario.get().upper()
        senha = self.txt_senha.get()

        try:  # Atualizando o banco de dados
            self.bd_user.cursor.execute(f"UPDATE usuarios SET nome = '{usuario}', senha = '{senha}' "
                                        f"WHERE perfil = 'Atlas'")
            self.bd_user.connect.commit()
            messagebox.showinfo(self.titulo, 'Login do Atlas atualizado com sucesso!')

        except:  # Em caso de erro
            messagebox.showerror(self.titulo, 'Falha na atualização do Atlas.')

        self.cancelar()

    def cancelar(self):
        self.bd_user.connect.close()
        self.window.destroy()

    def mostrar_senha(self):
        self.txt_senha["show"] = ""
        self.btn_show.destroy()

        # Criação do botão de ocultar a senha
        self.btn_hide = Button(self.window, image=self.img_hide, border=False, command=self.ocultar_senha,
                               activebackground='#F2F2F2')
        self.btn_hide.image = self.img_hide
        self.btn_hide.place(width=30, height=30, x=273, y=327)

    def ocultar_senha(self):
        self.txt_senha["show"] = "•"
        self.btn_hide.destroy()

        self.btn_show = Button(self.window, image=self.img_show, border=False, command=self.mostrar_senha,
                               activebackground='#F2F2F2')
        self.btn_show.image = self.img_show
        self.btn_show.place(width=30, height=30, x=273, y=329)


# Cadastro de materiais
class CadastroMateriais:
    def __init__(self, master):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-cad_materiais.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_salvar = PhotoImage(file=r'assets/images/btn-salvar.png')

        # Conectando ao banco de dados de parâmetros
        self.bd_param = BancoParametrosRede()

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(450, 650)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title("TPC PLH | Framework IN - Cadastro de Materiais")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Entrybox para digitar o código do material
        self.txt_codigo = Entry(self.window, font=self.fonte, justify=CENTER, background="#F2F2F2", border=False)
        self.txt_codigo.bind('<FocusOut>', self.update)
        self.txt_codigo.place(width=125, height=35, x=57, y=311)

        # Entrybox para digitar a quantidade padrão por caixa
        self.txt_quantidade = Entry(self.window, font=self.fonte, justify=CENTER, background="#F2F2F2", border=False)
        self.txt_quantidade.place(width=125, height=35, x=268, y=311)

        # Entrybox para digitar a descrição do material
        self.txt_descricao = Entry(self.window, font=self.fonte, justify=LEFT, background="#F2F2F2", border=False)
        self.txt_descricao.place(width=308, height=35, x=86, y=408)

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        # Combobox com as opções de tipo de giro
        self.cb_giro = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(), justify=CENTER)
        self.cb_giro.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_giro['values'] = ('ALTO GIRO', 'BAIXO GIRO', 'SUCATA')
        self.cb_giro['state'] = 'readonly'
        self.cb_giro.place(width=125, height=35, x=57, y=505)

        # Combobox com as opções de código de card interno
        self.cb_card = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(), justify=CENTER)
        self.cb_card.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_card['values'] = ('NÃO', '41000805', '41001186')
        self.cb_card['state'] = 'readonly'
        self.cb_card.place(width=125, height=35, x=268, y=505)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=45, y=565)

        # Botão para salvar
        self.btn_salvar = Button(self.window, image=self.img_salvar, border=False, command=self.salvar,
                                 activebackground='#FFFFFF')
        self.btn_salvar.place(width=150, height=35, x=255, y=565)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=280, y=625)

        # Entrando na caixa de texto de código do material
        self.txt_codigo.focus_set()

        # Manter a janela de cadastro ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de cadastro estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de cadastro
        self.window.grab_release()

    def salvar(self):
        # Atribuindo os valores em variáveis
        codigo = self.txt_codigo.get()
        quantidade = self.txt_quantidade.get()
        descricao = self.txt_descricao.get().upper()
        giro = self.cb_giro.get()
        card = self.cb_card.get()

        # Verificar se há alguma informação em branco
        if len(codigo) == 0:
            messagebox.showerror(self.titulo, 'Informe o código do material.')
            return self.txt_codigo.focus_set()

        elif len(quantidade) == 0:
            messagebox.showerror(self.titulo, 'Informe a quantidade padrão por caixa.')
            return self.txt_quantidade.focus_set()

        elif len(descricao) == 0:
            messagebox.showerror(self.titulo, 'Informe a descrição do material.')
            return self.txt_descricao.focus_set()

        elif len(giro) == 0:
            return messagebox.showerror(self.titulo, 'Informe o tipo de giro.')

        elif len(card) == 0:
            return messagebox.showerror(self.titulo, 'Informe se o material possui card interno.')

        # Validar tamanho e início do código
        if len(codigo) != 8 or codigo[:4] != '4100':
            messagebox.showerror(self.titulo, 'Código de material inválido.')
            return self.txt_codigo.focus_set()

        # Verificando se foram digitados apenas números no código e quantidade
        try:
            validar_cod = int(codigo)
            validar_qtd = int(quantidade)

        except:
            validar_cod = ''
            validar_qtd = ''

        # Mensagem de erro caso o código ou quantidade esteja inválida
        if type(validar_cod) == str:
            return messagebox.showerror(self.titulo, 'O código de material deve conter apenas números.')
        elif type(validar_qtd) == str:
            return messagebox.showerror(self.titulo, 'A quantidade padrão deve ser um número válido.')

        self.bd_param.cursor.execute(f"SELECT codigo FROM materiais WHERE codigo = '{codigo}'")
        data = self.bd_param.cursor.fetchone()

        try:
            if data:
                self.bd_param.cursor.execute(f"UPDATE materiais SET codigo = {codigo}, descricao = '{descricao}', "
                                             f"qtd = '{quantidade}', card = '{card}', giro = '{giro}' "
                                             f"WHERE codigo = '{codigo}'")

            else:
                self.bd_param.cursor.execute("INSERT INTO materiais (codigo, descricao, qtd, card, giro) "
                                             "VALUES (?, ?, ?, ?, ?)", (codigo, descricao, quantidade, card, giro))

            self.bd_param.connect.commit()

            messagebox.showinfo(self.titulo, 'Cadastro de material atualizado com sucesso!')

        except:
            messagebox.showerror(self.titulo, 'Falha ao atualizar cadastro de material.')

        self.bd_param.connect.close()
        self.cancelar()

    def update(self, event=None):
        codigo = self.txt_codigo.get()

        self.bd_param.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{codigo}'")
        data = self.bd_param.cursor.fetchone()

        if data:
            self.txt_quantidade.delete(0, END)
            self.txt_quantidade.insert(0, data[3])
            self.txt_descricao.delete(0, END)
            self.txt_descricao.insert(0, data[2])
            self.cb_card.set(data[4])
            self.cb_giro.set(data[5])

    def cancelar(self):
        self.bd_param.connect.close()
        self.window.destroy()


# Cadastrar novo fornecedor
class CadastroFornecedor:
    def __init__(self, master):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-cad_fornecedor.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_salvar = PhotoImage(file=r'assets/images/btn-salvar.png')

        # Conectando ao banco de dados de parâmetros
        self.bd_param = BancoParametrosRede()

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title("TPC PLH | Framework IN - Cadastrar Fornecedor")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Entrybox para digitar o nome do fornecedor
        self.txt_fornecedor = Entry(self.window, font=self.fonte, justify=CENTER, background="#F2F2F2", border=False)
        self.txt_fornecedor.place(width=240, height=35, x=80, y=274)

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        # Combobox com as opções de tipo de fornecedor
        self.cb_tipo = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(), justify=CENTER)
        self.cb_tipo.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_tipo['values'] = ('BASE', 'LOJA', 'PARCEIRA')
        self.cb_tipo['state'] = 'readonly'
        self.cb_tipo.place(width=200, height=35, x=100, y=355)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para salvar
        self.btn_salvar = Button(self.window, image=self.img_salvar, border=False, command=self.cadastrar,
                                 activebackground='#FFFFFF')
        self.btn_salvar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Entrando na caixa de texto de código do material
        self.txt_fornecedor.focus_set()

        # Manter a janela de cadastro ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de cadastro estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de cadastro
        self.window.grab_release()

    def cadastrar(self):
        # Verificando se o nome e tipo do fornecedor foi informado
        if len(self.txt_fornecedor.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o nome do fornecedor.')
            return self.txt_fornecedor.focus_set()
        elif len(self.cb_tipo.get()) == 0:
            return messagebox.showerror(self.titulo, 'Selecione o tipo de fornecedor.')

        # Atribuíndo o nome e tipo do fornecedor em variáveis
        fornecedor = self.txt_fornecedor.get().strip()
        fornecedor = fornecedor.upper()  # Transformando o nome inteiro em letras maiúsculas
        tipo = self.cb_tipo.get()

        # Verificando a existência de um cadastro com o nome informado
        self.bd_param.cursor.execute(f"SELECT nome FROM fornecedores WHERE nome = '{fornecedor}'")
        if self.bd_param.cursor.fetchone() is not None:
            messagebox.showerror(self.titulo, 'Já existe um fornecedor cadastrado com o nome informado.')
            self.txt_fornecedor.delete(0, END)
            return self.txt_fornecedor.focus_set()

        try:
            # Inserindo cadastro na tabela de parâmetros
            self.bd_param.cursor.execute("INSERT INTO fornecedores (nome, tipo) VALUES (?, ?)", (fornecedor, tipo))
            self.bd_param.connect.commit()
            self.bd_param.connect.close()

            # Transformando o nome em letras minúsculas e retirando caracteres como espaços, pontos e parenteses
            fornecedor = fornecedor.lower()
            fornecedor = fornecedor.replace(' ', '_').replace('.', '')
            fornecedor = fornecedor.replace('(', '').replace(')', '')

            # Conectando no banco de dados UREV local e criando a tabela do fornecedor
            bd_local = BancoUrevLocal()
            bd_local.create_table(fornecedor)
            bd_local.connect.close()

            # Conectando no banco de dados UREV da rede e criando a tabela do fornecedor
            bd_rede = BancoUrevRede()
            bd_rede.create_table(fornecedor)
            bd_rede.connect.close()

            self.txt_fornecedor.delete(0, END)
            self.txt_fornecedor.focus_set()

            messagebox.showinfo(self.titulo, 'Fornecedor cadastrado com sucesso!')

        except:
            messagebox.showerror(self.titulo, 'Falha no cadastro do fornecedor.')

    def cancelar(self):
        self.bd_param.connect.close()
        self.window.destroy()


# Selecionar fornecedor, para salvar packlist ou iniciar leitura
class SelecionarFornecedor:
    def __init__(self, master):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão
        self.fornecedor = ''

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-selec_fornecedor.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_continuar = PhotoImage(file=r'assets/images/btn-continuar.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 450)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title("TPC PLH | Framework IN - Selecionar Fornecedor")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Conectando ao banco de dados e consultando os fornecedores cadastrados
        self.bd_param = BancoParametrosRede()
        self.bd_param.cursor.execute(f"SELECT * FROM fornecedores ORDER BY nome ASC")
        data = self.bd_param.cursor.fetchall()
        self.bd_param.connect.close()
        lista = []  # Lista será usada para armazar os nomes dos fornecedores

        # Adicionando os fornecedores na lista
        for i in range(0, len(data)):
            lista.append(data[i][1])

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        # Combobox com a lista de fornecedores cadastrados
        self.cb_fornecedor = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(),
                                          justify=CENTER)
        self.cb_fornecedor.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_fornecedor['values'] = lista
        self.cb_fornecedor['state'] = 'readonly'
        self.cb_fornecedor.place(width=240, height=35, x=80, y=255)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=370)

        # Botão para confirmar a escolha do fornecedor
        self.btn_continuar = Button(self.window, image=self.img_continuar, border=False, command=self.continuar,
                                    activebackground='#FFFFFF')
        self.btn_continuar.place(width=150, height=35, x=220, y=370)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=425)

        # Manter a janela de seleção ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de seleção estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de seleção
        self.window.grab_release()

    def continuar(self):
        if len(self.cb_fornecedor.get()) == 0:
            messagebox.showerror(self.titulo, 'Selecione o nome do fornecedor.')
            return

        self.fornecedor = self.cb_fornecedor.get()
        self.cancelar()
        return self.fornecedor

    def cancelar(self):
        self.bd_param.connect.close()
        self.window.destroy()


# Leitura de seriais
class Leitura:
    def __init__(self, master, nf, total_seriais, tipo_leitura, bancada, user, lote, perfil, fornecedor=None):
        # Parâmetros gerais
        self.titulo = 'TPC PLH | Framework IN'
        self.nf = nf
        self.nota = nf[3:]
        self.tipo_leitura = tipo_leitura
        self.bancada = bancada
        self.user = user
        self.lote = lote
        self.perfil = perfil
        self.fornecedor = fornecedor
        self.param_local = BancoParametrosLocal()

        # Definição da fonte padrão
        self.fonte_txt = ("Calibri", 14)
        self.fonte_lbl = ("Calibri", 16)
        self.fonte_lista = ['Calibri', 16, 'normal']  # Fonte usada na TreeView

        # Importar áudios
        self.assinante = r'assets\audio\serial_assinante.mp3'
        self.duplicado = r'assets\audio\serial_duplicado.mp3'
        self.invalido = r'assets\audio\serial_invalido.mp3'
        self.reclassificacao = r'assets\audio\serial_reclassificacao.mp3'

        # Importar imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-leitura.png')
        self.img_apagar_serial = PhotoImage(file=r'assets/images/btn-apagar_serial.png')
        self.img_apagar_caixa = PhotoImage(file=r'assets/images/btn-apagar_caixa.png')
        self.img_imprimir_qr = PhotoImage(file=r'assets/images/btn-imprimir_qr.png')
        self.img_salvar_pallet = PhotoImage(file=r'assets/images/btn-salvar_pallet.png')
        self.img_voltar = PhotoImage(file=r'assets/images/btn-voltar.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(1280, 720)

        # Configurações principais da janela
        self.window = master
        self.window.title("TPC PLH | Framework IN - Leitura de Seriais")
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')
        self.frame = Frame(master)
        self.frame.pack()

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.frame, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Total de seriais na nota fiscal
        self.lbl_totSeriais = Label(text=total_seriais, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_totSeriais.place(width=85, height=30, x=541, y=72)

        # Tipo de leitura (TRANSFERÊNCIA / REVERSA / LABORATÓRIO / ASSINANTES)
        self.lbl_tipoLeitura = Label(text=self.tipo_leitura, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_tipoLeitura.place(width=140, height=30, x=730, y=72)

        # Número da bancada
        self.lbl_bancada = Label(text=self.bancada, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_bancada.place(width=60, height=30, x=958, y=72)

        # Nome do usuário
        self.lbl_usuario = Label(text=self.user, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_usuario.place(width=120, height=30, x=1093, y=72)

        # Código do lote
        self.lbl_lote = Label(text=self.lote, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_lote.place(width=150, height=30, x=1030, y=124)

        # Entry box para código do serial
        self.txt_serial = Entry(font=self.fonte_txt, background="#F2F2F2", border=False, justify=CENTER)
        self.txt_serial.place(width=177, height=35, x=60, y=299)

        # Ajustes com base no tipo de leitura
        if self.tipo_leitura == 'TRANSFERENCIA':
            self.deposito = 'USAD'
            self.txt_serial.bind("<Return>", self.add_usad)  # Evento para quando apertar 'enter'
            self.bd_local = BancoUsadLocal()
            self.bd_local.cursor.execute(f"SELECT DISTINCT material FROM {self.nf}")
            data = self.bd_local.cursor.fetchall()
            lista = []

            for i in range(0, len(data)):
                lista.append(data[i][0])

            # Organizar a lista em ordem númerica
            lista = sorted(lista)

            # Combobox com as opções de tipo de fornecedor
            self.cb_materiais = ttk.Combobox(font=self.fonte_txt, width=120, textvariable=StringVar(), justify=CENTER)
            self.cb_materiais.option_add('*TCombobox*Listbox.Justify', 'center')
            self.cb_materiais['values'] = lista
            self.cb_materiais['state'] = 'readonly'
            self.cb_materiais.bind('<<ComboboxSelected>>', self.update_info)
            self.cb_materiais.place(width=120, height=33, x=88, y=203)

        else:
            # Definição de StringVar para as informações que serão alteradas dinamicamente
            self.cod_material = StringVar()

            if self.tipo_leitura == 'LABORATORIO':
                self.deposito = 'USUS'
                self.bd_local = BancoUsusLocal()
                self.cod_material.set(self.nota)
                self.txt_serial.bind("<Return>", self.add_usus)  # Evento para quando apertar 'enter'

            elif self.tipo_leitura == 'REVERSA':
                self.deposito = 'UREV'
                self.bd_local = BancoUrevLocal()
                self.txt_serial.bind("<Return>", self.add_urev)  # Evento para quando apertar 'enter'
                self.nota = nf

            elif self.tipo_leitura == 'ASSINANTES':
                self.deposito = 'UASS'
                self.bd_local = BancoUassLocal()
                self.txt_serial.bind("<Return>", self.add_uass)  # Evento para quando apertar 'enter'

                # Definindo o número da caixa
                self.bd_local.cursor.execute(f"SELECT DISTINCT caixa FROM assinantes")
                caixas = self.bd_local.cursor.fetchall()
                if not caixas:
                    self.caixa = 1
                else:
                    self.caixa = len(caixas)

            # Label com o código do último material lido
            self.lbl_material = Label(textvariable=self.cod_material, font=self.fonte_lbl, background='#FFFFFF')
            self.lbl_material.place(width=120, height=33, x=88, y=203)

        # Número da nota fiscal da leitura
        self.lbl_nf = Label(text=self.nota, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_nf.place(width=95, height=30, x=326, y=72)

        # Quantidade padrão por caixa do material
        self.lbl_padraoCX = Label(font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_padraoCX.place(width=60, height=30, x=333, y=202)

        # Número da caixa atual do material
        self.lbl_cxAtual = Label(font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_cxAtual.place(width=60, height=30, x=513, y=202)
        if self.tipo_leitura == 'ASSINANTES':
            self.lbl_cxAtual['text'] = self.caixa

        # Quantidade bipada na caixa atual
        self.lbl_bipadoCX = Label(font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_bipadoCX.place(width=60, height=30, x=694, y=202)

        # Código do card interno do material (apenas se houver)
        self.lbl_codCard = Label(font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_codCard.place(width=110, height=30, x=308, y=303)

        # Quantidade de cards interno bipado (apenas se houver)
        self.lbl_qtdCard = Label(font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_qtdCard.place(width=60, height=30, x=513, y=303)

        # Quantidade total de seriais bipados
        self.lbl_totBipado = Label(font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_totBipado.place(width=80, height=30, x=684, y=303)

        # Configurando os estilos e cores da TreeView
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('Treeview', background='#F2F2F2', fieldbackground='#F2F2F2', rowheight=25,
                        font=self.fonte_lista)
        style.map('Treeview')

        # Criar e inserir a TreeView de seriais bipados
        self.lista_bipados = ttk.Treeview(columns=('col1', 'col2', 'col3', 'col4'), selectmode='browse')
        self.lista_bipados.place(width=799, height=277, x=49, y=400)

        # Criar e inserir a barra de rolagem na lista de seriais bipados
        self.scroll_bipados = ttk.Scrollbar(orient='vertical', command=self.lista_bipados.yview)
        self.scroll_bipados.place(width=16, height=255 + 20, x=629 + 200 + 2, y=401)
        self.lista_bipados.configure(yscrollcommand=self.scroll_bipados.set)

        # Cabeçalho da TreeView
        self.lista_bipados.heading('#0', text='ID')
        self.lista_bipados.heading('#1', text='Material')
        self.lista_bipados.heading('#2', text='Serial')
        self.lista_bipados.heading('#3', text='Tecnologia')
        self.lista_bipados.heading('#4', text='Caixa')
        self.lista_bipados.config(show='tree')  # Configuração pra ocultar o cabeçalho da TreeView

        # Largura das colunas da TreeView
        self.lista_bipados.column('#0', width=1, stretch=False)  # Configuração pra ocultar a primeira coluna (ID)
        self.lista_bipados.column('#1', width=80, anchor='center')
        self.lista_bipados.column('#2', width=150, anchor='center')
        self.lista_bipados.column('#3', width=350, anchor='center')
        self.lista_bipados.column('#4', width=50, anchor='center')

        # Criando e inserindo a TreeView de consolidação da nota fiscal
        self.lista_nf = ttk.Treeview(columns=('col1', 'col2', 'col3', 'col4'), selectmode='browse')
        self.lista_nf.place(width=350, height=352, x=878, y=200)

        # Inserindo a barra de rolagem na lista da nota fiscal
        self.scroll_nf = ttk.Scrollbar(orient='vertical', command=self.lista_nf.yview)
        self.scroll_nf.place(width=16, height=330 + 20, x=1009 + 200 + 2, y=201)
        self.lista_nf.configure(yscrollcommand=self.scroll_nf.set)

        # Cabeçalho da TreeView
        self.lista_nf.heading('#0', text='ID')
        self.lista_nf.heading('#1', text='MATERIAL')
        self.lista_nf.heading('#2', text='NF')
        self.lista_nf.heading('#3', text='LIDO')
        self.lista_nf.heading('#4', text='REST')
        self.lista_nf.config(show='tree')  # Configuração pra ocultar o cabeçalho da TreeView

        # Largura das colunas da TreeView
        self.lista_nf.column('#0', width=1, stretch=False)  # Configuração pra ocultar a primeira coluna (ID)
        self.lista_nf.column('#1', width=114, anchor='center')
        self.lista_nf.column('#2', width=76, anchor='center')
        self.lista_nf.column('#3', width=76, anchor='center')
        self.lista_nf.column('#4', width=76, anchor='center')

        # Botão para apagar o serial selecionado
        self.apagar_serial = Button(image=self.img_apagar_serial, border=False, command=self.apagar_serial,
                                    activebackground='#FFFFFF')
        self.apagar_serial.place(width=165, height=35, x=878, y=560)

        # Botão para apagar uma caixa completa
        self.apagar_caixa = Button(image=self.img_apagar_caixa, border=False, command=self.apagar_caixa,
                                   activebackground='#FFFFFF')
        self.apagar_caixa.place(width=165, height=35, x=878, y=602)

        # Botão para gerar QRCODE
        self.imprimir_qr = Button(image=self.img_imprimir_qr, border=False, command=self.imprimir_qrcode,
                                  activebackground='#FFFFFF')
        self.imprimir_qr.place(width=165, height=35, x=1063, y=560)

        # Botão para salvar o pallet atual e gerar novo número de lote
        self.salvar_plt = Button(image=self.img_salvar_pallet, border=False, command=self.salvar_pallet,
                                 activebackground='#FFFFFF')
        self.salvar_plt.place(width=165, height=35, x=1063, y=602)

        # Botão para voltar para o menu principal
        self.voltar_menu = Button(image=self.img_voltar, border=False, command=self.voltar, activebackground='#FFFFFF')
        self.voltar_menu.place(width=150, height=35, x=1078, y=647)

        # Definir horário de inicio da leitura
        hora_atual = datetime.now()
        hora_atual = hora_atual.strftime('%H%M')

        # Label com horário de ínicio em área não visível
        self.lbl_horario = Label(text=hora_atual, font=self.fonte_lbl, background="#FFFFFF", justify=CENTER)
        self.lbl_horario.place(width=1, height=1, x=20, y=20)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=1110, y=695)

        # Definindo o usuário e senha de acesso ao Atlas
        self.bd_user = BancoUsuarios()
        self.bd_user.cursor.execute(f"SELECT nome, senha FROM usuarios WHERE perfil = 'Atlas'")
        data = self.bd_user.cursor.fetchone()
        self.bd_user.connect.close()
        self.atlas_user = data[0]
        self.atlas_password = data[1]

        # Definindo os parâmetros para iniciar o Chrome de forma oculta
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        self.chrome = webdriver.Chrome(options=options)

        # Criar conexão com o Atlas
        self.connect_atlas(self.chrome)

        # Atualizar informações da leitura
        self.update_info()

    # Criar conexão com o Atlas
    def connect_atlas(self, driver):
        try:
            # Abrindo a página do Atlas e fazendo login
            driver.get('http://atlas/nethome/index.do')
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="usuario"]').send_keys(self.atlas_user)
            driver.find_element(By.XPATH, '//*[@id="senha"]').send_keys(self.atlas_password)
            driver.find_element(By.XPATH, '//*[@id="base"]').send_keys('PALHOCA')
            driver.find_element(By.XPATH, '//*[@id="tableLogin"]/tbody/tr[3]/td/table/tbody/tr/td/a[1]').click()

            # Aguarde dois segundos antes de entrar na página de consulta de logins, afim de evitar erros
            time.sleep(2)

            # Entrando na página de consulta por seriais
            driver.get('http://atlas/nethome/equipamento/relatorioEquipamentos.do?acao=prepareSearch')
            driver.find_element(By.XPATH, '//*[@id="ui-id-1"]').click()

        except:
            driver.close()
            return messagebox.showerror(self.titulo,
                                        'Falha na conexão com o Atlas.\n\n'
                                        'Volte ao menu e abra a leitura novamente.')

    # Consultar serial no Atlas
    def serial_atlas(self, driver, serial):
        try:
            # Informando o número do serial e clicando em consultar
            driver.find_element(By.XPATH, '//*[@id="tabs-2"]/table/tbody/tr[2]/td[2]/textarea').send_keys(serial)
            driver.find_element(By.XPATH, '//*[@id="btnConsultar"]').click()

            # Coletando as informações e armazenando em variáveis
            try:
                sn = driver.find_element(By.XPATH, '//*[@id="tabPorEquipamentoResult"]/table/tbody/tr/td[1]').text
            except:
                playsound(self.invalido, True)
                driver.find_element(By.XPATH, '//*[@id="btnClear"]').click()
                messagebox.showerror(self.titulo, f'O serial {serial} é inválido.')
                return self.txt_serial.delete(0, END)

            end = driver.find_element(By.XPATH, '//*[@id="tabPorEquipamentoResult"]/table/tbody/tr/td[2]').text
            mod = driver.find_element(By.XPATH, '//*[@id="tabPorEquipamentoResult"]/table/tbody/tr/td[4]').text
            cod = driver.find_element(By.XPATH, '//*[@id="tabPorEquipamentoResult"]/table/tbody/tr/td[5]').text
            est = driver.find_element(By.XPATH, '//*[@id="tabPorEquipamentoResult"]/table/tbody/tr/td[7]').text
            loc = driver.find_element(By.XPATH, '//*[@id="tabPorEquipamentoResult"]/table/tbody/tr/td[8]').text

            # Formatando os seriais primário e secundário
            serial_pri = end[:12]
            serial_sec = end[15:]

            # Verificando se o material tem classificação PBL
            if 'PBL' not in cod:
                cod = cod[10:]

            # Clicando no botão de limpar para tirar o serial do campo de consulta
            driver.find_element(By.XPATH, '//*[@id="btnClear"]').click()

            info = [cod, serial_pri, serial_sec, sn, mod, est, loc]
            return info

        except:
            messagebox.showerror(self.titulo,
                                 'Falha na conexão com o Atlas.\n\n'
                                 'Volte ao menu e abra a leitura novamente.')
            return self.txt_serial.delete(0, END)

    # Adicionar seriais na leitura de TRANSFERÊNCIA (USAD)
    def add_usad(self, event):
        # Verificar se o código do material foi informado
        if len(self.cb_materiais.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o código do material.')
            return self.txt_serial.delete(0, END)

        # Verificar se o serial foi informado
        if len(event.widget.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o serial.')
            return self.txt_serial.focus_set()

        # Formatando o serial, apagando quebras de linhas, espaços e deixando em letras maiúsculas
        box_serial = event.widget.get()
        box_serial = box_serial.replace('\n', '')
        box_serial = box_serial.strip()
        box_serial = box_serial.upper()

        # Verificar a quantidade de caracteres na caixa de serial
        qtd_caracteres = len(box_serial)

        # Verificar se a quantidade de caracteres corresponde à um QRCODE ou serial individual
        if qtd_caracteres > 24:
            if qtd_caracteres % 12 != 0:
                messagebox.showerror(self.titulo, 'QRCODE inválido! Faça a leitura individual.')
                return event.widget.delete(0, END)

            qtd_seriais = int(qtd_caracteres / 12)

        else:
            qtd_seriais = 1

        # Definindo as informações do material de acordo com o cabeçalho da interface
        cod_material = self.cb_materiais.get()
        cod_card = self.lbl_codCard.cget('text')
        qtd_padrao = int(self.lbl_padraoCX.cget('text'))
        caixa = int(self.lbl_cxAtual.cget('text'))

        # Verificando a existência do cadastro do material nos parâmetros
        self.param_local.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{cod_material}'")
        if not self.param_local.cursor.fetchone():
            messagebox.showerror(self.titulo,
                                 f'Material {cod_material} não cadastrado na base.\n\n'
                                 f'Solicite o cadastro à um ADM.')
            return self.txt_serial.delete(0, END)

        packlist = 'SIM'  # Valor padrão para definir se o serial está no packlist
        gerar_etq = False  # Valor padrão para definir se será impresso a etiqueta de QRCODE

        # Definir horário da leitura do serial
        agora = datetime.now()
        horario = agora.strftime('%d/%m %H:%M:%S')

        # Loop para definir o código do serial, de acordo com a quantidade identificada
        for index in range(0, qtd_seriais):
            if qtd_seriais > 1:
                serial = box_serial[:12]
            else:
                serial = box_serial

            # Fazer a leitura do serial
            try:
                # Pesquisando pelo serial no packlist
                while True:
                    # Procurando pelo serial primário
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.nf} WHERE serial_primario = '{serial}'")
                    data = self.bd_local.cursor.fetchone()
                    if data:
                        serial = data[2]
                        break

                    # Procurando pelo serial secundário
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.nf} WHERE serial_secundario = '{serial}'")
                    data = self.bd_local.cursor.fetchone()
                    if data:
                        serial = data[2]
                        break

                    # Procurando pelo serial number
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.nf} WHERE serial_number = '{serial}'")
                    data = self.bd_local.cursor.fetchone()
                    if data:
                        serial = data[2]
                        break

                    break

                # Se o serial estiver no packlist, faça validações do material e serial bipado
                if data:
                    # Verificar reclassificação
                    if data[1] != cod_material:
                        if cod_card != 'NÃO' and data[1] == cod_card or data[1] == '41001508':
                            cod_material = data[1]
                        else:
                            packlist = f'SIM (RECLASS.)'
                            playsound(self.reclassificacao, True)
                            messagebox.showerror(self.titulo, f'O serial {serial} pertence ao material {data[1]}.')

                    # Verificar assinante
                    if data[7] == 'ASSINANTE':
                        packlist = 'SIM (ASSIN.)'
                        playsound(self.assinante, True)

                    # Verificar serial duplicado
                    if data[9] is not None:
                        playsound(self.duplicado, True)
                        messagebox.showerror(self.titulo, f'O serial {serial} está duplicado!\n\n'
                                                          f'Material: {data[1]}\n'
                                                          f'Caixa: {data[10]}\n'
                                                          f'Bancada: {data[11]}\n'
                                                          f'Lote: {data[9]}')
                        return event.widget.delete(0, END)

                # Contagem de material e definição do número da caixa
                try:
                    i = 1
                    while True:
                        self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {self.nf} "
                                                     f"WHERE (material = '{cod_material}' AND lote = '{self.lote}') "
                                                     f"AND caixa = {i}")

                        cont_material = self.bd_local.cursor.fetchone()
                        cont_material = cont_material[0]

                        # Verificar se a quantidade já bipada, somada com o serial atual é igual à quantidade padrão
                        if cont_material + 1 == qtd_padrao:
                            if qtd_caracteres < 25:  # Se for bipado QRCODE, não será impresso outra etiqueta
                                gerar_etq = True

                        # Se a quantidade já bipada for igual à quantidade padrão, pule a contagem da próxima caixa
                        if cont_material == qtd_padrao:
                            i += 1
                        else:  # Se a quantidade já bipada for menor que a quantidade padrão, defina a caixa atual
                            caixa = i
                            break

                except:  # Se houver a contagem total for menor do que a quantidade padrão, defina caixa 1
                    caixa = 1

                # Se o serial for assinante, a caixa será definida como 0 para que não misture
                if packlist == 'SIM (ASSIN.)':
                    caixa = 0

                # Se o serial estiver no packlist, atualizar os seriais existentes no banco de dados
                if data:
                    self.bd_local.cursor.execute(f"UPDATE {self.nf} SET packlist = '{packlist}', lote = '{self.lote}', "
                                                 f"caixa = {caixa}, bancada = '{self.bancada}', "
                                                 f"usuario = '{self.user}', hora = '{horario}' "
                                                 f"WHERE serial_primario = '{serial}'")

                # Se não for encontrado no packlist, consulte o serial no Atlas
                else:
                    info = self.serial_atlas(self.chrome, serial)

                    if info is None:  # Se o serial for inválido
                        return event.widget.delete(0, END)

                    # Se o material informado for divergente do Atlas
                    elif cod_material != info[0]:
                        # Verificar se o material possui card interno, se corresponder com o card, ajuste o código
                        if cod_card != 'NÃO' and (info[0] == cod_card or info[0] == '41001508'):
                            cod_material = cod_card
                        # Caso contrário, defina como reclassificação
                        else:
                            packlist = 'NÃO (RECLASS.)'
                            playsound(self.reclassificacao, True)
                            messagebox.showerror(self.titulo, f'O serial {serial} pertence ao material {info[0]}.')

                    # Se o serial e código de material estiver correto, apenas defina o packlist como "NÃO"
                    else:
                        packlist = 'NÃO'

                    # Se o serial for assinante, a caixa será definida como 0 para que não misture
                    if info[6] == 'ASSINANTE':
                        caixa = 0
                        packlist = 'NÃO (ASSIN.)'
                        playsound(self.assinante, True)

                    # Atribuindo as informações do serial em variáveis
                    mat = info[0]
                    serial = info[1]
                    serial_sec = info[2]
                    serial_num = info[3]
                    tec = info[4]
                    est = info[5]
                    loc = info[6]

                    # Se o material não possuir serial secundário, será repetido o serial primário
                    if serial_sec == '':
                        serial_sec = serial

                    # Adicionar todas as informações do serial e leitura no banco de dados
                    self.bd_local.cursor.execute(f"INSERT INTO {self.nf} (material, serial_primario, "
                                                 f"serial_secundario, serial_number, tecnologia, estado, local, "
                                                 f"packlist, lote, caixa, bancada, usuario, hora) "
                                                 f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                                 (mat, serial, serial_sec, serial_num, tec, est, loc, packlist,
                                                  self.lote, caixa, self.bancada, self.user, horario))

                # Apagar o serial atual do box de seriais
                box_serial = box_serial[12:]

            except:  # Em caso de erro
                messagebox.showerror(self.titulo, 'Erro ao adicionar serial.')
                return self.txt_serial.delete(0, END)

        # Salvar as modificações no banco de dados
        self.bd_local.connect.commit()

        # Se a quantidade bipada atingir a quantidade padrão, gerar etiqueta de QRCODE
        if gerar_etq and len(packlist) == 3:
            # Buscar todos os seriais correspondentes ao código de material, lote e caixa
            self.bd_local.cursor.execute(f"SELECT serial_primario FROM {self.nf} "
                                         f"WHERE (material = {cod_material} AND lote = '{self.lote}') "
                                         f"AND caixa = '{caixa}'")
            lista = self.bd_local.cursor.fetchall()  # Armazenando os seriais em uma lista
            qtd_etq = len(lista)  # Quantidade de seriais que serão impressos
            seriais = ''  # Variável usada para concatenar os seriais para gerar o QRCODE

            # Concatenando os seriais com quebra de linha a cada serial
            for i in range(0, len(lista)):
                if i < len(lista) - 1:
                    seriais += f'{lista[i][0]}\n'
                else:
                    seriais += f'{lista[i][0]}'

            # Gerar a etiqueta de QRCODE
            GerarEtiqueta(cod_material, self.nota, self.bancada, self.lote, 'USAD', caixa, qtd_etq, seriais)

        # Apagar os valores da entrybox de seriais
        event.widget.delete(0, END)

        # Atualizar todas as informações da leitura
        self.update_info()

    # Adicionar seriais na leitura de REVERSA (UREV)
    def add_urev(self, event):
        # Verificar se o serial foi informado
        if len(event.widget.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o serial.')
            return self.txt_serial.focus_set()

        # Formatando o serial, apagando quebras de linhas, espaços e deixando em letras maiúsculas
        box_serial = event.widget.get()
        box_serial = box_serial.replace('\n', '')
        box_serial = box_serial.strip()
        box_serial = box_serial.upper()
        cod_material = ''  # Variável que irá armazenar o código do material

        # Verificar a quantidade de caracteres na caixa de serial
        qtd_caracteres = len(box_serial)

        # Verificar se a quantidade de caracteres corresponde à um QRCODE ou serial individual
        if qtd_caracteres > 24:
            if qtd_caracteres % 12 != 0:
                messagebox.showerror(self.titulo, 'QRCODE inválido! Faça a leitura individual.')
                return event.widget.delete(0, END)

            qtd_seriais = int(qtd_caracteres / 12)

        else:
            qtd_seriais = 1

        # Criação prévia de variáveis que serão usadas
        packlist = 'NÃO'  # Serial está no packlist ou não
        caixa = 1  # Número da caixa atual
        gerar_etq = False  # Definir se será impresso a etiqueta de QRCODE
        serial_sec = ''  # Serial secundário
        serial_num = ''  # Serial number
        tec = ''  # Tecnologia
        est = ''  # Estado
        loc = ''  # Local

        # Definir horário da leitura do serial
        agora = datetime.now()
        horario = agora.strftime('%d/%m %H:%M:%S')

        # Loop para definir o código do serial, de acordo com a quantidade identificada
        for index in range(0, qtd_seriais):
            if qtd_seriais > 1:
                serial = box_serial[:12]
            else:
                serial = box_serial

            # Fazer a leitura do serial
            try:
                # Pesquisando pelo serial no packlist
                while True:
                    # Procurando pelo serial primário
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.fornecedor} "
                                                 f"WHERE nota_fiscal = '{self.nf}' AND serial_primario = '{serial}'")
                    data = self.bd_local.cursor.fetchone()
                    if data:
                        serial = data[3]
                        break

                    # Procurando pelo serial secundário
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.fornecedor} "
                                                 f"WHERE nota_fiscal = '{self.nf}' AND serial_secundario = '{serial}'")
                    data = self.bd_local.cursor.fetchone()
                    if data:
                        serial = data[3]
                        break

                    # Procurando pelo serial number
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.fornecedor} "
                                                 f"WHERE nota_fiscal = '{self.nf}' AND serial_number = '{serial}'")
                    data = self.bd_local.cursor.fetchone()
                    if data:
                        serial = data[3]
                        break

                    break

                # Se o serial estiver no packlist, defina o código do material e verifique se está cadastrado
                if data:
                    packlist = 'SIM'
                    cod_material = data[2]
                    self.param_local.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{cod_material}'")
                    data_param = self.param_local.cursor.fetchone()

                    # Se não estiver cadastrado
                    if not data_param:
                        messagebox.showerror(self.titulo,
                                             f'Material {cod_material} não cadastrado na base.\n\n'
                                             f'Solicite o cadastro à um ADM.')
                        return self.txt_serial.delete(0, END)

                    # Definir quantidade padrão e tipo de giro do material
                    qtd_padrao = data_param[3]
                    tipo_giro = data_param[5]

                    # Verifique por assinante e tipo de giro, apenas materiais de alto giro serão destacados
                    if data[8] == 'ASSINANTE' and tipo_giro == 'ALTO GIRO':
                        caixa = 0  # Se o serial for assinante, a caixa será definida como 0 para que não misture
                        packlist = 'SIM (ASSIN.)'
                        playsound(self.assinante, True)

                # Se não for encontrado no packlist, consulte o serial no Atlas
                else:
                    info = self.serial_atlas(self.chrome, serial)

                    if info is None:  # Se o serial for inválido
                        return event.widget.delete(0, END)

                    # Armazenando informações do serial em variáveis
                    cod_material = info[0]
                    serial = info[1]
                    serial_sec = info[2]
                    serial_num = info[3]
                    tec = info[4]
                    est = info[5]
                    loc = info[6]
                    packlist = 'NÃO'

                    # Se o material não possuir serial secundário, será repetido o serial primário
                    if serial_sec == '':
                        serial_sec = serial

                    # Verificar se o material está cadastrado na base
                    self.param_local.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{cod_material}'")
                    data_param = self.param_local.cursor.fetchone()

                    # Se não estiver cadastrado
                    if not data_param:
                        messagebox.showerror(self.titulo,
                                             f'Material {cod_material} não cadastrado na base.\n\n'
                                             f'Solicite o cadastro à um ADM.')
                        return self.txt_serial.delete(0, END)

                    # Definir quantidade padrão e tipo de giro do material
                    qtd_padrao = data_param[3]
                    tipo_giro = data_param[5]

                    # Verifique por assinante e tipo de giro, apenas materiais de alto giro serão destacados
                    if loc == 'ASSINANTE' and tipo_giro == 'ALTO GIRO':
                        caixa = 0  # Se o serial for assinante, a caixa será definida como 0 para que não misture
                        packlist = 'NÃO (ASSIN.)'
                        playsound(self.assinante, True)

                # Pesquisar por serial duplicado
                self.bd_local.cursor.execute(f"SELECT * FROM {self.fornecedor} "
                                             f"WHERE nota_fiscal = '{self.nf}' AND serial_primario = '{serial}'")
                data = self.bd_local.cursor.fetchone()

                # Se o serial estiver duplicado
                if data:
                    if data[10] is not None:
                        playsound(self.duplicado, True)
                        messagebox.showerror(self.titulo, f'O serial {serial} está duplicado!\n\n'
                                                          f'Material: {data[2]}\n'
                                                          f'Caixa: {data[11]}\n'
                                                          f'Bancada: {data[12]}\n'
                                                          f'Lote: {data[10]}')
                        return event.widget.delete(0, END)

                # Contagem de material e definição do número da caixa
                if caixa != 0:
                    try:
                        i = 1
                        while True:
                            self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {self.fornecedor} "
                                                         f"WHERE (nota_fiscal = '{self.nf}' AND lote = '{self.lote}') "
                                                         f"AND (material = '{cod_material}' AND caixa = {i})")

                            cont_material = self.bd_local.cursor.fetchone()
                            cont_material = cont_material[0]

                            # Verificar se a quantidade já bipada, somada com o serial atual é igual à quantidade padrão
                            if cont_material + 1 == qtd_padrao:
                                if qtd_caracteres < 25:  # Se for bipado QRCODE, não será impresso outra etiqueta
                                    gerar_etq = True

                            # Se a quantidade já bipada for igual à quantidade padrão, pule a contagem da próxima caixa
                            if cont_material == qtd_padrao:
                                i += 1
                            else:  # Se a quantidade já bipada for menor que a quantidade padrão, defina a caixa atual
                                caixa = i
                                break

                    except:  # Se houver a contagem total for menor do que a quantidade padrão, defina caixa 1
                        caixa = 1

                # Se o serial estiver no packlist, atualizar os seriais existentes no banco de dados
                if data:
                    self.bd_local.cursor.execute(f"UPDATE {self.fornecedor} SET packlist = '{packlist}', "
                                                 f"lote = '{self.lote}', caixa = {caixa}, bancada = '{self.bancada}', "
                                                 f"usuario = '{self.user}', hora = '{horario}' "
                                                 f"WHERE nota_fiscal = '{self.nf}' AND serial_primario = '{serial}'")

                # Se não estiver no packlist, adicionar todas as informações do serial e leitura no banco de dados
                else:
                    # Adicionar todas as informações do serial e leitura no banco de dados
                    self.bd_local.cursor.execute(f"INSERT INTO {self.fornecedor} (nota_fiscal, material, "
                                                 f"serial_primario, serial_secundario, serial_number, tecnologia, "
                                                 f"estado, local, packlist, lote, caixa, bancada, usuario, hora) "
                                                 f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                                 (self.nf, cod_material, serial, serial_sec, serial_num, tec, est, loc,
                                                  packlist, self.lote, caixa, self.bancada, self.user, horario))

                # Apagar o serial atual do box de seriais
                box_serial = box_serial[12:]

            except:  # Em caso de erro
                messagebox.showerror(self.titulo, 'Erro ao adicionar serial.')
                return self.txt_serial.delete(0, END)

        # Salvar as modificações no banco de dados
        self.bd_local.connect.commit()

        # Se a quantidade bipada atingir a quantidade padrão, gerar etiqueta de QRCODE
        if gerar_etq and len(packlist) == 3:
            # Buscar todos os seriais correspondentes ao código de material, lote e caixa
            self.bd_local.cursor.execute(f"SELECT serial_primario FROM {self.fornecedor} "
                                         f"WHERE (nota_fiscal = '{self.nf}' AND material = {cod_material}) "
                                         f"AND (lote = '{self.lote}' AND caixa = '{caixa}')")
            lista = self.bd_local.cursor.fetchall()  # Armazenando os seriais em uma lista
            qtd_etq = len(lista)  # Quantidade de seriais que serão impressos
            seriais = ''  # Variável usada para concatenar os seriais para gerar o QRCODE

            # Concatenando os seriais com quebra de linha a cada serial
            for i in range(0, len(lista)):
                if i < len(lista) - 1:
                    seriais += f'{lista[i][0]}\n'
                else:
                    seriais += f'{lista[i][0]}'

            # Gerar a etiqueta de QRCODE
            GerarEtiqueta(cod_material, self.nota, self.bancada, self.lote, 'UREV', caixa, qtd_etq, seriais)

        # Atualizar a label com o código do material
        self.cod_material.set(cod_material)
        self.lbl_material['text'] = self.cod_material

        # Apagar os valores da entrybox de seriais
        event.widget.delete(0, END)

        # Atualizar todas as informações da leitura
        self.update_info()

    # Adicionar seriais na leitura de LABORATÓRIO (USUS)
    def add_usus(self, event):
        # Verificar se o serial foi informado
        if len(event.widget.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o serial.')
            return self.txt_serial.focus_set()

        # Formatando o serial, apagando quebras de linhas, espaços e deixando em letras maiúsculas
        box_serial = event.widget.get()
        box_serial = box_serial.replace('\n', '')
        box_serial = box_serial.strip()
        box_serial = box_serial.upper()

        # Verificar a quantidade de caracteres na caixa de serial
        qtd_caracteres = len(box_serial)

        # Verificar se a quantidade de caracteres corresponde à um QRCODE ou serial individual
        if qtd_caracteres > 24:
            if qtd_caracteres % 12 != 0:
                messagebox.showerror(self.titulo, 'QRCODE inválido! Faça a leitura individual.')
                return event.widget.delete(0, END)

            qtd_seriais = int(qtd_caracteres / 12)

        else:
            qtd_seriais = 1

        # Definindo as informações do material de acordo com o cabeçalho da interface
        cod_material = self.lbl_material.cget('text')
        cod_card = self.lbl_codCard.cget('text')
        qtd_padrao = int(self.lbl_padraoCX.cget('text'))
        caixa = int(self.lbl_cxAtual.cget('text'))

        # Verificando a existência do cadastro do material nos parâmetros
        self.param_local.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{cod_material}'")
        if not self.param_local.cursor.fetchone():
            messagebox.showerror(self.titulo,
                                 f'Material {cod_material} não cadastrado na base.\n\n'
                                 f'Solicite o cadastro à um ADM.')
            return self.txt_serial.delete(0, END)

        gerar_etq = False  # Valor padrão para definir se será impresso a etiqueta de QRCODE

        # Definir horário da leitura do serial
        agora = datetime.now()
        horario = agora.strftime('%d/%m %H:%M:%S')

        # Loop para definir o código do serial, de acordo com a quantidade identificada
        for index in range(0, qtd_seriais):
            if qtd_seriais > 1:
                serial = box_serial[:12]
            else:
                serial = box_serial

            # Fazer a leitura do serial
            try:
                info = self.serial_atlas(self.chrome, serial)

                if info is None:  # Se o serial for inválido
                    return event.widget.delete(0, END)

                # Se o material informado for divergente do Atlas
                elif cod_material != info[0]:
                    # Verificar se o material possui card interno, se corresponder com o card, ajuste o código
                    if cod_card != 'NÃO' and (info[0] == cod_card or info[0] == '41001508'):
                        cod_material = cod_card
                    # Caso contrário, defina como reclassificação
                    else:
                        playsound(self.reclassificacao, True)
                        messagebox.showerror(self.titulo, f'O serial {serial} pertence ao material {info[0]}.')

                # Atribuindo as informações do serial em variáveis
                mat = info[0]
                serial = info[1]
                serial_sec = info[2]
                serial_num = info[3]
                tec = info[4]
                est = info[5]
                loc = info[6]

                # Se o material não possuir serial secundário, será repetido o serial primário
                if serial_sec == '':
                    serial_sec = serial

                # Pesquisar por serial duplicado
                self.bd_local.cursor.execute(f"SELECT * FROM {self.nf} "
                                             f"WHERE serial_primario = '{serial}' AND lote = '{self.lote}'")
                data = self.bd_local.cursor.fetchone()

                # Se o serial estiver duplicado
                if data:
                    playsound(self.duplicado, True)
                    messagebox.showerror(self.titulo, f'O serial {serial} está duplicado!\n\n'
                                                      f'Material: {data[1]}\n'
                                                      f'Caixa: {data[9]}\n'
                                                      f'Bancada: {data[10]}\n'
                                                      f'Lote: {data[8]}')
                    return event.widget.delete(0, END)

                # Contagem de material e definição do número da caixa
                try:
                    i = 1
                    while True:
                        self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {self.nf} "
                                                     f"WHERE (material = '{cod_material}' AND lote = '{self.lote}') "
                                                     f"AND caixa = {i}")

                        cont_material = self.bd_local.cursor.fetchone()
                        cont_material = cont_material[0]

                        # Verificar se a quantidade já bipada, somada com o serial atual é igual à quantidade padrão
                        if cont_material + 1 == qtd_padrao:
                            if qtd_caracteres < 25:  # Se for bipado QRCODE, não será impresso outra etiqueta
                                gerar_etq = True

                        # Se a quantidade já bipada for igual à quantidade padrão, pule a contagem da próxima caixa
                        if cont_material == qtd_padrao:
                            i += 1
                        else:  # Se a quantidade já bipada for menor que a quantidade padrão, defina a caixa atual
                            caixa = i
                            break

                # Se houver a contagem total for menor do que a quantidade padrão, defina caixa 1
                except:
                    caixa = 1

                # Se o serial for assinante, a caixa será definida como 0 para que não misture
                if loc == 'ASSINANTE':
                    caixa = 0
                    playsound(self.assinante, True)

                # Adicionar todas as informações do serial e leitura no banco de dados
                self.bd_local.cursor.execute(f"INSERT INTO {self.nf} (material, serial_primario, "
                                             f"serial_secundario, serial_number, tecnologia, estado, local, "
                                             f"lote, caixa, bancada, usuario, hora) "
                                             f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                             (mat, serial, serial_sec, serial_num, tec, est, loc, self.lote, caixa,
                                              self.bancada, self.user, horario))

                # Apagar o serial atual do box de seriais
                box_serial = box_serial[12:]

            except:  # Em caso de erro
                messagebox.showerror(self.titulo, 'Erro ao adicionar serial.')
                return self.txt_serial.delete(0, END)

        # Salvar as modificações no banco de dados
        self.bd_local.connect.commit()

        # Se a quantidade bipada atingir a quantidade padrão, gerar etiqueta de QRCODE
        if gerar_etq:
            # Buscar todos os seriais correspondentes ao código de material, lote e caixa
            self.bd_local.cursor.execute(f"SELECT serial_primario FROM {self.nf} "
                                         f"WHERE (material = {cod_material} AND lote = '{self.lote}') "
                                         f"AND caixa = '{caixa}'")
            lista = self.bd_local.cursor.fetchall()  # Armazenando os seriais em uma lista
            qtd_etq = len(lista)  # Quantidade de seriais que serão impressos
            seriais = ''  # Variável usada para concatenar os seriais para gerar o QRCODE

            # Concatenando os seriais com quebra de linha a cada serial
            for i in range(0, len(lista)):
                if i < len(lista) - 1:
                    seriais += f'{lista[i][0]}\n'
                else:
                    seriais += f'{lista[i][0]}'

            # Gerar a etiqueta de QRCODE
            GerarEtiqueta(cod_material, 'LABORATORIO', self.bancada, self.lote, 'USUS', caixa, qtd_etq, seriais)

        # Apagar os valores da entrybox de seriais
        event.widget.delete(0, END)

        # Atualizar todas as informações da leitura
        self.update_info()

    # Adicionar seriais na leitura de ASSINANTES (UASS)
    def add_uass(self, event):
        # Verificar se o serial foi informado
        if len(event.widget.get()) == 0:
            messagebox.showerror(self.titulo, 'Informe o serial.')
            return self.txt_serial.focus_set()

        # Formatando o serial, apagando quebras de linhas, espaços e deixando em letras maiúsculas
        box_serial = event.widget.get()
        box_serial = box_serial.replace('\n', '')
        box_serial = box_serial.strip()
        serial = box_serial.upper()

        # Definir horário da leitura do serial
        agora = datetime.now()
        horario = agora.strftime('%d/%m %H:%M:%S')

        # Definir número da caixa atual
        caixa = self.lbl_cxAtual.cget('text')

        # Fazer a leitura do serial
        info = self.serial_atlas(self.chrome, serial)

        if info is None:  # Se o serial for inválido
            return event.widget.delete(0, END)

        # Armazenando informações do serial em variáveis
        cod_material = info[0]
        serial = info[1]
        serial_sec = info[2]
        serial_num = info[3]
        tec = info[4]
        est = info[5]
        loc = info[6]

        # Verificar se o material está cadastrado na base
        self.param_local.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{cod_material}'")
        data_param = self.param_local.cursor.fetchone()

        # Se não estiver cadastrado
        if not data_param:
            messagebox.showerror(self.titulo,
                                 f'Material {cod_material} não cadastrado na base.\n\n'
                                 f'Solicite o cadastro à um ADM.')
            return self.txt_serial.delete(0, END)

        # Se o material não possuir serial secundário, será repetido o serial primário
        if serial_sec == '':
            serial_sec = serial

        # Pesquisar por serial duplicado
        self.bd_local.cursor.execute(f"SELECT * FROM assinantes "
                                     f"WHERE serial_primario = '{serial}' AND lote = '{self.lote}'")
        data = self.bd_local.cursor.fetchone()

        if data:
            playsound(self.duplicado, True)
            messagebox.showerror(self.titulo, f'O serial {serial} está duplicado!\n\n'
                                              f'Material: {data[1]}\n'
                                              f'Caixa: {data[9]}\n'
                                              f'Bancada: {data[10]}\n'
                                              f'Lote: {data[8]}')
            return event.widget.delete(0, END)

        # Adicionar todas as informações do serial e leitura no banco de dados
        try:
            self.bd_local.cursor.execute(f"INSERT INTO assinantes (material, serial_primario, serial_secundario, "
                                         f"serial_number, tecnologia, estado, local, lote, caixa, bancada, "
                                         f"usuario, hora) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                         (cod_material, serial, serial_sec, serial_num, tec, est, loc, self.lote, caixa,
                                          self.bancada, self.user, horario))

            # Salvar as modificações no banco de dados
            self.bd_local.connect.commit()

        except:  # Em caso de erro
            messagebox.showerror(self.titulo, 'Erro ao adicionar serial.')
            return self.txt_serial.delete(0, END)

        # Atualizar a label com o código do material
        self.cod_material.set(cod_material)
        self.lbl_material['text'] = self.cod_material

        # Apagar os valores da entrybox de seriais
        event.widget.delete(0, END)

        # Atualizar todas as informações da leitura
        self.update_info()

    # Atualizar as informações
    def update_info(self, event=None):
        # Definir tabela para buscar as informações no banco de dados
        if self.tipo_leitura == 'REVERSA':
            tabela = self.fornecedor
        elif self.tipo_leitura == 'ASSINANTES':
            tabela = 'assinantes'
        else:
            tabela = self.nf

        # Definir código do material
        if self.tipo_leitura == 'TRANSFERENCIA':
            cod_material = self.cb_materiais.get()
        else:
            cod_material = self.lbl_material.cget('text')

        # Definir quantidade padrão por caixa do material
        if len(cod_material) == 0:
            qtd_padrao = '-'
            cod_card = '-'

        else:
            self.param_local.cursor.execute(f"SELECT * FROM materiais WHERE codigo = '{cod_material}'")
            data = self.param_local.cursor.fetchone()
            qtd_padrao = data[3]
            cod_card = data[4]

        # Contagem da quantidade de materiais
        try:
            i = 1
            if self.tipo_leitura == 'TRANSFERENCIA':
                while True:
                    self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {tabela} "
                                                 f"WHERE (material = '{cod_material}' AND lote = '{self.lote}') "
                                                 f"AND caixa = {i}")

                    cont_material = self.bd_local.cursor.fetchone()
                    cont_material = cont_material[0]

                    if cont_material == qtd_padrao:
                        i += 1
                    else:
                        caixa = i
                        break

            elif self.tipo_leitura == 'LABORATORIO':
                while True:
                    self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {tabela} "
                                                 f"WHERE material = '{cod_material}' AND lote = '{self.lote}' "
                                                 f"AND caixa = {i}")

                    cont_material = self.bd_local.cursor.fetchone()
                    cont_material = cont_material[0]

                    if cont_material == qtd_padrao:
                        i += 1
                    else:
                        caixa = i
                        break

            elif self.tipo_leitura == 'REVERSA':
                while True:
                    self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {tabela} "
                                                 f"WHERE (nota_fiscal = '{self.nf}' AND material = '{cod_material}') "
                                                 f"AND (lote = '{self.lote}' AND caixa = {i})")

                    cont_material = self.bd_local.cursor.fetchone()
                    cont_material = cont_material[0]

                    if cont_material == qtd_padrao:
                        i += 1
                    else:
                        caixa = i
                        break

            else:
                self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {tabela} "
                                             f"WHERE material = '{cod_material}' AND lote = '{self.lote}' "
                                             f"AND caixa = {self.caixa}")
                cont_material = self.bd_local.cursor.fetchone()
                cont_material = cont_material[0]
                caixa = 1

        except:
            cont_material = 0
            caixa = 1

        # Verificar se o material possui card
        if cod_card == 'NÃO' or cod_card == '-':
            self.lbl_qtdCard['text'] = '-'

        else:
            # Contagem da quantidade de cards
            try:
                if self.tipo_leitura == 'REVERSA':
                    query_01 = f"SELECT COUNT(*) FROM {tabela} " \
                               f"WHERE (nota_fiscal = '{self.nf}' AND material = '{cod_card}') " \
                               f"AND (lote = '{self.lote}' AND caixa = '{caixa}')"

                elif self.tipo_leitura == 'ASSINANTES':
                    caixa = self.lbl_cxAtual.cget('text')
                    query_01 = f"SELECT COUNT(*) FROM {tabela} " \
                               f"WHERE (material = '{cod_card}' AND lote = '{self.lote}') AND caixa = '{caixa}'"

                else:
                    query_01 = f"SELECT COUNT(*) FROM {tabela} " \
                               f"WHERE (material = '{cod_card}' AND lote = '{self.lote}') AND caixa = '{caixa}'"

                # Fazer a contagem no banco de dados
                self.bd_local.cursor.execute(query_01)
                qtd_card = self.bd_local.cursor.fetchone()
                qtd_card = qtd_card[0]

            except:
                qtd_card = 0

            self.lbl_qtdCard['text'] = qtd_card

        # Contagem de total de materiais bipados
        try:
            if self.tipo_leitura == 'REVERSA':
                query_02 = f"SELECT COUNT(*) FROM {tabela} WHERE nota_fiscal = '{self.nf}' AND lote = '{self.lote}'"
            else:
                query_02 = f"SELECT COUNT(*) FROM {tabela} WHERE lote = '{self.lote}'"

            self.bd_local.cursor.execute(query_02)
            tot_bipado = self.bd_local.cursor.fetchone()
            tot_bipado = tot_bipado[0]

        except:
            tot_bipado = 0

        # Atualizar cabeçalho de informações da leitura
        self.lbl_padraoCX['text'] = qtd_padrao
        self.lbl_codCard['text'] = cod_card
        self.lbl_cxAtual['text'] = caixa
        self.lbl_bipadoCX['text'] = cont_material
        self.lbl_totBipado['text'] = tot_bipado

        # Comando para buscar todos os seriais lidos
        if self.tipo_leitura == 'REVERSA':
            query_03 = f"SELECT * FROM {tabela} " \
                       f"WHERE nota_fiscal = '{self.nf}' AND lote = '{self.lote}' ORDER BY hora ASC"

        else:
            query_03 = f"SELECT * FROM {tabela} WHERE lote = '{self.lote}' ORDER BY hora ASC"

        # Executar o comando no banco de dados
        self.bd_local.cursor.execute(query_03)
        bipagem = self.bd_local.cursor.fetchall()

        # Apagar a lista atual
        self.lista_bipados.delete(*self.lista_bipados.get_children())

        # Loop para adicionar os seriais encontrados
        for i in range(0, len(bipagem)):
            tag = 'default'

            if self.tipo_leitura == 'LABORATORIO':
                cod_material = self.lbl_material.cget('text')
                # Verificar se o código de material e código de card está preenchido
                if cod_material != '-' and cod_card != '-':
                    # Se o código for divergente do material ou do card, definir como reclassificação
                    if bipagem[i][1] != cod_material and bipagem[i][1] != cod_card:
                        tag = 'reclassificacao'

                # Se o material estiver em assinante
                if bipagem[i][7] == 'ASSINANTE':
                    tag = 'assinante'

                # Lista de valores para adicionar na TreeView
                valores = (f'{bipagem[i][1]}', f'{bipagem[i][2]}', f'{bipagem[i][5]}', f'{bipagem[i][9]}')

            elif self.tipo_leitura == 'REVERSA':
                # Se o material estiver em assinante
                if 'ASSIN' in bipagem[i][9]:
                    tag = 'assinante'

                # Lista de valores para adicionar na TreeView
                valores = (f'{bipagem[i][2]}', f'{bipagem[i][3]}', f'{bipagem[i][6]}', f'{bipagem[i][11]}')

            elif self.tipo_leitura == 'ASSINANTES':
                # Se o material não estiver em assinante, definir como "OK"
                if bipagem[i][7] != 'ASSINANTE':
                    tag = 'ok'

                # Lista de valores para adicionar na TreeView
                valores = (f'{bipagem[i][1]}', f'{bipagem[i][2]}', f'{bipagem[i][5]}', f'{bipagem[i][9]}')

            else:
                # Se o código for divergente do material ou do card, definir como reclassificação
                if 'RECLASS.' in bipagem[i][8]:
                    tag = 'reclassificacao'

                # Se o material estiver em assinante
                if bipagem[i][7] == 'ASSINANTE':
                    tag = 'assinante'

                # Lista de valores para adicionar na TreeView
                valores = (f'{bipagem[i][1]}', f'{bipagem[i][2]}', f'{bipagem[i][5]}', f'{bipagem[i][10]}')

            # Criando e configurando tags para destacar as cores das linhas
            self.lista_bipados.tag_configure('default', background='#F2F2F2', foreground='black')
            self.lista_bipados.tag_configure('assinante', background='#C00000', foreground='white')
            self.lista_bipados.tag_configure('reclassificacao', background='#FFD966', foreground='black')
            self.lista_bipados.tag_configure('ok', background='#00B050', foreground='black')

            # Inserindo os seriais na lista
            self.lista_bipados.insert('', END, values=valores, tags=tag)
            self.lista_bipados.see(self.lista_bipados.get_children()[-1])

        # Comando para consolidar os materiais e quantidades da nota fiscal
        if self.tipo_leitura == 'REVERSA':
            query_04 = f"SELECT material, COUNT(material) FROM {self.fornecedor} " \
                       f"WHERE nota_fiscal = '{self.nf}' AND packlist LIKE 'SIM%' GROUP BY material"

        elif self.tipo_leitura == 'ASSINANTES':
            query_04 = f"SELECT material, COUNT(material) FROM assinantes WHERE lote = '{self.lote}' GROUP BY material"

        elif self.tipo_leitura == 'TRANSFERENCIA':
            query_04 = f"SELECT material, COUNT(material) FROM {self.nf} WHERE packlist LIKE 'SIM%' GROUP BY material"

        else:
            query_04 = f"SELECT material, COUNT(material) FROM {self.nf} GROUP BY material"

        # Executando o comando no banco de dados
        self.bd_local.cursor.execute(query_04)
        consolidado = self.bd_local.cursor.fetchall()

        # Apagar a lista atual
        self.lista_nf.delete(*self.lista_nf.get_children())

        # Loop para adicionar os materiais e quantidades
        for i in range(0, len(consolidado)):
            material = consolidado[i][0]
            qtd_nf = consolidado[i][1]

            # Comando para fazer a contagem de materiais bipados
            if self.tipo_leitura == 'REVERSA':
                query_05 = f"SELECT COUNT(*) FROM {self.fornecedor} " \
                           f"WHERE (nota_fiscal = '{self.nf}' AND material = '{material}') AND lote = '{self.lote}'"

            elif self.tipo_leitura == 'ASSINANTES':
                query_05 = f"SELECT COUNT(*) FROM assinantes WHERE material = '{material}' AND lote = '{self.lote}'"

            else:
                query_05 = f"SELECT COUNT(*) FROM {self.nf} WHERE material = '{material}' AND lote = '{self.lote}'"

            # Executar o comando no banco de dados
            self.bd_local.cursor.execute(query_05)

            # Armazenar a quantidade em variável
            qtd_lido = self.bd_local.cursor.fetchone()
            qtd_lido = qtd_lido[0]

            # Definir a quantidade restante
            qtd_rest = qtd_nf - qtd_lido

            # Lista de valores para adicionar na TreeView
            info = (f'{material}', f'{qtd_nf}', f'{qtd_lido}', f'{qtd_rest}')

            # Criando e configurando tags para destacar as cores das linhas
            self.lista_nf.tag_configure('default', background='#F2F2F2', foreground='black')
            self.lista_nf.tag_configure('ok', background='#00B050', foreground='black')
            self.lista_nf.tag_configure('sobra', background='#FFD966', foreground='black')

            # Verificar a quantidade restante para destacar cor da linha
            if qtd_rest == 0:
                tag = 'ok'
            elif qtd_rest < 0:
                tag = 'sobra'
            else:
                tag = 'default'

            # Se o tipo de leitura for Laboratório ou Assinantes, sempre utilizar a tag "default"
            if self.tipo_leitura == 'LABORATORIO' or self.tipo_leitura == 'ASSINANTES':
                tag = 'default'

            # Inserindo as informações na lista
            self.lista_nf.insert('', END, values=info, tags=tag)

        # Definindo hora atual e hora de início da leitura
        hora_atual = datetime.now()
        hora_atual = int(hora_atual.strftime('%H%M'))
        hora_inicio = self.lbl_horario.cget('text')
        hora_inicio = int(hora_inicio)

        # Salvando produtividade a cada 10 minutos
        if hora_atual >= hora_inicio + 5:
            self.salvar_produtividade()
            self.lbl_horario['text'] = hora_atual

        # Selecionar foco na caixa de serial
        self.txt_serial.focus_set()

    # Imprimir QRCODE manual
    def imprimir_qrcode(self):
        # Variável para armazenar o resultado de busca de códigos de materiais
        data = None

        # Verificação de tipo de leitura
        if self.tipo_leitura == 'TRANSFERENCIA' or self.tipo_leitura == 'LABORATORIO':
            self.bd_local.cursor.execute(f"SELECT DISTINCT material FROM {self.nf} WHERE lote = '{self.lote}'")
            data = self.bd_local.cursor.fetchall()

        elif self.tipo_leitura == 'REVERSA':
            self.bd_local.cursor.execute(f"SELECT DISTINCT material FROM {self.fornecedor} "
                                         f"WHERE nota_fiscal = {self.nota} AND lote = '{self.lote}'")
            data = self.bd_local.cursor.fetchall()

        # Se a o resultado da buscar encontrar materiais, adicionar os códigos em uma lista
        if data is not None:
            materiais = []

            for i in range(0, len(data)):
                materiais.append(data[i][0])

        # Se não encontrar, definir a variável da lista como nada
        else:
            materiais = None

        info = get_material(self.tipo_leitura, materiais)
        if not info:
            return

        cod_material = info[0]
        caixa = info[1]

        if self.tipo_leitura == 'ASSINANTES':
            nota = 'ASSIN.'
            query = f"SELECT serial_primario FROM assinantes WHERE lote = '{self.lote}' AND caixa = '{caixa}'"

        elif self.tipo_leitura == 'TRANSFERENCIA':
            nota = self.nota
            query = f"SELECT serial_primario FROM {self.nf} " \
                    f"WHERE (material = {cod_material} AND lote = '{self.lote}') AND caixa = '{caixa}'"

        elif self.tipo_leitura == 'REVERSA':
            nota = self.nota
            query = f"SELECT serial_primario FROM {self.fornecedor} " \
                    f"WHERE (nota_fiscal = '{self.nf}' AND material = {cod_material}) " \
                    f"AND (lote = '{self.lote}' AND caixa = '{caixa}')"

        else:
            nota = 'LABORATORIO'
            query = f"SELECT serial_primario FROM {self.nf} " \
                    f"WHERE (material = {cod_material} AND lote = '{self.lote}') AND caixa = '{caixa}'"

        try:
            self.bd_local.cursor.execute(query)
            lista = self.bd_local.cursor.fetchall()

            # Verificar se existe seriais para geração do QRCODE
            if lista:
                seriais = ''
                qtd_etq = len(lista)
                for i in range(0, len(lista)):
                    if i < len(lista) - 1:
                        seriais += f'{lista[i][0]}\n'
                    else:
                        seriais += f'{lista[i][0]}'

                GerarEtiqueta(cod_material, nota, self.bancada, self.lote, self.deposito, caixa, qtd_etq, seriais)

                messagebox.showinfo(self.titulo, f'QRCODE da caixa {caixa} do material {cod_material} gerado!')
                return

            # Se não existir, apresentar mensagem de erro
            else:
                messagebox.showerror(self.titulo, f'Não existem materiais na caixa {caixa} para geração de QRCODE.')
                return

        except:
            return messagebox.showerror(self.titulo, f'Erro ao gerar o QRCODE.')

    # Apagar apenas a linha selecionada
    def apagar_serial(self):
        try:
            # Armazenar serial em variável
            linha = self.lista_bipados.selection()[0]
            serial = self.lista_bipados.item(linha, 'values')[1]

            if self.tipo_leitura == 'TRANSFERENCIA':
                # Comando para verificar se o serial está no packlist
                self.bd_local.cursor.execute(f"SELECT packlist FROM {self.nf} WHERE serial_primario = '{serial}'")
                data = self.bd_local.cursor.fetchone()

                # Se estiver no packlist, apenas apague informações da leitura, mantendo o serial
                if data[0][:3] == 'SIM':
                    self.bd_local.cursor.execute(f"UPDATE {self.nf} SET lote = NULL, caixa = NULL, bancada = NULL, "
                                                 f"usuario = NULL, hora = NULL WHERE serial_primario = '{serial}'")
                # Se não estiver, apague toda as informações do serial selecionado
                else:
                    self.bd_local.cursor.execute(f"DELETE FROM {self.nf} WHERE serial_primario = '{serial}'")

            elif self.tipo_leitura == 'LABORATORIO':
                # Apagar todas as informações do serial selecionado
                self.bd_local.cursor.execute(f"DELETE FROM {self.nf} "
                                             f"WHERE serial_primario = '{serial}' AND lote = '{self.lote}'")

            elif self.tipo_leitura == 'REVERSA':
                # Comando para verificar se o serial está no packlist
                self.bd_local.cursor.execute(f"SELECT packlist FROM {self.fornecedor} "
                                             f"WHERE serial_primario = '{serial}'")
                data = self.bd_local.cursor.fetchone()

                # Se estiver no packlist, apenas apague informações da leitura, mantendo o serial
                if data[0][:3] == 'SIM':
                    self.bd_local.cursor.execute(f"UPDATE {self.fornecedor} SET lote = NULL, caixa = NULL, "
                                                 f"bancada = NULL, usuario = NULL, hora = NULL "
                                                 f"WHERE nota_fiscal = '{self.nf}' AND serial_primario = '{serial}'")
                # Se não estiver, apague toda as informações do serial selecionado
                else:
                    self.bd_local.cursor.execute(f"DELETE FROM {self.fornecedor} "
                                                 f"WHERE nota_fiscal = '{self.nf}' AND serial_primario = '{serial}'")

            else:
                # Apagar todas as informações do serial selecionado
                self.bd_local.cursor.execute(f"DELETE FROM assinantes "
                                             f"WHERE serial_primario = '{serial}' AND lote = '{self.lote}'")

            # Salvar as modificações no banco de dados
            self.bd_local.connect.commit()

            # Atualizar todas as informações da leitura
            self.update_info()
            return messagebox.showinfo(self.titulo, 'Serial apagado com sucesso.')

        except:
            return messagebox.showerror(self.titulo, 'Erro ao apagar o serial.')

    # Apagar todos os seriais da caixa escolhida
    def apagar_caixa(self):
        # Variável para armazenar o resultado de busca de códigos de materiais
        data = None

        # Verificação de tipo de leitura
        if self.tipo_leitura == 'TRANSFERENCIA' or self.tipo_leitura == 'LABORATORIO':
            self.bd_local.cursor.execute(f"SELECT DISTINCT material FROM {self.nf} WHERE lote = '{self.lote}'")
            data = self.bd_local.cursor.fetchall()

        elif self.tipo_leitura == 'REVERSA':
            self.bd_local.cursor.execute(f"SELECT DISTINCT material FROM {self.fornecedor} "
                                         f"WHERE nota_fiscal = {self.nota} AND lote = '{self.lote}'")
            data = self.bd_local.cursor.fetchall()

        # Se a o resultado da buscar encontrar materiais, adicionar os códigos em uma lista
        if data is not None:
            materiais = []

            for i in range(0, len(data)):
                materiais.append(data[i][0])

        # Se não encontrar, definir a variável da lista como nada
        else:
            materiais = None

        info = get_material(self.tipo_leitura, materiais)
        if not info:
            return

        cod_material = info[0]
        caixa = info[1]

        try:
            if self.tipo_leitura == 'ASSINANTES':
                # Apagar todas as informações do serial
                self.bd_local.cursor.execute(f"DELETE FROM assinantes WHERE caixa = '{caixa}' AND lote = '{self.lote}'")

                msg = f'Caixa {caixa} apagada com sucesso!'

            else:

                if self.tipo_leitura == 'LABORATORIO':
                    # Apagar todas as informações do serial
                    self.bd_local.cursor.execute(f"DELETE FROM {self.nf} "
                                                 f"WHERE (material = '{cod_material}' AND caixa = '{caixa}') "
                                                 f"AND lote = '{self.lote}'")

                else:
                    if self.tipo_leitura == 'TRANSFERENCIA':
                        # Apagar todas as informações do serial que não esteja no packlist
                        query_01 = f"DELETE FROM {self.nf} WHERE (material = '{cod_material}' AND caixa = '{caixa}') " \
                                   f"AND (packlist LIKE 'NÃO%' AND lote = '{self.lote}')"

                        # Apagar apenas as informações da leitura, mantendo o serial
                        query_02 = f"UPDATE {self.nf} SET lote = NULL, caixa = NULL, bancada = NULL, usuario = NULL, " \
                                   f"hora = NULL WHERE (material = '{cod_material}' AND caixa = '{caixa}') " \
                                   f"AND (packlist LIKE 'SIM%' and lote = '{self.lote}')"

                    else:
                        # Apagar todas as informações do serial que não esteja no packlist
                        query_01 = f"DELETE FROM {self.fornecedor} WHERE nota_fiscal = '{self.nf}' " \
                                   f"AND (material = '{cod_material}' AND caixa = '{caixa}') " \
                                   f"AND (packlist LIKE 'NÃO%' AND lote = '{self.lote}')"

                        # Apagar apenas as informações da leitura, mantendo o serial
                        query_02 = f"UPDATE {self.fornecedor} SET lote = NULL, caixa = NULL, bancada = NULL, " \
                                   f"usuario = NULL, hora = NULL " \
                                   f"WHERE (nota_fiscal = '{self.nf}' AND material = '{cod_material}') " \
                                   f"AND (caixa = '{caixa}' AND lote = '{self.lote}')"

                    # Executar comandos no banco de dados
                    self.bd_local.cursor.execute(query_01)
                    self.bd_local.cursor.execute(query_02)

                msg = f'Caixa {caixa} do material {cod_material} apagada com sucesso!'

            # Salvar as modificações no banco de dados
            self.bd_local.connect.commit()

            # Atualizar todas as informações da leitura
            self.update_info()
            return messagebox.showinfo(self.titulo, msg)

        except:
            return messagebox.showerror(self.titulo, f'Erro ao apagar a caixa.')

    # Salvar pallet no banco de dados da rede
    def salvar_pallet(self):
        try:
            # Primeiro salvar o pallet local
            self.salvar_pallet_local()

            # Verificar e definir tipo de leitura
            if self.tipo_leitura == 'TRANSFERENCIA':
                # Armazenar todas as informações de seriais e leitura em uma lista
                self.bd_local.cursor.execute(f"SELECT * FROM {self.nf} WHERE lote = '{self.lote}'")
                dados_leitura = self.bd_local.cursor.fetchall()

                # Definir banco de dados da rede
                bd_rede = BancoUsadRede()

                # Armazenar informações da leitura do banco de dados local em variáveis
                for i in range(0, len(dados_leitura)):
                    material = dados_leitura[i][1]
                    serial_primario = dados_leitura[i][2]
                    serial_secundario = dados_leitura[i][3]
                    serial_number = dados_leitura[i][4]
                    tecnologia = dados_leitura[i][5]
                    estado = dados_leitura[i][6]
                    local = dados_leitura[i][7]
                    packlist = dados_leitura[i][8]
                    lote = dados_leitura[i][9]
                    caixa = dados_leitura[i][10]
                    bancada = dados_leitura[i][11]
                    usuario = dados_leitura[i][12]
                    hora = dados_leitura[i][13]

                    # Verificar se o serial está no packlist
                    bd_rede.cursor.execute(f"SELECT * FROM {self.nf} WHERE serial_primario = '{serial_primario}'")
                    data = bd_rede.cursor.fetchone()

                    if data:
                        # Se o serial estiver no packlist, apenas atualize as informações de leitura
                        bd_rede.cursor.execute(f"UPDATE {self.nf} SET packlist = '{packlist}', lote = '{lote}', "
                                               f"caixa = '{caixa}', bancada = '{bancada}', usuario = '{usuario}', "
                                               f"hora = '{hora}' WHERE serial_primario = '{serial_primario}'")

                    else:
                        # Se o serial não estiver no packlist, adicione todas as informações do serial e de leitura
                        bd_rede.cursor.execute(f"INSERT INTO {self.nf} (material, serial_primario, serial_secundario, "
                                               f"serial_number, tecnologia, estado, local, packlist, lote, caixa, "
                                               f"bancada, usuario, hora) "
                                               f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                               (material, serial_primario, serial_secundario, serial_number, tecnologia,
                                                estado, local, packlist, lote, caixa, bancada, usuario, hora))

            elif self.tipo_leitura == 'REVERSA':
                # Armazenar todas as informações de seriais e leitura em uma lista
                self.bd_local.cursor.execute(f"SELECT * FROM {self.fornecedor} "
                                             f"WHERE nota_fiscal = '{self.nf}' AND lote = '{self.lote}'")
                dados_leitura = self.bd_local.cursor.fetchall()

                # Definir banco de dados da rede
                bd_rede = BancoUrevRede()

                # Armazenar informações da leitura do banco de dados local em variáveis
                for i in range(0, len(dados_leitura)):
                    nota_fiscal = dados_leitura[i][1]
                    material = dados_leitura[i][2]
                    serial_primario = dados_leitura[i][3]
                    serial_secundario = dados_leitura[i][4]
                    serial_number = dados_leitura[i][5]
                    tecnologia = dados_leitura[i][6]
                    estado = dados_leitura[i][7]
                    local = dados_leitura[i][8]
                    packlist = dados_leitura[i][9]
                    lote = dados_leitura[i][10]
                    caixa = dados_leitura[i][11]
                    bancada = dados_leitura[i][12]
                    usuario = dados_leitura[i][13]
                    hora = dados_leitura[i][14]

                    # Verificar se o serial está no packlist
                    bd_rede.cursor.execute(f"SELECT * FROM {self.fornecedor} "
                                           f"WHERE nota_fiscal = '{self.nf}' AND serial_primario = '{serial_primario}'")
                    data = bd_rede.cursor.fetchone()

                    if data:
                        # Se o serial estiver no packlist, apenas atualize as informações de leitura
                        bd_rede.cursor.execute(f"UPDATE {self.fornecedor} SET packlist = '{packlist}', "
                                               f"lote = '{lote}', caixa = '{caixa}', bancada = '{bancada}', "
                                               f"usuario = '{usuario}', hora = '{hora}' "
                                               f"WHERE nota_fiscal = '{self.nf}' "
                                               f"AND serial_primario = '{serial_primario}'")

                    else:
                        # Se o serial não estiver no packlist, adicione todas as informações do serial e de leitura
                        bd_rede.cursor.execute(f"INSERT INTO {self.fornecedor} (nota_fiscal, material, "
                                               f"serial_primario, serial_secundario, serial_number, tecnologia, "
                                               f"estado, local, packlist, lote, caixa, bancada, usuario, hora) "
                                               f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                               (nota_fiscal, material, serial_primario, serial_secundario,
                                                serial_number, tecnologia, estado, local, packlist, lote, caixa,
                                                bancada, usuario, hora))

            else:
                if self.tipo_leitura == 'LABORATORIO':
                    # Armazenar todas as informações de seriais e leitura em uma lista
                    self.bd_local.cursor.execute(f"SELECT * FROM {self.nf} WHERE lote = '{self.lote}'")
                    dados_leitura = self.bd_local.cursor.fetchall()

                    # Definir banco de dados da rede
                    bd_rede = BancoUsusRede()

                    # Definir tabela do banco de dados
                    tabela = self.nf

                    # Criar tabela, caso não exista
                    bd_rede.create_table(tabela)

                else:
                    self.bd_local.cursor.execute(f"SELECT * FROM assinantes WHERE lote = '{self.lote}'")
                    dados_leitura = self.bd_local.cursor.fetchall()

                    # Definir banco de dados da rede
                    bd_rede = BancoUassRede()

                    # Definir tabela do banco de dados
                    tabela = 'assinantes'

                # Armazenar informações da leitura do banco de dados local em variáveis
                for i in range(0, len(dados_leitura)):
                    material = dados_leitura[i][1]
                    serial_primario = dados_leitura[i][2]
                    serial_secundario = dados_leitura[i][3]
                    serial_number = dados_leitura[i][4]
                    tecnologia = dados_leitura[i][5]
                    estado = dados_leitura[i][6]
                    local = dados_leitura[i][7]
                    lote = dados_leitura[i][8]
                    caixa = dados_leitura[i][9]
                    bancada = dados_leitura[i][10]
                    usuario = dados_leitura[i][11]
                    hora = dados_leitura[i][12]

                    # Adicionar todas as informações do serial e leitura no banco de dados
                    bd_rede.cursor.execute(f"INSERT INTO {tabela} (material, serial_primario, serial_secundario, "
                                           f"serial_number, tecnologia, estado, local, lote, caixa, bancada, usuario, "
                                           f"hora) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                           (material, serial_primario, serial_secundario, serial_number, tecnologia,
                                            estado, local, lote, caixa, bancada, usuario, hora))

            # Salvar as modificações no banco de dados
            bd_rede.connect.commit()

            # Fechar conexão com o banco de dados da rede
            bd_rede.connect.close()

            # Salvar produtividade
            self.salvar_produtividade()

            # Resetar número de lote
            self.lote = ''
            self.apagar_lote()

            # Atualizar todas as informações da leitura
            self.update_info()
            messagebox.showinfo(self.titulo, 'Pallet salvo com sucesso!')
            self.voltar()

        except:
            return messagebox.showerror(self.titulo, 'Erro ao salvar pallet. Tente novamente.')

    # Salvar o pallet na pasta local
    def salvar_pallet_local(self):
        # Definir tipo de leitura
        if self.tipo_leitura == 'TRANSFERENCIA':
            db_path = 'assets/database/usad.db'
            query = f"SELECT * FROM {self.nf}"
            filename = f'NF {self.nf[3:]} - {self.lote}'

        elif self.tipo_leitura == 'REVERSA':
            db_path = 'assets/database/urev.db'
            query = f"SELECT * FROM {self.fornecedor} WHERE nota_fiscal = {self.nf}"
            fornecedor = str(self.fornecedor).replace('_', ' ')
            fornecedor = fornecedor.capitalize()
            filename = f'Forn. {fornecedor} NF {self.nf} - {self.lote}'

        elif self.tipo_leitura == 'LABORATÓRIO':
            db_path = 'assets/database/usus.db'
            tabela = f'nf_{self.cod_material}'
            query = f"SELECT * FROM {tabela} WHERE lote = '{self.lote}'"
            filename = f'{self.cod_material} - {self.lote}'

        else:
            db_path = 'assets/database/uass.db'
            query = f"SELECT * FROM assinantes WHERE lote = '{self.lote}'"
            filename = f'Assinantes - {self.lote}'

        pasta = pasta_leitura(self.tipo_leitura)
        filename = rf'{pasta}\{filename}.xlsx'

        # Realizando a leitura do banco de dados
        bd_export = sqlite3.connect(db_path)
        df = pd.read_sql_query(query, bd_export).drop(columns=['id'])

        # Exportando do banco de dados para o excel
        df.to_excel(filename, index=False)
        bd_export.close()

    # Apagar o último número de lote do usuário
    def apagar_lote(self):
        # Conectando ao banco de dados de usuários
        bd_user = BancoUsuarios()

        # Apagando o número de lote atual
        bd_user.cursor.execute(f"UPDATE usuarios SET lote = NULL WHERE nome = '{self.user}'")

        # Salvar as modificações no banco de dados
        bd_user.connect.commit()

        # Fechar conexão com o banco de dados
        bd_user.connect.close()

    # Salvar a produtividade no banco de dados
    def salvar_produtividade(self):
        # Definindo o ano atual
        agora = datetime.now()
        ano = agora.strftime('%Y')

        # Conectando no banco de dados de produtividade
        bd_param = BancoParametrosRede()

        try:
            # Apagando os números de produtividade antigo, para depois inserir os números atualizados
            bd_param.cursor.execute(f"DELETE FROM produtividade WHERE lote = '{self.lote}'")

            # Verificar tipo de leitura para definir tabela e número da nota fiscal
            if self.tipo_leitura == 'TRANSFERENCIA':
                tabela = self.nf
                nota_fiscal = self.nota

            elif self.tipo_leitura == 'REVERSA':
                tabela = self.fornecedor
                nota_fiscal = self.nf

            elif self.tipo_leitura == 'LABORATORIO':
                tabela = self.nf
                nota_fiscal = 'LOTE'

            else:
                tabela = 'assinantes'
                nota_fiscal = 'ASSIN.'

            # Obter a lista distinta de materiais bipados no lote
            self.bd_local.cursor.execute(f"SELECT DISTINCT material FROM {tabela} WHERE lote = '{self.lote}'")
            lista_materiais = self.bd_local.cursor.fetchall()

            # Loop para percorrer todos os códigos de materiais
            for i in range(0, len(lista_materiais)):
                # Definir código de material
                cod_sap = lista_materiais[i][0]

                # Obter a lista distinta de datas em que foram bipados os materiais
                self.bd_local.cursor.execute(f"SELECT DISTINCT substr(hora, 1, 5) FROM {tabela} "
                                             f"WHERE material = '{cod_sap}' AND lote = '{self.lote}'")
                lista_data = self.bd_local.cursor.fetchall()

                # Loop para percorrer toda a lista de datas encontradas com o material definido
                for j in range(0, len(lista_data)):
                    # Definir data
                    data = lista_data[j][0]

                    # Obter a lista distinta de horas em que foram bipados os materiais
                    self.bd_local.cursor.execute(f"SELECT DISTINCT substr(hora, 7, 2) FROM {tabela} "
                                                 f"WHERE (material = '{cod_sap}' AND lote = '{self.lote}') "
                                                 f"AND substr(hora, 1, 5) = '{data}'")
                    lista_hora = self.bd_local.cursor.fetchall()

                    # Loop para percorrer toda a lista de horas encontradas com materiais e datas definidas
                    for k in range(0, len(lista_hora)):
                        # Definir data
                        data = lista_data[j][0]

                        # Definir hora
                        hora = lista_hora[k][0]

                        # Obter a quantidade de materiais bipados nos parâmetros definidos
                        self.bd_local.cursor.execute(f"SELECT COUNT(*) FROM {tabela} "
                                                     f"WHERE (lote = '{self.lote}' AND material = '{cod_sap}') "
                                                     f"AND (substr(hora, 1, 5) = '{data}' "
                                                     f"AND substr(hora, 7, 2) = '{hora}')")
                        qtd = self.bd_local.cursor.fetchone()
                        qtd = qtd[0]

                        # Ajustar o formato de data para armazenar na tabela
                        dia = data[:2]
                        mes = data[3:]
                        data = f'{ano}-{mes}-{dia}'

                        # Inserir informações de produtividade na tabela
                        bd_param.cursor.execute(f"INSERT INTO produtividade (data, usuario, bancada, nota_fiscal, "
                                                f"deposito, cod_sap, quantidade, hora, lote) "
                                                f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                                (data, self.user, self.bancada, nota_fiscal, self.deposito, cod_sap,
                                                 qtd, hora, self.lote))

            # Salvar as modificações no banco de dados
            bd_param.connect.commit()

        except:
            pass

        # Fechar conexão com o banco de dados
        bd_param.connect.close()

    # Voltar ao menu principal
    def voltar(self):
        encerrar_chromedriver()
        voltar_menu(self.user, self.perfil, self.lote)


# Selecionar material e caixa para gerar QRCODE
class SelecionarMaterial:
    def __init__(self, master, tipo_leitura, materiais):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão
        self.tipo_leitura = tipo_leitura
        self.info = []  # Lista usada para retornar informações

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-get_material.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_continuar = PhotoImage(file=r'assets/images/btn-continuar.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title('TPC PLH | Framework IN - Selecionar material')
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        if self.tipo_leitura == 'ASSINANTES':
            # Para leitura de assinantes, apagar todos os seriais de acordo com o número da caixa
            self.lbl_material = Label(self.window, text='ASSINANTES', font=self.fonte, background="#FFFFFF",
                                      justify=CENTER)
            self.lbl_material.place(width=110, height=30, x=145, y=276)

        else:
            # Combobox com os materiais disponíveis para gerar QRCODE
            self.cb_materiais = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(),
                                             justify=CENTER)
            self.cb_materiais.option_add('*TCombobox*Listbox.Justify', 'center')
            self.cb_materiais['values'] = materiais
            self.cb_materiais['state'] = 'readonly'
            self.cb_materiais.place(width=160, height=35, x=120, y=274)

        # Criação da caixa de texto para inserir número da caixa
        self.txt_caixa = Entry(self.window, font=self.fonte, justify=CENTER, background="#F2F2F2", border=False)
        self.txt_caixa.place(width=50, height=35, x=175, y=346)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para confirmar
        self.btn_continuar = Button(self.window, image=self.img_continuar, border=False, command=self.continuar,
                                    activebackground='#FFFFFF')
        self.btn_continuar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def continuar(self):
        # Definir código do material
        if self.tipo_leitura == 'ASSINANTES':
            cod_material = 'ASSINANTES'
        else:
            cod_material = self.cb_materiais.get()

        # Verificar se o material foi selecionado.
        if len(cod_material) == 0:
            return messagebox.showerror(self.titulo, 'Selecione o código do material.')

        # Definir o número da caixa
        caixa = self.txt_caixa.get()

        # Verificar se o número da caixa foi digitado
        if len(caixa) == 0:
            return messagebox.showerror(self.titulo, 'Informe o número da caixa.')

        # Se for digitado um número com mais de 2 digitos, retornar caixa inválida
        elif len(caixa) > 2:
            return messagebox.showerror(self.titulo, 'Número de caixa inválida.')

        # Verificar se a caixa contém apenas números
        else:
            try:
                caixa = int(caixa)
            except:
                return messagebox.showerror(self.titulo, 'Número de caixa inválida.')

        # Retornar informações para gerar QRCODE
        self.info = [cod_material, caixa]
        self.cancelar()
        return self.info

    def cancelar(self):
        self.window.destroy()


# Selecionar material e caixa para gerar QRCODE
class ExportarLote:
    def __init__(self, master, cod_material):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão
        self.cod_material = cod_material

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-export_lote.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_exportar = PhotoImage(file=r'assets/images/btn-exportar.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title('TPC PLH | Framework IN - Exportar Lote')
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # EntryBox com calendário para selecionar data
        self.txt_data = DateEntry(self.window, locale='pt_BR', font=self.fonte, justify=CENTER)
        self.txt_data.bind('<<DateEntrySelected>>', self.update)
        self.txt_data.place(width=140, height=35, x=130, y=276)

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        # Combobox para listar os números de lotes
        self.cb_lotes = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(), justify=CENTER)
        self.cb_lotes.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_lotes['state'] = 'readonly'
        self.cb_lotes.place(width=180, height=35, x=110, y=346)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para salvar
        self.btn_exportar = Button(self.window, image=self.img_exportar, border=False, command=self.exportar,
                                   activebackground='#FFFFFF')
        self.btn_exportar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Atualizar ao iniciar a janela
        self.update()

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def update(self, event=None):
        # Apagar o valor atual da combobox
        self.cb_lotes.set('')

        # Definir data para criar lista dos lotes
        data = self.txt_data.get_date()

        # Conectar no banco de dados da rede
        bd_param = BancoParametrosRede()

        # Executando a busca dos números de lotes no banco de dados
        bd_param.cursor.execute(f"SELECT lote from lotes WHERE data = '{data}' AND nota_fiscal = '{self.cod_material}'")
        dados = bd_param.cursor.fetchall()
        bd_param.connect.close()

        lista_lotes = []

        # Se encontrar lotes, adicionar os valores na combobox
        if dados:

            for i in range(0, len(dados)):
                lista_lotes.append(dados[i][0])

        self.cb_lotes['values'] = lista_lotes

    def exportar(self):
        lote = self.cb_lotes.get()

        if len(lote) == 0:
            messagebox.showerror(self.titulo, 'Selecione o número do lote.')
            return

        # Definir caminho do banco de dados da rede
        db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\usus.db'

        # Definir nome da tabela
        tabela = f'nf_{self.cod_material}'

        try:
            # Conectando no banco de dados
            bd_rede = BancoUsusRede()

            # Executando a busca e armazenando o resultado em uma variável
            bd_rede.cursor.execute(f"SELECT * FROM {tabela} WHERE lote = '{lote}'")
            data = bd_rede.cursor.fetchone()

            # Fechando a conexão
            bd_rede.connect.close()

            if not data:
                messagebox.showerror(self.titulo, f'A leitura do lote {lote} não foi finalizada.')
                return

            filename = filedialog.asksaveasfilename(filetypes=[('Excel file (xlsx)', '.xlsx')], defaultextension='xlsx',
                                                    initialfile=f'{self.cod_material} - {lote}')

            if len(filename) < 1:
                messagebox.showwarning(self.titulo, 'Processo de exportação cancelado.')
                return

            bd_export = sqlite3.connect(db_path)
            df = pd.read_sql_query(f"SELECT * FROM {tabela} WHERE lote = '{lote}'", bd_export).drop(columns=['id'])
            df.to_excel(filename, index=False)
            bd_export.close()
            messagebox.showinfo(self.titulo, f'Lote {lote} do material {self.cod_material} exportado com sucesso!')
            self.cancelar()
            return 'sucesso'

        except:
            messagebox.showerror(self.titulo, f'Falha ao exportar o lote {lote} do material {self.cod_material}.')
            return

    def cancelar(self):
        self.window.destroy()


# Consultar e exportar o relatório de produtividade
class Produtividade:
    def __init__(self, master, perfil, usuario):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão
        self.fonte_lista = ['Calibri', 16, 'normal']  # Fonte usada na TreeView
        self.perfil = perfil  # Permissão de perfil do usuário para definir consulta
        self.usuario = usuario

        # Definir data atual em números inteiros
        self.hoje = datetime.now()
        self.hoje = self.hoje.strftime('%Y%m%d')
        self.hoje = int(self.hoje)

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-produtividade.png')
        self.img_limpar = PhotoImage(file=r'assets/images/btn-limpar_filtro.png')
        self.img_exportar = PhotoImage(file=r'assets/images/btn-exportar.png')
        self.img_voltar = PhotoImage(file=r'assets/images/btn-voltar.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(1000, 650)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title('TPC PLH | Framework IN - Exportar Lote')
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # EntryBox com calendário para selecionar data inicial
        self.data_inicial = DateEntry(self.window, locale='pt_BR', font=self.fonte, justify=CENTER)
        self.data_inicial.bind('<<DateEntrySelected>>', self.update)
        self.data_inicial.place(width=140, height=35, x=70, y=174)

        # EntryBox com calendário para selecionar data final
        self.data_final = DateEntry(self.window, locale='pt_BR', font=self.fonte, justify=CENTER)
        self.data_final.bind('<<DateEntrySelected>>', self.update)
        self.data_final.place(width=140, height=35, x=290, y=174)

        # Definindo o estilo a ser usado na combobox
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('TCombobox', rowheight=25, border=False, font=self.fonte)
        style.map('TCombobox')

        # Combobox para listar os depósitos
        self.cb_depositos = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(),
                                         justify=CENTER)
        self.cb_depositos.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_depositos['state'] = 'readonly'
        self.cb_depositos.bind('<<ComboboxSelected>>', self.update)
        self.cb_depositos.place(width=180, height=35, x=490, y=174)
        self.cb_depositos.set('TODOS')

        # Combobox para listar as bancadas
        self.cb_bancadas = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(),
                                        justify=CENTER)
        self.cb_bancadas.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_bancadas['state'] = 'readonly'
        self.cb_bancadas.bind('<<ComboboxSelected>>', self.update)
        self.cb_bancadas.place(width=140, height=35, x=70, y=272)
        self.cb_bancadas.set('TODAS')

        # Combobox para listar os usuários
        self.cb_usuarios = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(),
                                        justify=CENTER)
        self.cb_usuarios.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_usuarios['state'] = 'readonly'
        self.cb_usuarios.bind('<<ComboboxSelected>>', self.update)
        self.cb_usuarios.place(width=140, height=35, x=290, y=272)
        if self.perfil == 'Default':
            self.cb_usuarios.set(self.usuario)
        else:
            self.cb_usuarios.set('TODOS')

        # Combobox para listar os horários
        self.cb_horas = ttk.Combobox(self.window, font=self.fonte, width=200, textvariable=StringVar(), justify=CENTER)
        self.cb_horas.option_add('*TCombobox*Listbox.Justify', 'center')
        self.cb_horas['state'] = 'readonly'
        self.cb_horas.bind('<<ComboboxSelected>>', self.update)
        self.cb_horas.place(width=140, height=35, x=510, y=272)
        self.cb_horas.set('TODAS')

        # Configurando os estilos e cores da TreeView
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('Treeview', background='#F2F2F2', fieldbackground='#F2F2F2', rowheight=25,
                        font=self.fonte_lista)
        style.map('Treeview')

        # Criar e inserir a TreeView de seriais bipados
        self.lista_produtividade = ttk.Treeview(self.window, columns=('col1', 'col2', 'col3', 'col4', 'col5'),
                                                selectmode='browse')
        self.lista_produtividade.place(width=620, height=277, x=49, y=360)

        # Criar e inserir a barra de rolagem na lista de seriais bipados
        self.scroll_produtividade = ttk.Scrollbar(self.window, orient='vertical',
                                                  command=self.lista_produtividade.yview)
        self.scroll_produtividade.place(width=16, height=255 + 20, x=450 + 200 + 2, y=361)
        self.lista_produtividade.configure(yscrollcommand=self.scroll_produtividade.set)

        # Cabeçalho da TreeView
        self.lista_produtividade.heading('#0', text='ID')
        self.lista_produtividade.heading('#1', text='Data')
        self.lista_produtividade.heading('#2', text='Hora')
        self.lista_produtividade.heading('#3', text='Bancada')
        self.lista_produtividade.heading('#4', text='Depósito')
        self.lista_produtividade.heading('#5', text='Quantidade')
        self.lista_produtividade.config(show='tree')  # Configuração pra ocultar o cabeçalho da TreeView

        # Largura das colunas da TreeView
        self.lista_produtividade.column('#0', width=1, stretch=False)  # Configuração pra ocultar a primeira coluna (ID)
        self.lista_produtividade.column('#1', width=150, anchor='center')
        self.lista_produtividade.column('#2', width=100, anchor='center')
        self.lista_produtividade.column('#3', width=100, anchor='center')
        self.lista_produtividade.column('#4', width=140, anchor='center')
        self.lista_produtividade.column('#5', width=110, anchor='center')

        # Botão para exportar o relatório
        if self.perfil != 'Default':
            self.btn_exportar = Button(self.window, image=self.img_exportar, border=False, command=self.exportar,
                                       activebackground='#FFFFFF')
            self.btn_exportar.place(width=150, height=35, x=760, y=460)

        # Botão para limpar todos os filtros
        self.btn_limpar = Button(self.window, image=self.img_limpar, border=False, command=self.limpar_filtro,
                                 activebackground='#FFFFFF')
        self.btn_limpar.place(width=150, height=35, x=760, y=505)

        # Botão para voltar ao menu principal
        self.btn_voltar = Button(self.window, image=self.img_voltar, border=False, command=self.voltar,
                                 activebackground='#FFFFFF')
        self.btn_voltar.place(width=150, height=35, x=760, y=550)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=830, y=625)

        # Conectando no banco de dados dos parâmetros da rede
        self.bd_param = BancoParametrosRede()

        # Fechar a conexão com o banco de dados
        self.bd_param.connect.close()

        # Atualizar ao iniciar a janela
        self.update()

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def update_deposito(self):
        # Obtendo os valores selecionados nas combobox
        data_inicio = self.data_inicial.get_date()
        data_fim = self.data_final.get_date()
        bancada = self.cb_bancadas.get()
        usuario = self.cb_usuarios.get()
        hora = self.cb_horas.get()

        # Verificando e definindo a bancada
        if not bancada or bancada == 'TODAS':
            bancada = '%'

        # Verificando e definindo o usuário
        if self.perfil == 'Default':
            usuario = self.usuario
        else:
            if not usuario or usuario == 'TODOS':
                usuario = '%'

        # Verificando e definindo a hora
        if not hora or hora == 'TODAS':
            hora = '%'

        # Buscandos os as informações no banco de dados
        self.bd_param.cursor.execute(f"SELECT DISTINCT deposito FROM produtividade "
                                     f"WHERE (data BETWEEN '{data_inicio}' AND '{data_fim}') "
                                     f"AND (bancada LIKE '{bancada}' AND usuario LIKE '{usuario}') "
                                     f"AND hora LIKE '{hora}'")
        dados = self.bd_param.cursor.fetchall()
        depositos = []

        if dados:
            for i in range(0, len(dados)):
                if dados[i][0] == 'UASS':
                    deposito = 'ASSINANTES'
                elif dados[i][0] == 'USUS':
                    deposito = 'LABORATÓRIO'
                elif dados[i][0] == 'UREV':
                    deposito = 'REVERSA'
                else:
                    deposito = 'TRANSFERÊNCIA'

                depositos.append(deposito)

            # Organizar a lista em ordem alfabética
            depositos = sorted(depositos)

            # Adicionar a opção para todos os depósitos no início da lista
            depositos.insert(0, 'TODOS')

        # Adicionar a lista nas opções da combobox
        self.cb_depositos['values'] = depositos

    def update_bancada(self):
        # Obtendo os valores selecionados nas combobox
        data_inicio = self.data_inicial.get_date()
        data_fim = self.data_final.get_date()
        deposito = self.cb_depositos.get()
        usuario = self.cb_usuarios.get()
        hora = self.cb_horas.get()

        # Verificando e definindo o depósito
        if deposito == 'ASSINANTES':
            deposito = 'UASS'
        elif deposito == 'LABORATÓRIO':
            deposito = 'USUS'
        elif deposito == 'REVERSA':
            deposito = 'UREV'
        elif deposito == 'TRANSFERÊNCIA':
            deposito = 'USAD'
        else:
            deposito = '%'

        # Verificando e definindo o usuário
        if self.perfil == 'Default':
            usuario = self.usuario
        else:
            if not usuario or usuario == 'TODOS':
                usuario = '%'

        # Verificando e definindo a hora
        if not hora or hora == 'TODAS':
            hora = '%'

        # Buscandos os as informações no banco de dados
        self.bd_param.cursor.execute(f"SELECT DISTINCT bancada FROM produtividade "
                                     f"WHERE (data BETWEEN '{data_inicio}' AND '{data_fim}') "
                                     f"AND (deposito LIKE '{deposito}' AND usuario LIKE '{usuario}') "
                                     f"AND hora LIKE '{hora}'")
        dados = self.bd_param.cursor.fetchall()
        bancadas = []

        if dados:
            for i in range(0, len(dados)):
                bancadas.append(dados[i])

            # Organizar a lista em ordem alfabética
            bancadas = sorted(bancadas)

            # Adicionar a opção para todos os depósitos no início da lista
            bancadas.insert(0, 'TODAS')

        # Adicionar a lista nas opções da combobox
        self.cb_bancadas['values'] = bancadas

    def update_usuario(self):
        # Obtendo os valores selecionados nas combobox
        data_inicio = self.data_inicial.get_date()
        data_fim = self.data_final.get_date()
        deposito = self.cb_depositos.get()
        bancada = self.cb_bancadas.get()
        hora = self.cb_horas.get()

        # Verificando e definindo o depósito
        if deposito == 'ASSINANTES':
            deposito = 'UASS'
        elif deposito == 'LABORATÓRIO':
            deposito = 'USUS'
        elif deposito == 'REVERSA':
            deposito = 'UREV'
        elif deposito == 'TRANSFERÊNCIA':
            deposito = 'USAD'
        else:
            deposito = '%'

        # Verificando e definindo o usuário
        if not bancada or bancada == 'TODAS':
            bancada = '%'

        # Verificando e definindo a hora
        if not hora or hora == 'TODAS':
            hora = '%'

        # Se for um usuário padrão, consultar apenas a própria produtividade
        if self.perfil == 'Default':
            usuarios = [self.usuario]
        else:
            # Buscandos os as informações no banco de dados
            self.bd_param.cursor.execute(f"SELECT DISTINCT usuario FROM produtividade "
                                         f"WHERE (data BETWEEN '{data_inicio}' AND '{data_fim}') "
                                         f"AND (deposito LIKE '{deposito}' AND bancada LIKE '{bancada}') "
                                         f"AND hora LIKE '{hora}'")
            dados = self.bd_param.cursor.fetchall()
            usuarios = []

            if dados:
                for i in range(0, len(dados)):
                    usuarios.append(dados[i])

                # Organizar a lista em ordem alfabética
                usuarios = sorted(usuarios)

                # Adicionar a opção para todos os depósitos no início da lista
                usuarios.insert(0, 'TODOS')

        # Adicionar a lista nas opções da combobox
        self.cb_usuarios['values'] = usuarios

    def update_hora(self):
        # Obtendo os valores selecionados nas combobox
        data_inicio = self.data_inicial.get_date()
        data_fim = self.data_final.get_date()
        deposito = self.cb_depositos.get()
        bancada = self.cb_bancadas.get()
        usuario = self.cb_usuarios.get()

        # Verificando e definindo o depósito
        if deposito == 'ASSINANTES':
            deposito = 'UASS'
        elif deposito == 'LABORATÓRIO':
            deposito = 'USUS'
        elif deposito == 'REVERSA':
            deposito = 'UREV'
        elif deposito == 'TRANSFERÊNCIA':
            deposito = 'USAD'
        else:
            deposito = '%'

        # Verificando e definindo o usuário
        if not bancada or bancada == 'TODAS':
            bancada = '%'

        # Verificando e definindo o usuário
        if self.perfil == 'Default':
            usuario = self.usuario
        else:
            if not usuario or usuario == 'TODOS':
                usuario = '%'

        # Buscandos os as informações no banco de dados
        self.bd_param.cursor.execute(f"SELECT DISTINCT hora FROM produtividade "
                                     f"WHERE (data BETWEEN '{data_inicio}' AND '{data_fim}') "
                                     f"AND (deposito LIKE '{deposito}' AND bancada LIKE '{bancada}') "
                                     f"AND usuario LIKE '{usuario}'")
        dados = self.bd_param.cursor.fetchall()
        horas = []

        if dados:
            for i in range(0, len(dados)):
                horas.append(dados[i])

            # Organizar a lista em ordem alfabética
            horas = sorted(horas)

            # Adicionar a opção para todos os depósitos no início da lista
            horas.insert(0, 'TODAS')

        # Adicionar a lista nas opções da combobox
        self.cb_horas['values'] = horas

    def update_all(self):
        # Obtendo os valores selecionados nas combobox
        data_inicio = self.data_inicial.get_date()
        data_fim = self.data_final.get_date()
        deposito = self.cb_depositos.get()
        bancada = self.cb_bancadas.get()
        usuario = self.cb_usuarios.get()
        hora = self.cb_horas.get()

        # Verificando e convertendo o depósito
        if deposito == 'ASSINANTES':
            deposito = 'UASS'
        elif deposito == 'LABORATÓRIO':
            deposito = 'USUS'
        elif deposito == 'REVERSA':
            deposito = 'UREV'
        elif deposito == 'TRANSFERÊNCIA':
            deposito = 'USAD'
        else:
            deposito = '%'

        # Verificando e definindo a bancada
        if not bancada or bancada == 'TODAS':
            bancada = '%'

        # Verificando e definindo o usuário
        if self.perfil == 'Default':
            usuario = self.usuario
        else:
            if not usuario or usuario == 'TODOS':
                usuario = '%'

        # Verificando e definindo a hora
        if not hora or hora == 'TODAS':
            hora = '%'

        # Buscandos os as informações no banco de dados
        self.bd_param.cursor.execute(f"SELECT data, hora, bancada, deposito, SUM(quantidade) FROM produtividade "
                                     f"WHERE (data BETWEEN '{data_inicio}' AND '{data_fim}') "
                                     f"AND (deposito LIKE '{deposito}' AND bancada LIKE '{bancada}') "
                                     f"AND (usuario LIKE '{usuario}' AND hora LIKE '{hora}') "
                                     f"GROUP BY data, hora, bancada, deposito "
                                     f"ORDER BY data ASC, hora ASC, bancada ASC, deposito ASC")
        dados = self.bd_param.cursor.fetchall()

        # Fechar a conexão com o banco de dados
        self.bd_param.connect.close()

        # Apagar a lista atual
        self.lista_produtividade.delete(*self.lista_produtividade.get_children())

        # Loop para adicionar os seriais encontrados
        for i in range(0, len(dados)):
            # Formatar data para exibição na TreeView
            data = str(dados[i][0])
            data = datetime.strptime(data, '%Y-%m-%d')
            data = data.strftime('%d/%m/%Y')

            # Lista de valores para adicionar na TreeView
            valores = (f'{data}', f'{dados[i][1]}', f'{dados[i][2]}', f'{dados[i][3]}', f'{dados[i][4]}')

            # Criando e configurando a tag para definir a cor da TreeView
            self.lista_produtividade.tag_configure('default', background='#F2F2F2', foreground='black')

            # Inserindo os seriais na lista
            self.lista_produtividade.insert('', END, values=valores, tags='default')

    def update(self, event=None):
        # Definir data atual
        hoje = datetime.now()
        hoje = hoje.strftime('%Y%m%d')
        hoje = int(hoje)

        # Definir data inicial
        inicio = self.data_inicial.get_date()
        inicio = str(inicio).replace('-', '')
        inicio = int(inicio)

        # Definir data final
        fim = self.data_inicial.get_date()
        fim = str(fim).replace('-', '')
        fim = int(fim)

        # Validar informações
        if inicio > hoje:
            messagebox.showerror(self.titulo, 'A data inicial não pode ser maior do que a data atual.')
            self.data_inicial.set_date(datetime.now())
            return
        elif fim > hoje:
            messagebox.showerror(self.titulo, 'A data final não pode ser maior do que a data atual.')
            self.data_final.set_date(datetime.now())
            return
        elif inicio > fim:
            messagebox.showerror(self.titulo, 'A data inicial não pode ser maior do que a data final.')
            self.data_inicial.set_date(datetime.now())
            return

        # Conectando no banco de dados dos parâmetros da rede
        self.bd_param = BancoParametrosRede()

        # Atualizar informações
        self.update_deposito()
        self.update_bancada()
        self.update_usuario()
        self.update_hora()
        self.update_all()

    def limpar_filtro(self):
        # Apagando todos os valores das combobox
        self.data_inicial.set_date(datetime.now())
        self.data_final.set_date(datetime.now())
        self.cb_depositos.set('TODOS')
        self.cb_bancadas.set('TODAS')
        if self.perfil == 'Default':
            self.cb_usuarios.set(self.usuario)
        else:
            self.cb_usuarios.set('TODOS')
        self.cb_horas.set('TODAS')

        # Atualizando todas as informações
        self.update()

    def exportar(self):
        # Obtendo os valores selecionados nas combobox
        data_inicio = self.data_inicial.get_date()
        data_fim = self.data_final.get_date()
        deposito = self.cb_depositos.get()
        bancada = self.cb_bancadas.get()
        usuario = self.cb_usuarios.get()
        hora = self.cb_horas.get()

        # Verificando e convertendo o depósito
        if deposito == 'ASSINANTES':
            deposito = 'UASS'
        elif deposito == 'LABORATÓRIO':
            deposito = 'USUS'
        elif deposito == 'REVERSA':
            deposito = 'UREV'
        elif deposito == 'TRANSFERÊNCIA':
            deposito = 'USAD'
        else:
            deposito = '%'

        # Verificando e definindo a bancada
        if not bancada or bancada == 'TODAS':
            bancada = '%'

        # Verificando e definindo o usuário
        if self.perfil == 'Default':
            usuario = self.usuario
        else:
            if not usuario or usuario == 'TODOS':
                usuario = '%'

        # Verificando e definindo a hora
        if not hora or hora == 'TODAS':
            hora = '%'

        # Comando para efetuar a busca no banco de dados
        query = f"SELECT * FROM produtividade " \
                f"WHERE (data BETWEEN '{data_inicio}' AND '{data_fim}') " \
                f"AND (deposito LIKE '{deposito}' AND bancada LIKE '{bancada}') " \
                f"AND (usuario LIKE '{usuario}' AND hora LIKE '{hora}') " \
                f"ORDER BY data ASC, hora ASC, bancada ASC, deposito ASC"

        # Verificando se existe dados com os filtros selecionados
        bd_param = BancoParametrosRede()
        data = bd_param.cursor.execute(query).fetchone()
        bd_param.connect.close()

        # Se não existir, exibir mensagem de erro
        if not data:
            return messagebox.showerror(self.titulo, f'Não existe leitura para salvar a produtividade.')

        # Solicitando directório para salvar a planilha
        filename = filedialog.asksaveasfilename(filetypes=[('Excel files', '.xlsx')], defaultextension='xlsx',
                                                initialfile=f'Produtividade Bancadas')

        if len(filename) < 1:
            return messagebox.showwarning(self.titulo, 'Processo cancelado.')

        # Definir directório do banco de dados da rede
        db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\parametros.db'

        # Conectar no banco de dados
        bd_export = sqlite3.connect(db_path)

        # Executar a busca e excluir a coluna ID da exportação
        df = pd.read_sql_query(query, bd_export).drop(columns=['id'])

        # Salvar em excel
        df.to_excel(filename, index=False)

        # Fechar a conexão com o banco de dados
        bd_export.close()

        # Mensagem de aviso de sucesso
        messagebox.showinfo(self.titulo, f'Produtividade das bancadas exportada com sucesso!')

    def voltar(self):
        self.window.destroy()


# Obter o número da bancada
class SelecionarBancada:
    def __init__(self, master, bancada):
        self.titulo = 'TPC PLH | Framework IN'  # Título usado na messagebox
        self.fonte = ("Calibri", 15)  # Fonte padrão
        self.bancada = bancada

        # Importar as imagens
        self.img_fundo = PhotoImage(file=r'assets/images/itfc-selec_bancada.png')
        self.img_cancelar = PhotoImage(file=r'assets/images/btn-cancelar.png')
        self.img_continuar = PhotoImage(file=r'assets/images/btn-continuar.png')

        # Obter coordenadas para centralizar a janela na tela do usuário
        window_info = coordenadas(400, 500)

        # Configurações principais da janela
        self.window = Toplevel(master)
        self.window.title('TPC PLH | Framework IN - Selecionar bancada')
        self.window.geometry('%dx%d+%d+%d' % window_info)
        self.window.resizable(False, False)
        self.window.iconbitmap(default=r'assets/images/icon.ico')

        # Criação da label com a imagem de fundo
        self.lbl_fundo = Label(self.window, image=self.img_fundo)
        self.lbl_fundo.image = self.img_fundo
        self.lbl_fundo.pack()

        # Entrybox para informar o número da bancada
        self.txt_bancada = Entry(self.window, font=self.fonte, justify=CENTER, background="#F2F2F2", border=False)
        self.txt_bancada.bind("<Return>", self.continuar)
        self.txt_bancada.place(width=50, height=35, x=176, y=311)

        # Botão para cancelar
        self.btn_cancelar = Button(self.window, image=self.img_cancelar, border=False, command=self.cancelar,
                                   activebackground='#FFFFFF')
        self.btn_cancelar.place(width=150, height=35, x=30, y=420)

        # Botão para salvar
        self.btn_continuar = Button(self.window, image=self.img_continuar, border=False, command=self.continuar,
                                    activebackground='#FFFFFF')
        self.btn_continuar.place(width=150, height=35, x=220, y=420)

        # Label com os créditos de criação
        self.lbl_creditos = Button(self.window, text='by Wanderson Carelli', font=("Calibri", "13", "bold", "italic"),
                                   foreground='#1E3C96', border=False, background='#FFFFFF', activebackground='#FFFFFF',
                                   command=open_credits)
        self.lbl_creditos.place(width=170, height=25, x=230, y=475)

        # Entrar na caixa de texto da bancada
        self.txt_bancada.focus_set()

        # Manter a janela de alteração ativa
        self.window.grab_set()

        # Manter bloqueada a janela principal enquanto a janela de alteração estiver aberta
        root.wait_window(self.window)

        # Desbloquear a janela principal ao fechar a janela de alteração
        self.window.grab_release()

    def continuar(self, event=None):
        self.bancada = self.txt_bancada.get()

        if not self.bancada:
            messagebox.showerror(self.titulo, 'Informe o número da bancada.')
            self.txt_bancada.focus_set()
            return

        elif len(self.bancada) > 2:
            messagebox.showerror(self.titulo, 'Número de bancada inválido.')
            self.txt_bancada.delete(0, END)
            return

        else:
            try:
                self.bancada = int(self.bancada)
            except:
                self.bancada = None
                messagebox.showerror(self.titulo, 'Número de bancada inválido.')
                self.txt_bancada.delete(0, END)
                return

        if type(self.bancada) == int:
            self.bancada = f'{self.bancada:02d}'
            self.window.destroy()

            return self.bancada

    def cancelar(self):
        self.bancada = None
        self.window.destroy()
        return self.bancada


root = Tk()
Application(root)
root.mainloop()

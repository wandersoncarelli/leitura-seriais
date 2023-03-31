from datetime import datetime
import sqlite3


# Usuários
class BancoUsuarios:
    # db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\users.db'
    db_path = 'assets/database/rede/users.db'  # Arquivo local usado para testes e desenvolvimento

    def __init__(self):
        # Conectando o banco de dados ao iniciar a classe
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()
        self.create_table()

    # Criando a tabela de usuários, se não existir
    def create_table(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY,
        nome TEXT NOT NULL,
        senha TEXT NOT NULL,
        perfil TEXT NOT NULL,
        lote TEXT,
        UNIQUE(nome))""")


# Parâmetros de materiais local
class BancoParametrosLocal:
    db_path = 'assets/database/parametros.db'

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_materiais(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS materiais (
        id INTEGER PRIMARY KEY,
        codigo TEXT NOT NULL,
        descricao TEXT NOT NULL,
        qtd INTEGER NOT NULL,
        card TEXT NOT NULL,
        giro TEXT NOT NULL,
        UNIQUE(codigo))""")

    def drop_materiais(self):
        self.cursor.execute("DROP TABLE materiais")
        self.connect.commit()


# Parâmetros gerais na rede
class BancoParametrosRede:
    # db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\parametros.db'
    db_path = 'assets/database/rede/parametros.db'    # Arquivo local usado para testes e desenvolvimento

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_materiais(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS materiais (
        id INTEGER PRIMARY KEY,
        codigo TEXT NOT NULL,
        descricao TEXT NOT NULL,
        qtd INTEGER NOT NULL,
        card TEXT NOT NULL,
        giro TEXT NOT NULL,
        UNIQUE(codigo))""")

    def create_fornecedores(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS fornecedores (
        id INTEGER PRIMARY KEY,
        nome TEXT NOT NULL,
        tipo TEXT NOT NULL,
        UNIQUE(nome))""")

    def create_produtividade(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS produtividade (
        id INTEGER PRIMARY KEY,
        data TEXT NOT NULL,
        usuario TEXT NOT NULL,
        bancada TEXT NOT NULL,
        nota_fiscal TEXT NOT NULL,
        deposito TEXT NOT NULL,
        cod_sap TEXT NOT NULL,
        quantidade INTEGER NOT NULL,
        hora TEXT NOT NULL,
        lote TEXT NOT NULL)""")

    def create_lotes(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS lotes (
        id INTEGER PRIMARY KEY,
        data TEXT NOT NULL,
        lote TEXT NOT NULL,
        usuario TEXT NOT NULL,
        bancada TEXT NOT NULL,
        tipo_leitura TEXT NOT NULL,
        nota_fiscal TEXT NOT NULL,
        UNIQUE(lote))""")

    def gerar_lote(self, usuario, bancada, tipo_leitura, nf):
        # Definindo a data atual
        agora = datetime.now()
        data = agora.strftime('%Y-%m-%d')
        hoje = agora.strftime('%d%m%y')
        lote = f'LT{hoje}B{bancada}'
        bd_user = BancoUsuarios()

        try:
            # Conectando ao banco de dados e consultando o último número de lote da bancada
            self.cursor.execute(f"SELECT * FROM lotes WHERE data = '{data}'")
            dados = self.cursor.fetchall()

            for i in range(0, len(dados)):
                if dados[i][4] == bancada:
                    lote = dados[i][2]

            if len(lote) == 11:
                lote = f'{lote}P01'
            else:
                plt = int(lote[12:]) + 1
                lote = f'{lote[:11]}P{plt:02d}'

            # Salvando o novo número de lote na tabela de lotes
            self.cursor.execute("INSERT INTO lotes (data, lote, usuario, bancada, tipo_leitura, nota_fiscal) "
                                "VALUES (?, ?, ?, ?, ?, ?)", (data, lote, usuario, bancada, tipo_leitura, nf))
            self.connect.commit()

            bd_user.cursor.execute(f"UPDATE usuarios SET lote = '{lote}' WHERE nome = '{usuario}'")
            bd_user.connect.commit()

        except:
            self.connect.close()
            bd_user.connect.close()
            lote = None

        return lote


# TRANSFERÊNCIA (USAD) local
class BancoUsadLocal:
    db_path = 'assets/database/usad.db'

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self, nf):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS {nf} (
        id INTEGER PRIMARY KEY,
        material TEXT NOT NULL,
        serial_primario TEXT NOT NULL,
        serial_secundario TEXT NOT NULL,
        serial_number TEXT NOT NULL,
        tecnologia TEXT NOT NULL,
        estado TEXT NOT NULL,
        local TEXT NOT NULL,
        packlist TEXT NOT NULL,
        lote TEXT,
        caixa INTEGER,
        bancada TEXT,
        usuario TEXT,
        hora TEXT)""")


# TRANSFERÊNCIA (USAD) na rede
class BancoUsadRede:
    # db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\usad.db'
    db_path = 'assets/database/rede/usad.db'  # Arquivo local usado para testes e desenvolvimento

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self, nf):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS {nf} (
        id INTEGER PRIMARY KEY,
        material TEXT NOT NULL,
        serial_primario TEXT NOT NULL,
        serial_secundario TEXT NOT NULL,
        serial_number TEXT NOT NULL,
        tecnologia TEXT NOT NULL,
        estado TEXT NOT NULL,
        local TEXT NOT NULL,
        packlist TEXT NOT NULL,
        lote TEXT,
        caixa INTEGER,
        bancada TEXT,
        usuario TEXT,
        hora TEXT)""")


# REVERSA (UREV) local
class BancoUrevLocal:
    db_path = 'assets/database/urev.db'

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self, fornecedor):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS {fornecedor} (
            id INTEGER PRIMARY KEY,
            nota_fiscal TEXT NOT NULL,
            material TEXT NOT NULL,
            serial_primario TEXT NOT NULL,
            serial_secundario TEXT NOT NULL,
            serial_number TEXT NOT NULL,
            tecnologia TEXT NOT NULL,
            estado TEXT NOT NULL,
            local TEXT NOT NULL,
            packlist TEXT NOT NULL,
            lote TEXT,
            caixa INTEGER,
            bancada TEXT,
            usuario TEXT,
            hora TEXT)""")


# REVERSA (UREV) na rede
class BancoUrevRede:
    # db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\urev.db'
    db_path = 'assets/database/rede/urev.db'  # Arquivo local usado para testes e desenvolvimento

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self, fornecedor):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS {fornecedor} (
            id INTEGER PRIMARY KEY,
            nota_fiscal TEXT NOT NULL,
            material TEXT NOT NULL,
            serial_primario TEXT NOT NULL,
            serial_secundario TEXT NOT NULL,
            serial_number TEXT NOT NULL,
            tecnologia TEXT NOT NULL,
            estado TEXT NOT NULL,
            local TEXT NOT NULL,
            packlist TEXT NOT NULL,
            lote TEXT,
            caixa INTEGER,
            bancada TEXT,
            usuario TEXT,
            hora TEXT)""")


# LABORATÓRIO (USUS) local
class BancoUsusLocal:
    db_path = 'assets/database/usus.db'

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self, nf):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS {nf} (
        id INTEGER PRIMARY KEY,
        material TEXT NOT NULL,
        serial_primario TEXT NOT NULL,
        serial_secundario TEXT NOT NULL,
        serial_number TEXT NOT NULL,
        tecnologia TEXT NOT NULL,
        estado TEXT NOT NULL,
        local TEXT NOT NULL,
        lote TEXT,
        caixa INTEGER,
        bancada TEXT,
        usuario TEXT,
        hora TEXT)""")


# LABORATÓRIO (USUS) nada rede
class BancoUsusRede:
    # db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\usus.db'
    db_path = 'assets/database/rede/usus.db'  # Arquivo local usado para testes e desenvolvimento

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self, nf):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS {nf} (
        id INTEGER PRIMARY KEY,
        material TEXT NOT NULL,
        serial_primario TEXT NOT NULL,
        serial_secundario TEXT NOT NULL,
        serial_number TEXT NOT NULL,
        tecnologia TEXT NOT NULL,
        estado TEXT NOT NULL,
        local TEXT NOT NULL,
        lote TEXT,
        caixa INTEGER,
        bancada TEXT,
        usuario TEXT,
        hora TEXT)""")


# ASSINANTES (UASS) local
class BancoUassLocal:
    db_path = 'assets/database/uass.db'

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS assinantes (
        id INTEGER PRIMARY KEY,
        material TEXT NOT NULL,
        serial_primario TEXT NOT NULL,
        serial_secundario TEXT NOT NULL,
        serial_number TEXT NOT NULL,
        tecnologia TEXT NOT NULL,
        estado TEXT NOT NULL,
        local TEXT NOT NULL,
        lote TEXT,
        caixa INTEGER,
        bancada TEXT,
        usuario TEXT,
        hora TEXT)""")


# ASSINANTES (UASS) na rede
class BancoUassRede:
    # db_path = r'\\nflosv0010\FLO-TEC\SISTEMAS PLH\data\framework_in\uass.db'
    db_path = 'assets/database/rede/uass.db'  # Arquivo local usado para testes e desenvolvimento

    def __init__(self):
        self.connect = sqlite3.connect(self.db_path)
        self.cursor = self.connect.cursor()

    def create_table(self):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS assinantes (
        id INTEGER PRIMARY KEY,
        material TEXT NOT NULL,
        serial_primario TEXT NOT NULL,
        serial_secundario TEXT NOT NULL,
        serial_number TEXT NOT NULL,
        tecnologia TEXT NOT NULL,
        estado TEXT NOT NULL,
        local TEXT NOT NULL,
        lote TEXT,
        caixa INTEGER,
        bancada TEXT,
        usuario TEXT,
        hora TEXT)""")

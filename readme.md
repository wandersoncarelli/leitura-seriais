# Leitura de Seriais - TPC PLH

Sistema desenvolvido com interface gráfica, com o objetivo de realizar a triagem e 
leitura seriais de aparelhos terminais recebidos em logística reversa, transferência 
de aparelhos usados e geração de lote de aparelhos à serem enviados pro laboratório.

Todas as caixas deverão ser identificadas com  etiquetas, contendo o código do material, 
data da leitura, número da nota fiscal e número de lote gerado para o pallet. Também deverá 
conter as informações como centro, depósito, número da caixa e quantidade de terminais, 
e o mais importante, o QRCODE com os códigos dos seriais triados em cada caixa.

Durante a leitura, o usuário poderá bipar uma das três opções de seriais do aparelho, sendo o 
principal, secundário e serial number, pois o sistema irá efetuar uma consulta em tempo real
no sistema Atlas, retornando o serial principal e obtendo as informações referente ao modelo do 
aparelho. Caso o serial bipado seja inexistente, irá retornar o erro de serial inválido.

Para os seriais inválidos, duplicados, com reclassificação ou status de ASSINANTE, será exibido
uma mensagem ao usuário descrevendo o erro e irá tocar um arquivo de áudio avisando o erro. 

O sistema estará conectado com o banco de dados no servidor de rede compartilhada para 
armazenar as informações, desta forma, todo o progresso de leitura será salvo em tempo real,
evitando perdas de informações.

A produtividade dos usuários serão contabilizadas e salvas automaticamente a cada 5 minutos,
desta forma, conseguimos ter as informações do progresso de triagem atualizadas a qualquer momento.

---
## Funções:
* ### Usuários: 
  - Cadastro de usuário e senha
  - Login
  - Grupos de permissões de usuário

---
* ### Menu Principal:
  #### *• Tipos de leitura:*
  - Transferência (USAD)
  - Reversa (UREV)
  - Laboratório (USUS)
  - Assinantes (UASS)

* ### Opções de menu:
  #### *• Configurações:*
  - Alterar senha de usuário
  - Alterar senha do Atlas
  - Cadastro de materiais

  #### *• Produção Transferência:*
  - Exportar leitura
  - Importar packlist

  #### *• Produção Reversa:*
  - Cadastro de fornecedores
  - Exportação de leitura
  - Importação de packlist

  #### *• Produção Laboratório:*
  - Exportação de leitura
  
  #### *• Produção Assinantes:*
  - Exportação de leitura

  #### *• Produtividade:*
  - Consulta de produtividade
  - Exportação de relatório

  #### *• Sair:*
  - Logoff

---
* ### Importação de packlist:
  - Validação de seriais da nota fiscal no sistema Atlas
  - Extração de seriais primários, secundários e serial number
  - Extração do modelo, status e local do aparelho

* ### Exportação de leitura:
  - Unificação das leituras de todos os usuários
  - Número de lote, caixa, bancada, usuário e hora informados em cada serial
  - Exportação da leitura em planilha excel (xlsx)

* ### Interface de leitura: 
  - Cabeçalho com informações referente à nota fiscal
  - Exibição de informações referente ao material bipado
  - Lista com a quantidade consolidada de materiais bipados
  - Lista de informações detalhadas referentes aos seriais bipados
  - Geração/impressão de etiqueta QRCODE
  - Campo para informar o serial
  - Botão para apagar serial bipado
  - Botão para apagar caixa bipada
  - Botão para salvar/finalizar o lote atual

* ### Interface de produtividade:
  - Campos de seleção de intervalos de data da consulta
  - Campo de filtro de depósito, bancada, usuário e hora
  - Lista detalhada com todas as informações selecionadas nos filtros
  - Usuários operacionais visualizam apenas a própria produção
  - Usuários administradores visualizam a produção geral de todos os usuários
  - Usuários administradores podem exportar o relatório filtrado em uma planilha excel (xlsx)

---
## Imagens:
* ### Login: 
<p align="center"><img src="https://i.imgur.com/DIOhcjw.png"/></p>

* ### Menu Principal: 
<p align="center"><img src="https://i.imgur.com/S7r9tB6.png"/></p>

* ### Interface de leitura: 
<p align="center"><img src="https://i.imgur.com/PlwRUD5.png"/></p>

* ### Consulta de produtividade: 
<p align="center"><img src="https://i.imgur.com/ZfEGEAO.png"/></p>

---
### Bibliotecas utilizadas:
* Datetime
* Openpyxl
* Pandas
* Pathlib
* Playsound (v1.2.2)
* Psutil
* Pyperclip
* Pywin32
* QRCode
* Selenium
* Socket
* SQLite
* Subprocess
* Time
* Tkinter
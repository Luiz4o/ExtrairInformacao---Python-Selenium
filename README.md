﻿# Web Scraping Magazine Luiza Notebooks

Este projeto automatiza a busca de produtos no site da Magazine Luiza, separando notebooks em duas categorias:

Produtos com mais de 100 avaliações ("melhores")

Produtos com 100 ou menos avaliações ("piores")

Os resultados são salvos em um arquivo Excel e enviados automaticamente por email.

# Tecnologias Utilizadas
Python

Selenium (Web scraping)

Pandas (Manipulação de dados)

OpenPyXL (Manipulação de arquivos Excel)

SMTPLib e email (Envio de emails)

# Pré-requisitos
Antes de rodar a aplicação, certifique-se de ter instalado:

Python 3.8 ou superior

Google Chrome instalado

ChromeDriver compatível com a sua versão do Chrome

# Instalação
Clone o repositório e instale as dependências:

bash

git clone https://github.com/Luiz4o/ExtrairInformacao---Python-Selenium.git
cd ExtrairInformacao---Python-Selenium
pip install -r requirements.txt

# Configurações Necessárias

Antes de executar, defina no ambiente a variável EMAIL_PASS com uma senha que você deve gerar do seu Gmail, é super fácil basta apenas 
acessar https://myaccount.google.com/apppasswords com sua conta e irá conseguir gerar uma senha especial para usar no seu script(Para gerar 
esta senha é importante que você tenha o autentificador de duas etapas cadastrado em seu email).

Também é necessário definir variáveis de ambiente para EMAIL_FROM que seria referente ao seu email que você gerou a senha, e o EMAIL_TO que seria
para quem o email deve ser enviado.


# Como executar

No terminal, rode:

bash

python main.py

A aplicação irá:

Acessar o site da Magazine Luiza.

Pesquisar por notebooks.

Percorrer as páginas, coletando dados dos produtos.

Salvar um relatório Notebook.xlsx com:

Melhores produtos

Piores produtos

Logs de execução

Enviar o relatório por email automaticamente.

# Estrutura do Arquivo Excel
melhores: Produtos com mais de 100 avaliações.

piores: Produtos com 100 ou menos avaliações.

logs: Informações sobre eventuais erros ou falhas.

# Observações

Em caso de erro ao salvar o arquivo Excel, será gerado um arquivo notebooks_logs_error.txt contendo os logs.

Para prevenir bloqueios, o robô aguarda 8 segundos entre as páginas.

As mensagens de erro são registradas no arquivo final.

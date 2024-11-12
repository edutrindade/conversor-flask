# Conversor de Dados Firebird para Excel

Este projeto é uma aplicação web desenvolvida com **Flask** que permite a exportação de dados de um banco de dados Firebird específico para um arquivo Excel. A aplicação facilita a exportação e importação de dados para outras bases de forma prática, através de uma interface simples e intuitiva.

## Pré-requisitos

Antes de iniciar, certifique-se de ter os seguintes componentes instalados em sua máquina:

* **Python 3.x** : Linguagem de programação usada para desenvolver a aplicação.
* **pip** : Gerenciador de pacotes do Python para instalar as dependências do projeto.
* **Firebird** : Sistema de gerenciamento de banco de dados que será acessado pela aplicação.
* **Git** : Para clonar o repositório do projeto.
* **Acesso ao banco de dados Firebird** que você deseja converter.

## Instalação

### 1. Clone o Repositório

Primeiro, faça o download do código fonte clonando o repositório do projeto:

```
git clone https://github.com/seu-usuario/seu-repositorio.git

cd seu-repositorio.git
```

### 2. Crie um Ambiente Virtual

Crie um ambiente virtual para isolar as dependências do projeto:

```
python3 -m venv venv
source venv/bin/activate  # No Windows, use `venv\Scripts\activate`
```

### 3. Instale as Dependências

Com o ambiente virtual ativo, instale as dependências do projeto listadas no arquivo `requirements.txt`:

```
pip install -r requirements.txt
```

### 4. Configuração do Banco de Dados

Verifique se o banco de dados Firebird está acessível e ajuste o caminho para o banco no arquivo `app.py`, utilizando as credenciais corretas para acesso.

## Executando a Aplicação

### 1. Inicie o Servidor Flask

Ative o servidor Flask com o seguinte comando:

```
python3 app.py
```

### 2. Acesse a Aplicação

Abra um navegador e acesse a URL `http://localhost:5000` para visualizar a interface da aplicação.

### 3. Utilizando a Aplicação

* **Insira o Caminho do Banco de Dados** : Cole o caminho completo do arquivo do banco de dados Firebird no campo de entrada.
* **Exportar Dados** : Clique no botão "Exportar" para iniciar o processo de conversão. Um arquivo Excel será gerado e disponibilizado para download.

# Documentação do Projeto: Extração de Dados de Arquivos PDF

## Visão Geral

Este projeto é um código Python desenvolvido para monitorar uma pasta específica em busca de arquivos PDF e extrair informações relevantes desses arquivos. As informações são extraídas com base em palavras-chave específicas e são posteriormente formatadas e armazenadas em um arquivo Excel.

## Tecnologias Utilizadas

- Python 3.x
- Bibliotecas Python:
  - pdfplumber: Para extrair o texto de arquivos PDF.
  - watchdog: Para monitorar a pasta em busca de novos arquivos PDF.
  - openpyxl: Para criar e atualizar o arquivo Excel.
  - re: Para utilizar expressões regulares na extração de informações.

## Uso

Siga as etapas a seguir para utilizar este projeto:

### Pré-requisitos

- Certifique-se de ter o Python 3.x instalado no seu sistema.
- Instale as bibliotecas Python necessárias executando o seguinte comando:

  ```bash
  pip install pdfplumber watchdog openpyxl

## Manual do Código Python para Leitura e Extração de Dados de Arquivos PDF

Este manual descreve as principais funcionalidades do código Python fornecido para a leitura e extração de
informações de arquivos PDF. O código utiliza bibliotecas como pdfplumber, watchdog, openpyxl, e re para realizar a
tarefa. O código tem como objetivo monitorar uma pasta específica para arquivos PDF e extrair informações desses
arquivos, armazenando-as em um arquivo Excel.

### 1. Configuração Inicial:

Certifique-se de que todas as bibliotecas necessárias estejam instaladas no seu ambiente Python.

Configure o diretório onde estão localizados os arquivos PDF. Substitua o valor da variável pdf_folder pelo caminho
da sua pasta de interesse.

O código cria um arquivo Excel para armazenar as informações extraídas. Os cabeçalhos das colunas são definidos na
primeira linha da planilha.

### 2. Extração de Informações de Arquivos PDF:

O código monitora a pasta definida e extrai informações dos arquivos PDF sempre que um novo arquivo é adicionado à
pasta.

Para a extração de informações, o código procura por palavras-chave específicas nos arquivos PDF. Atualmente,
ele suporta duas palavras-chave: "DAJE" e "GRERJ".

Se a palavra-chave "DAJE" for encontrada no PDF, o código utiliza expressões regulares para localizar informações
como banco, emissor, série, número, CPF/CNPJ, valor do ato, valor a pagar, data de pagamento e código de barras. Os
valores são então formatados e adicionados ao arquivo Excel.

Se a palavra-chave "GRERJ" for encontrada no PDF, um processo semelhante é realizado para extrair informações
específicas, como banco, data de pagamento, CPF/CNPJ, valor a pagar e código de barras.

Qualquer arquivo PDF que não contenha nenhuma das palavras-chave é ignorado.

### 3. Monitoramento Contínuo:

O código utiliza a biblioteca watchdog para monitorar a pasta de PDFs continuamente. Quando um novo arquivo PDF é
adicionado à pasta, a função de manipulação é chamada para extrair informações desse arquivo.

Para encerrar o monitoramento, você pode pressionar Ctrl + C no terminal onde o código está sendo executado.

### 4. Relatório e Exportação:

Após o processamento de todos os arquivos PDF, o código gera um relatório que lista os arquivos PDF adicionados,
juntamente com a data e hora da adição. O relatório é salvo em um arquivo de texto.

O arquivo Excel contendo as informações extraídas é salvo com um nome baseado na data e hora atual.

### 5. Personalização:

Antes de executar o código, certifique-se de configurar corretamente o caminho da pasta de PDFs, bem como os detalhes
de extração específicos para o seu caso de uso.

Você pode personalizar o código para adicionar ou modificar palavras-chave, bem como ajustar as expressões regulares
para atender às suas necessidades de extração.

Lembre-se de que este código é um ponto de partida e pode ser personalizado de acordo com os requisitos específicos
do seu projeto. Certifique-se de ter as bibliotecas necessárias instaladas no seu ambiente Python e ajuste as
configurações de diretório e extração de informações conforme necessário.

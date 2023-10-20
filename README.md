
# Projeto de Extração de Dados de Arquivos PDF 🎲

## Tecnologias Utilizadas 🚀

- Python 3.x
- Bibliotecas Python: pdfplumber, watchdog, openpyxl, re

## Descrição do Projeto 📄

Este projeto consiste em um código Python que monitora uma pasta específica em busca de arquivos PDF e extrai informações relevantes desses arquivos. As informações são extraídas com base em palavras-chave específicas e são posteriormente formatadas e armazenadas em um arquivo Excel. O projeto foi desenvolvido para fins de automatização e organização de dados provenientes de documentos em formato PDF.

## Como Utilizar 👩‍💻

Siga as etapas abaixo para utilizar o projeto:

### Pré-requisitos 🧾

- Certifique-se de ter o Python 3.x instalado no seu sistema.
- Instale as bibliotecas Python necessárias. Você pode fazer isso usando o seguinte comando:

    ```bash
    pip install pdfplumber watchdog openpyxl
    ```

### Passos 📂

1. Clone ou faça o download deste repositório para o seu ambiente local.

2. Abra o código Python (arquivo `.py`) em um editor de código ou ambiente de desenvolvimento de sua escolha.

3. Configure o diretório onde estão localizados os arquivos PDF que você deseja processar, alterando a variável `pdf_folder` no código para o caminho da pasta desejada.

4. Personalize o código conforme necessário. Você pode adicionar ou modificar palavras-chave e expressões regulares para corresponder às suas necessidades de extração de informações.

5. Execute o código Python.

6. O código iniciará o monitoramento da pasta especificada. Quando um novo arquivo PDF for adicionado à pasta, o código tentará extrair informações com base nas palavras-chave especificadas.

7. As informações extraídas são formatadas e adicionadas a um arquivo Excel.

8. Após o processamento de todos os arquivos PDF, o código gera um relatório que lista os arquivos PDF processados. O relatório é salvo em um arquivo de texto.

Lembre-se de que este projeto é um ponto de partida e pode ser personalizado para atender aos requisitos específicos do seu projeto. Certifique-se de ajustar as configurações do diretório e de extração de informações de acordo com suas necessidades.

Para obter uma explicação detalhada do funcionamento do código e instruções de uso, consulte o arquivo [DOCUMENTAÇÃO.md](DOCUMENTAÇÃO.md) no repositório.

## 

 
<p align="center">
Feito com ♥ by Lorena Mattos :wave:
<a href="https://lorena-mattos.github.io/links-da-lorena/">Faça parte das minhas redes!</a>
</p> 

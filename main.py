from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime
import pdfplumber
import re
import os
import openpyxl

# Diretório onde estão localizados os PDFs
pdf_folder = r'C:\Users\lorena.machado\Documents\leitor-pdf-boletos\pdf'

# Crie um arquivo Excel
excel_file = openpyxl.Workbook()
excel_worksheet = excel_file.active

# Defina os cabeçalhos das colunas na primeira linha da planilha
excel_worksheet.append(
    ["Contribuinte", "Emissor", "Série", "Nº", "CPF/CNPJ", "VALOR DO ATO", "VALOR A PAGAR", "PAGÁVEL ATÉ",
     "Código de Barras"])

# Lista para armazenar informações sobre os PDFs adicionados
pdf_info_list = []


# Função para extrair informações de um arquivo PDF
def extract_info_and_write_to_excel(pdf_file_path):
    global excel_worksheet  # Acesso à variável global
    text_content = ""

    with pdfplumber.open(pdf_file_path) as pdf:
        for page in pdf.pages:
            text_content += page.extract_text()
            keyword_found = False

            if "DAJE" in text_content:
                banco_match = re.search(r'(BANCO|Banco) (.+?)\s\d{2}/\d{2}/\d{4}', text_content)
                emissor_match = re.search(r'Emissor (\d+)', text_content)
                serie_match = re.search(r'Série (\d+)', text_content)
                numero_match = re.search(r'Nº (\d+)', text_content)
                cpf_match = re.search(r'BANCO (.+?)\s\d{2}/\d{2}/\d{4}', text_content)
                valor_ato_match = re.search(r'REQUISIÇÃO (.+)', text_content)
                valor_pagar_match = re.search(r'CUSTAS_JUDICIAIS (.+)', text_content)
                data_pagamento_match = re.search(r'BANCO (.+)', text_content)
                codigo_de_barras_match = re.search(r'(\d{11}\s\d\s\d{11}\s\d\s\d{11}\s\d\s\d{11}\s\d)', text_content)

                if banco_match:
                    banco = banco_match.group(2)
                else:
                    banco = "Não encontrado"

                if emissor_match:
                    emissor = emissor_match.group(1)
                else:
                    emissor = "Não encontrado"

                if serie_match:
                    serie = serie_match.group(1)
                else:
                    serie = "Não encontrado"

                if numero_match:
                    numero = numero_match.group(1)
                else:
                    numero = "Não encontrado"

                if cpf_match:
                    cpf = cpf_match.group(1)
                else:
                    cpf = "Não encontrado"

                if valor_ato_match:
                    valor_ato = valor_ato_match.group(1)
                else:
                    valor_ato = "Não encontrado"

                if valor_pagar_match:
                    valor_pagar = valor_pagar_match.group(1)
                else:
                    valor_pagar = "Não encontrado"

                if data_pagamento_match:
                    data_pagamento = data_pagamento_match.group(1)
                else:
                    data_pagamento = "Não encontrado"

                if codigo_de_barras_match:
                    codigo_de_barras = codigo_de_barras_match.group(1)
                else:
                    codigo_de_barras = "Não encontrado"

                contribuinte_text = "BANCO " + banco[:-18]
                emissor_text = emissor
                serie_text = serie
                numero_text = numero
                cpf_text = cpf[-18:]
                valor_ato_text = valor_ato[35:]
                valor_pagar_text = valor_pagar[35:]
                data_pagamento_text = data_pagamento[-10:]
                codigo_de_barras_text = codigo_de_barras

                excel_worksheet.append(
                    [contribuinte_text, emissor_text, serie_text, numero_text, cpf_text, valor_ato_text,
                     valor_pagar_text, data_pagamento_text, codigo_de_barras_text])

                keyword_found = True
                print(f"Palavra-chave 'DAJE' encontrada no arquivo: {pdf_file_path}")
                # Faça o que você precisa fazer com a palavra-chave 'DAJE' aqui
                break

            if "GRERJ" in text_content:
                # Use expressões regulares para encontrar as informações
                banco_match = re.search(r'NOME DE QUEM FAZ O RECOLHIMENTO: (.+)', text_content)
                data_pagamento_match = re.search(r'VALIDADE PARA PAGAMENTO: (.+)', text_content)
                cpf_match = re.search(r'CNPJ OU CPF DE QUEM FAZ O RECOLHIMENTO: (.+)', text_content)
                valor_pagar_match = re.search(r'CAARJ / IAB (.+)', text_content)
                codigo_de_barras_match = re.search(r'(\d{11}\s\d\s\d{11}\s\d\s\d{11}\s\d\s\d{11}\s\d)', text_content)

                # Extrair as informações correspondentes
                if banco_match:
                    banco = banco_match.group(1)
                else:
                    banco = "Não encontrado"

                if data_pagamento_match:
                    data_pagamento = data_pagamento_match.group(1)
                else:
                    data_pagamento = "Não encontrado"

                if cpf_match:
                    cpf = cpf_match.group(1)
                else:
                    cpf = "Não encontrado"

                if valor_pagar_match:
                    valor_pagar = valor_pagar_match.group(1)
                else:
                    valor_pagar = "Não encontrado"

                if codigo_de_barras_match:
                    codigo_de_barras = codigo_de_barras_match.group(1)
                else:
                    codigo_de_barras = "Não encontrado"

                contribuinte_text = banco
                emissor_text = ""
                serie_text = ""
                numero_text = ""
                cpf_text = cpf
                valor_ato_text = ""
                valor_pagar_text = "R$: " + valor_pagar[26:]
                data_pagamento_text = data_pagamento[:-68]
                codigo_de_barras_text = codigo_de_barras

                excel_worksheet.append(
                    [contribuinte_text, emissor_text, serie_text, numero_text, cpf_text, valor_ato_text,
                     valor_pagar_text, data_pagamento_text, codigo_de_barras_text])

                keyword_found = True
                print(f"Palavra-chave 'GRERJ' encontrada no arquivo: {pdf_file_path}")
                # Faça o que você precisa fazer com a palavra-chave 'GRERJ' aqui
                break

    if not keyword_found:
        # Palavra-chave não encontrada, continue o monitoramento
        print(f"Palavra-chave não encontrada no arquivo: {pdf_file_path}")


# Função para lidar com eventos de novos arquivos na pasta "pdf"
class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith('.pdf') and "pdf" in event.src_path:
            pdf_file_path = event.src_path
            file_name = os.path.basename(pdf_file_path)
            added_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            added_time_cleaned = re.sub(r'\.\d+', '', added_time)  # Remove o ponto e as numerações após o segundo
            pdf_info = f"Arquivo PDF adicionado: {file_name} às {added_time_cleaned}"
            pdf_info_list.append(pdf_info)
            print(f"Arquivo PDF adicionado: {os.path.basename(pdf_file_path)} às {added_time_cleaned}")
            extract_info_and_write_to_excel(pdf_file_path)


# Inicialize o observador para monitorar a pasta
event_handler = PDFHandler()
observer = Observer()
observer.schedule(event_handler, path=pdf_folder, recursive=False)
observer.start()

try:
    while True:
        pass
except KeyboardInterrupt:
    observer.stop()

observer.join()

data_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
# Salvar as informações em um arquivo de texto
report_file_path = rf'C:\Users\lorena.machado\Documents\leitor-pdf-boletos\pdf_report_{data_atual}.txt'
with open(report_file_path, 'w') as report_file:
    for info in pdf_info_list:
        report_file.write(info + '\n')

# Salvar o arquivo Excel após o processamento de todos os arquivos PDF
excel_file.save(rf"C:\Users\lorena.machado\Documents\leitor-pdf-boletos\relatorio_{data_atual}.xlsx")

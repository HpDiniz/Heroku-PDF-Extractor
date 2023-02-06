import re
import os
import json
import uuid
import base64
import PyPDF2
import openpyxl
import pandas as pd
from io import StringIO
from datetime import datetime
from flask import Flask, request
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter

app = Flask(__name__)
pasta_temp = 'Temp/' + datetime.today().strftime('%Y%m%d') + "/"

@app.route('/obtem-ativos', methods=['POST'])
def obtem_ativos():

    if not os.path.isdir("Temp"):
        os.makedirs("Temp")
    
    if not os.path.isdir('Temp/' + datetime.today().strftime('%Y%m%d')):
        os.makedirs('Temp/' + datetime.today().strftime('%Y%m%d'))

    file_name = 'Listagem dos Ativos.xlsx'
    request_data = json.loads(request.data)

    # Cria o DataFrame resultante
    df_result = create_new_df()

    for res in request_data:

        text_pypdf, text_pdfminer = convert_base64_pdf_to_text(res['Value'])

        if(text_pypdf + text_pdfminer == ""):
            print('Arquivo ' + res['FileName'] + ' vazio.')
            continue

        tipo_extrato = get_tipo_extrato(text_pypdf, text_pdfminer)

        if tipo_extrato == "Extrato de Cotista XP":
            df_aux = obtem_extrato_cotista_xp(text_pypdf, text_pdfminer)
        elif tipo_extrato == "Posição Consolidada XP":
            df_aux = obtem_posicao_consolidada_xp(text_pypdf)
        elif tipo_extrato == "Posição e Performance XP":
            df_aux = obtem_posicao_performance_xp(text_pypdf)
        elif tipo_extrato == "Extrato Consolidado Modal":
            df_aux = obtem_extrato_consolidado_modal(text_pypdf, text_pdfminer)

        df_result = pd.concat([df_result, df_aux])

    response = {
        "Message": "Nenhum consolidado foi gerado.",
        "Value": ""
    }

    if (len(df_result) > 0):
        print(str(len(df_result)) + " ativos foram consolidados.")
        response["Message"] = "Consolidado gerado com sucesso!"
        response["Value"] = df_to_excel(df_result)

    return response

    #if (len(df_result) > 0):
    #    return df_to_excel(df_result, file_name)
    #else:
    #    return "ERROR"
    #    return {
    #        "Message": "Consolidado gerado com sucesso!",
    #        "Value": df_to_excel(df_result, file_name)
    #    }
    
    #return {
    #    "Message": "Nenhum consolidado foi gerado.",
    #    "Value": ""
    #}


# Converte o conteúdo do PDF para texto
def get_pdfminer_text(path):

    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    device = TextConverter(rsrcmgr, retstr, laparams=LAParams())
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    for page in PDFPage.get_pages(fp, set(), maxpages=0, password="",caching=True, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()

    return text.strip()

def get_pypdf_text(path):

    text, reader = "", PyPDF2.PdfReader(path)
    for page in reader.pages:
        text = text + page.extract_text()

    return text.strip()

def create_new_df():

    # Cria um DataFrame utilizado no consolidado final
    return pd.DataFrame({
        'CLASSIFICAÇÃO': pd.Series(dtype='str'),
        'ATIVO': pd.Series(dtype='str'),
        'EXPOSIÇÃO': pd.Series(dtype='str'),
        'INSTITUIÇÃO': pd.Series(dtype='str'),
        'ATUAL': pd.Series(dtype='float64'),
        'DATA': pd.Series(dtype='str'),
        'ON/OFF': pd.Series(dtype='str'),
        'DATA DE VENCIMENTO': pd.Series(dtype='str'),
        'INDEXADOR': pd.Series(dtype='str')
    })

def get_tipo_extrato(text_pypdf, text_pdfminer):

    text_1 = re.sub('\s+','', text_pypdf[0:100].lower())
    text_2 = re.sub('\s+','', text_pdfminer[0:100].lower())

    if 'extratodecotista' in text_1 and 'extratodecotista' in text_2:
        return "Extrato de Cotista XP"
    elif 'posição&performance' in text_1 and 'posição&performance' in text_2:
        return "Posição e Performance XP"
    elif 'extratoconsolidadoinvestimentos' in text_2:
        return "Extrato Consolidado Modal"

    return "Posição Consolidada XP"

def obtem_classificacao(ativo):

    # Verifica se é uma ativo no formato de Ticker
    founded = re.search(r'\b[A-Z]{4}(3|4|5|6|11|32|33|34|35)\b', ativo, flags=(re.IGNORECASE))
    if founded:
        return "Ações"
    
    # Define as palavras chave para cada Classificação
    dict_classificacoes = {
        "Ações": [
            "FIC FIA","FIA","FI Ações", "FI Açoes"
        ],
        "Renda Fixa": [
            "Renda Fixa","RF","CDB","LC","LF","LFSN","LFSC","LCA","LCI","CRA",
            "CRI","Deb","Debenture","Debênture","Tesouro","Pré","Pós","Pre",
            "Pos","IPCA","IPCA+","LTN","NTN","NTN-B","NTNB","DI","REF DI",
            "CP","Credito Privado","Crédito Privado","LP","Longo Prazo"
        ],
        "Renda Fixa (Previdência)": [
            "Previdencia","Previdência","Prev","VGBL","PGBL","VGBL/PGBL"
        ],
        "Multimercado": [
            "FIC FIM","FIM","MM","Multimercado","Multi","COE","Multiestratégia",
            "Multiestratégia"
        ]
    }

    classes = []
    # Percorre cada uma das classificações
    for classificacao in dict_classificacoes:

        # Percorre as palavras-chave de cada uma das classificações
        for key in dict_classificacoes[classificacao]:

            # Verifica se a palavra-chave atual existe no nome do ativo
            founded = re.search(r'\b' + key + r'\b', ativo, flags=(re.IGNORECASE))
            if founded and not classificacao in classes:
                classes.append(classificacao)

    if len(classes) > 0:
        return ' / '.join(classes)

    # Se nenhuma classificação foi encontrada, verifica se é um ETF de investimento estrangeiro
    founded = re.search(r'\b[A-Z]{3}\b', ativo, flags=(re.IGNORECASE))
    if founded:
        return "Ações"

    return "***"

def remove_tuple_position(result, pos):

    new_result = []
    for res in result:
        value = list(res)
        new_result.append(tuple(res[:pos] + res[pos+1:]))

    return new_result

def add_text_in_tuple_position(result, pos, text, front = True):

    new_result = []
    for res in result:
        value = list(res)
        if front:
            value[pos] = text + value[pos]
        else:
            value[pos] = value[pos] + text
        new_result.append(tuple(value))

    return new_result

def invert_tuple_positions(result, pos_1, pos_2):

    new_result = []
    for res in result:
        value = list(res)
        aux = value[pos_1]
        value[pos_1] = value[pos_2]
        value[pos_2] = aux
        new_result.append(tuple(value))

    return new_result

def reset_text_in_tuple_position(result, positions):

    if type(positions) is not list:
        positions = list(positions)

    new_result = []
    for res in result:
        value = list(res)
        for pos in positions:
            value[pos] = ""
        new_result.append(tuple(value))

    return new_result

def write_consolidado(df, regex_result, data_emissao, instituicao, classification = None):
    
    # Preenche o consolidado final
    for res in regex_result:

        # Faça atribuições
        ativo = re.sub('\s+',' ', res[0].strip())
        data_vencimento = re.sub('\s+',' ', res[1].strip())
        valor_bruto = re.sub('\s+',' ', res[3].strip())
        indexador = re.sub('^\+\s*','', re.sub('\s+',' ', res[2].strip()))

        indexador = indexador if len(indexador) > 1 else "-"
        valor_bruto = float(valor_bruto.replace('.','').replace(',','.'))
        classificacao = classification if classification != None else obtem_classificacao(ativo)
        
        # Se for encontrada alguma data no nome do ativo, algum registro inválido foi coletado
        if re.search(r'\d{2}\/\d{2}\/\d{4}', ativo):
            continue

        # Caso a data de vencimento encontrada, não seja uma data, altere seu valor para "-"
        if not re.search(r'\d{2}\/\d{2}\/\d{4}', data_vencimento):
            data_vencimento = "-"

        # Adiciona os novos valores ao DataFrame
        df = df.append({
            'CLASSIFICAÇÃO': classificacao,
            'ATIVO': ativo,
            'EXPOSIÇÃO': "Real",
            'INSTITUIÇÃO': instituicao,
            'ATUAL': valor_bruto,
            'DATA': data_emissao,
            'ON/OFF': "ON",
            'DATA DE VENCIMENTO': data_vencimento,
            'INDEXADOR': indexador
        }, ignore_index=True)
    
        df.astype({'ATUAL': 'float'}).dtypes

    return df

def convert_base64_pdf_to_text(value):

    bytes = base64.b64decode(value, validate=True)

    if bytes[0:4] != b'%PDF':
        raise ValueError('Missing the PDF file signature')

    filename = pasta_temp + str(uuid.uuid4()) + '.pdf'

    # Write the PDF contents to a local file
    f = open(filename, 'wb')
    f.write(bytes)
    f.close()

    # Lê o conteúdo do arquivo PDF
    text_pdfminer = get_pdfminer_text(filename)
    text_pypdf = get_pypdf_text(filename)

    return text_pypdf, text_pdfminer

def df_to_excel(df):

    filename = pasta_temp + str(uuid.uuid4()) + '.xlsx'

    # Criar o arquivo Excel
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, index=False, float_format="%.2f")

    # Ajustar o tamanho das colunas para a largura da maior célula em cada coluna
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for i, col in enumerate(df.columns):
        column_len = df[col].astype(str).str.len().max()
        column_len = max(column_len, len(col)) + 2
        worksheet.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = column_len

    # Salvar o arquivo Excel
    writer.close()

    with open(filename, "rb") as f:
        content = f.read()

    return base64.b64encode(content).decode("utf-8")

def obtem_extrato_consolidado_modal(text_pypdf, text_pdfminer):

    # Cria o dataFrame resultante
    df_result = create_new_df()

    # Regex para obter a Data de Emissão
    regex = "(?<=Per.odo de refer.ncia \d{2}\/\d{2}\/\d{4} a )\d{2}\/\d{2}\/\d{4}"
    data_emissao = re.search(regex, text_pdfminer, flags=(re.IGNORECASE)).group(0)

    # Regex para obter todas as rendas fixas
    regex = r'\n[\d\/]*([a-zA-Z].+?)\s*\d{2}\/\d{2}\/\d{4}\s*(\d{2}\/\d{2}\/\d{4})\s*R\$\s*[\d,.]+[.,]\d{2}\s*[\d.]+(\s*)R\$\s*([\d,.]+[.,]\d{2})\s*R\$\s*[\d,.]+[.,]\d{2}\s*[\d.,]+'
    result = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    # Regex para obter todas as ações
    regex = r'(BRUTO|\,\d\d)\s*([a-zA-Z].+?\-(.|\n)*?[A-Z]{4}(3|4|5|6|11|32|33|34|35))\s*[\d.]+\s*R\$\s*[\d.]+\,\d\d\s*R\$\s*([\d.]+\,\d\d)'
    result_acoes = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))
    result = result + remove_tuple_position(result_acoes,0)
    
    # Regex para obter todas os CRAs
    regex = r'(L.QUIDO|\,\d\d)\s*([a-zA-Z].+?\-(.|\n)*?)\d{2}\/\d{2}\/\d{4}\s*[\d.]+\s*R\$\s*([\d.]+\,\d\d)\s*R\$\s*[\d.]+\,\d\d\s*R\$\s*[\d.]+\,\d\d'
    result_cra = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))
    result_cra = reset_text_in_tuple_position(result_cra,[0,2])
    result = result + invert_tuple_positions(result_cra,0,1)

    df_result = write_consolidado(df_result, result, data_emissao, "Modal")

    return df_result

def obtem_posicao_performance_xp(text_pypdf):

    # Cria o dataFrame resultante
    df_result = create_new_df()

    # Regex para obter todas as linhas que possuírem ativos de Renda Fixa
    regex = r'\s*(.*)\d{2}\/\d{2}\/\d{4}\s*(\d{2}\/\d{2}\/\d{4}|\-)\s*(\d{2}\/\d{2}\/\d{4})([A-Z\-\+\s]*\s*[\d,.]+%[A-Z\-\s]*)([\d.]+\,\d\d)[\d.]+\,\d\d'
    result = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    # Regex para obter todas as linhas que possuírem ativos de Renda Fixa
    regex = r'\s*(.*)\d{2}\/\d{2}\/\d{4}[\s\d.]+\,\d\d([\d.]+\,\d\d)[\d.]+\,\d\d'
    result = result + re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    return df_result

def obtem_posicao_consolidada_xp(text_pypdf):

    # Cria o dataFrame resultante
    df_result = create_new_df()

    # Regex para obter a Data de Emissão
    regex = "(?<=Data da Consulta: )\d{2}\/\d{2}\/\d{4}"
    data_emissao = re.search(regex, text_pypdf, flags=(re.IGNORECASE)).group(0)

    # Regex para obter todas as linhas que possuírem ativos de Renda Fixa
    regex = r'\n[\d\/]*([a-zA-Z].+?\s*-\s*[A-Z]{3}\/\d{4})\s*\d{2}\/\d{2}\/\d{4}\s*\d{2}\/\d{2}\/\d{4}\s*(\d{2}\/\d{2}\/\d{4})\s*([A-Z\-\+\s]*\s*[\d,.]+%[A-Z\-\s]*)\d+\s+\d+\s*R\$\s*[\d,.]+[.,]\d{2}\s*R\$\s*([\d,.]+[.,]\d{2})\s*R\$\s*[\d,.]+[.,]\d{2}'
    result = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    # Regex para obter todas as linhas que possuírem ativos de Renda Fixa sem carência
    regex = r'\n[\d\/]*([a-zA-Z].+?\s*-\s*[A-Z]{3}\/\d{4})\s*\d{2}\/\d{2}\/\d{4}\s*\-\s*(\d{2}\/\d{2}\/\d{4})\s*([A-Z\-\+\s]*\s*[\d,.]+%[A-Z\-\s]+)\d+\s+\d+\s*R\$\s*[\d,.]+[.,]\d{2}\s*R\$\s*([\d,.]+[.,]\d{2})\s*R\$\s*[\d,.]+[.,]\d{2}'
    result = result + re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    # Regex para obter todas as linhas que possuírem ativos de Renda Fixa Pós-Fixados
    regex = r'\n[\d\/]*([a-zA-Z].+?)\s*\d{2}\/\d{2}\/(\d{4})\s*[\d,.]+\s*[\d,.]+\s*R\$\s*[\d,.]+(\s*)R\$\s*([\d,.]+)\s*R\$\s*[\d,.]+'
    result = result + re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    # Regex para obter todas as linhas que possuírem Fundos Imobiliários
    regex = r'([A-Z]{4}(3|4|5|6|11|32|33|34|35))\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+R\$\s*[\d,.]+[.,]\d{2}(\s*)R\$\s*([\d,.]+[.,]\d{2})'
    result = result + re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    # Regex para obter todas as linhas que possuírem COEs
    regex = r'\n[\d\/]*([a-zA-Z].+?)\s*-\s*[\w\s,.]*\s*-\s*\s*\d{2}\.\d{2}\.\d{4}\s*[\w\s,.]*\d{2}\/\d{2}\/\d{4}\s+(\d{2}\/\d{2}\/\d{4})\s+\d+\s*R\$\s*[\d,.]+[.,]\d{2}\s*R\$\s*[\d,.]+[.,]\d{2}(\s*)R\$\s*([\d,.]+[.,]\d{2})'
    result_coes = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))
    result_coes = add_text_in_tuple_position(result_coes, 0, "COE - ")

    # Regex para obter todas as linhas que possuírem Ações
    regex = r'([A-Z]{4}(3|4|5|6|11|32|33|34|35))\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+R\$\s*[\d,.]+[.,]\d{2}(\s*)R\$\s*([\d,.]+[.,]\d{2})'
    result_acoes = re.findall(regex, text_pypdf, flags=(re.IGNORECASE | re.MULTILINE))

    df_result = write_consolidado(df_result, result, data_emissao, "XP Investimentos")
    df_result = write_consolidado(df_result, result_coes, data_emissao, "XP Investimentos", "Multimercado")
    df_result = write_consolidado(df_result, result_acoes, data_emissao, "XP Investimentos")

    return df_result

def obtem_extrato_cotista_xp(text_pypdf, text_pdfminer):

    # Cria o dataFrame resultante
    df_result = create_new_df()

    # Regex para obter a data de emissão
    regex_0 = "(?<=Movimenta..o de \d{2}\/\d{2}\/\d{4} a )\d{2}\/\d{2}\/\d{4}"
    data_emissao = re.search(regex_0, text_pdfminer, flags=(re.IGNORECASE)).group(0)

    # Regex para obter o texto situado entre o termo "POSIÇÃO CONSOLIDADA" e o termo "Emissão"
    regex_1 = "(?<=POSI..O CONSOLIDADA).*?(?=Emissão:)"
    result_1 = re.search(regex_1, text_pdfminer, flags=(re.IGNORECASE | re.DOTALL)).group(0)

    # Regex para obter todas as linhas que possuírem pelo menos uma letra e um espaço
    regex_2 = "^.*[a-zA-Z] .*$"
    result_2 = re.finditer(regex_2, result_1, flags=(re.MULTILINE))

    # Regex para obter o texto situado entre o termo "POSIÇÃOCONSOLIDADA" e o termo "TotalnaInstituição"
    regex_3 = "(?<=POSI..OCONSOLIDADA\n).*?(?=TotalnaInstituição)"
    result_3 = re.search(regex_3, text_pypdf, flags=(re.IGNORECASE | re.DOTALL)).group(0)

    # Substitui os espaços por um ponto e vírgula, para simular uma tabela CSV
    result_3 = result_3.replace(" ",";")

    # Corrige termos com espaço
    for result in result_2:
        value = result.group(0)
        key = value.replace(" ","")
        result_3 = result_3.replace(key,value)

    # Converte o CSV para o tipo DataFrame
    df = pd.read_csv(StringIO(result_3), sep=";")
    
    # Cria um array de tuplas
    result = []

    # Preenche o consolidado final
    for index, row in df.iterrows():
        result.append((row['Fundo'],"-","-",row["Valor Bruto"]))

    df_result = write_consolidado(df_result, result, data_emissao, "XP Investimentos")

    return df_result

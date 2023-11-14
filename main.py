import streamlit as st
import pandas as pd
from openpyxl.styles import Font
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

import requests
from PIL import Image
from io import BytesIO, StringIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

st.title("Formatar arquivos RNCAC")

file_upload = st.file_uploader("Anexe o arquivo original baixado do monday", type='xlsx')

if file_upload is not None:
    # Salva o arquivo temporário
    temp_file = 'arquivo_formatado.xlsx'
    # temp_file = r"C:\Users\pcp2\Downloads\RNCAC_CQ_1698677290 (3).xlsx"

    with open(temp_file, 'wb') as file:
        file.write(file_upload.getvalue())

    font_bold = Font(bold=True)

    excelOriginal = pd.read_excel(temp_file)

    id = excelOriginal['Unnamed: 2'][4] #B7
    # item = excelOriginal['RNCAC'][4] 
    # log_criacao = excelOriginal['Unnamed: 3'][4]
    # origem = excelOriginal['Unnamed: 4'][4]
    setor = excelOriginal['Unnamed: 5'][4] #D7
    # conjunto = excelOriginal['Unnamed: 6'][4] 
    responsavel = excelOriginal['Unnamed: 9'][4] #G8 G15
    # gravidade = excelOriginal['Unnamed: 10'][4]
    # prioridade = excelOriginal['Unnamed: 11'][4]
    # status = excelOriginal['Unnamed: 12'][4] 
    # ultima_atualizacao = excelOriginal['Unnamed: 16'][4]

    descricao = excelOriginal['Unnamed: 7'][4] #A11
    arquivos = excelOriginal['Unnamed: 7'][6]
    arquivos_split = arquivos.split()
    foto = arquivos_split[0].replace(",","")

    import requests

    img_data = requests.get(foto).content
    with open('image_name.jpg', 'wb') as handler:
        handler.write(img_data)


    # conclusao = excelOriginal['Unnamed: 14'][4]
    acao_contencao = excelOriginal['Unnamed: 13'][4] #A17


    mao_de_obra = excelOriginal['Unnamed: 17'][4] #B25
    materia_prima = excelOriginal['Unnamed: 18'][4] #B26
    medicao = excelOriginal['Unnamed: 19'][4] #B27
    metodo = excelOriginal['Unnamed: 20'][4] #B28
    meio_ambiente = excelOriginal['Unnamed: 21'][4] #B29

    participantes = excelOriginal['Unnamed: 23'][4] #C30

    excelOriginal_cortado = excelOriginal.iloc[5:].reset_index(drop=True)
    # excelOriginal_cortado = excelOriginal_cortado.set_axis(excelOriginal_cortado.iloc[0], axis='columns', inplace=False)
    excelOriginal_cortado = excelOriginal_cortado.set_axis(excelOriginal_cortado.iloc[0], axis='columns')
    excelOriginal_cortado = excelOriginal_cortado[1:].reset_index(drop=True)
    
    print(excelOriginal_cortado)


    wb = load_workbook('modelo_final_v2.xlsx')
    ws = wb.active

    ws['B7'] = id
    ws['D7'] = setor
    ws['G8'] = responsavel
    ws['G14'] = responsavel
    ws['A10'] = descricao
    ws['A16'] = acao_contencao
    ws['B24'] = mao_de_obra
    ws['B25'] = materia_prima
    ws['B26'] = medicao
    ws['B27'] = metodo
    ws['B28'] = meio_ambiente
    ws['C29'] = participantes

    # Adiciona a imagem à planilha do Excel
    excel_img = ExcelImage(img_data)
    ws.add_image(excel_img, 'A1')

    # ws['B14'].font = font_bold

    # ws['B15'] = prioridade
    # ws['B15'].font = font_bold

    # ws['B16'] = status
    # ws['B16'].font = font_bold

    # ws['A20'] = descricao
    # ws['A18'] = arquivos
    # ws['D45'] = ultima_atualizacao

    # ws['A42'] = conclusao
    # ws['D44'] = avaliacao

    excelOriginal_cortado['Previsão'] = pd.to_datetime(excelOriginal_cortado['Previsão'])
    excelOriginal_cortado['Previsão'] = excelOriginal_cortado['Previsão'].dt.strftime('%d-%m-%Y')
    excelOriginal_cortado = excelOriginal_cortado.fillna('')

    u = 34
    ultimaLinha = len(excelOriginal_cortado) + u-1

    df_status = excelOriginal_cortado.iloc[:,5:6]

    try:

        while u < ultimaLinha:
            for i in range(len(excelOriginal_cortado)):
                ws['A' + str(u)] = excelOriginal_cortado['Name'][i]
                ws['F' + str(u)] = excelOriginal_cortado['Previsão'][i]
                ws['E' + str(u)] = excelOriginal_cortado['Responsável'][i]
                u += 1
    except:
        pass

    wb.template = False
    wb.save(temp_file)

    # Crie um botão de download
    download_button = st.download_button(
        label="Clique aqui para baixar o arquivo formatado",
        data=open(temp_file, 'rb').read(),
        file_name=temp_file,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # Remove o arquivo temporário
    import os
    os.remove(img_path)
    
    # Exibe o botão de download
    st.write(download_button)

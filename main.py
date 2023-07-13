import streamlit as st
import pandas as pd
from openpyxl.styles import Font
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

st.title("Formatar arquivos RNCAC")

file_upload = st.file_uploader("Anexe o arquivo original baixado do monday", type='xlsx')

if file_upload is not None:
    # Salva o arquivo temporário
    temp_file = 'arquivo_formatado.xlsx'
    
    with open(temp_file, 'wb') as file:
        file.write(file_upload.getvalue())

    font_bold = Font(bold=True)

    excelOriginal = pd.read_excel(temp_file)

    id = excelOriginal['Unnamed: 2'][4]
    item = excelOriginal['RNCAC'][4]
    log_criacao = excelOriginal['Unnamed: 3'][4]
    origem = excelOriginal['Unnamed: 4'][4]
    setor = excelOriginal['Unnamed: 5'][4]
    conjunto = excelOriginal['Unnamed: 6'][4]
    responsavel = excelOriginal['Unnamed: 9'][4]
    gravidade = excelOriginal['Unnamed: 10'][4]
    prioridade = excelOriginal['Unnamed: 11'][4]
    status = excelOriginal['Unnamed: 12'][4]
    ultima_atualizacao = excelOriginal['Unnamed: 17'][4]

    descricao = excelOriginal['Unnamed: 7'][4]
    arquivos = excelOriginal['Unnamed: 8'][4]

    excelOriginal_cortado = excelOriginal.iloc[5:].reset_index(drop=True)
    # excelOriginal_cortado = excelOriginal_cortado.set_axis(excelOriginal_cortado.iloc[0], axis='columns', inplace=False)
    excelOriginal_cortado = excelOriginal_cortado.set_axis(excelOriginal_cortado.iloc[0], axis='columns')
    excelOriginal_cortado = excelOriginal_cortado[1:].reset_index(drop=True)

    wb = load_workbook('modeloFinal.xlsx')
    ws = wb.active

    ws['B3'] = id
    ws['B4'] = item
    ws['B5'] = log_criacao
    ws['B6'] = origem
    ws['B7'] = setor
    ws['B8'] = conjunto
    ws['B9'] = responsavel
    ws['B10'] = gravidade
    ws['B10'].font = font_bold

    ws['B11'] = prioridade
    ws['B11'].font = font_bold

    ws['B12'] = status
    ws['B12'].font = font_bold

    ws['A14'] = descricao
    ws['A16'] = arquivos
    ws['D41'] = ultima_atualizacao

    excelOriginal_cortado['Prazo'] = pd.to_datetime(excelOriginal_cortado['Prazo'])
    excelOriginal_cortado['Prazo'] = excelOriginal_cortado['Prazo'].dt.strftime('%d-%m-%Y')
    excelOriginal_cortado = excelOriginal_cortado.fillna('')

    ultimaLinha = len(excelOriginal_cortado) + 19 - 2
    u = 19

    df_status = excelOriginal_cortado.iloc[:,5:6]

    try:

        while u < ultimaLinha:
            for i in range(len(excelOriginal_cortado)):
                ws['A' + str(u)] = excelOriginal_cortado['Name'][i]
                ws['B' + str(u)] = excelOriginal_cortado['Ação Imediata'][i]
                ws['C' + str(u)] = excelOriginal_cortado['Prazo'][i]
                ws['D' + str(u)] = excelOriginal_cortado['Pessoas'][i]
                ws['E' + str(u)] = df_status['Status'][i]
                ws['F' + str(u)] = excelOriginal_cortado['Avaliação'][i]
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

    # Exibe o botão de download
    st.write(download_button)

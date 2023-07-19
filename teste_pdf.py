from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
import pandas as pd

def calcular_altura_tabela(dados_tabela, largura_colunas, altura_linha):
    num_linhas = len(dados_tabela)
    altura_tabela = altura_linha * num_linhas
    return altura_tabela

nome_arquivo = 'relatorio_new.pdf'

def criar_cabecalho(nome_arquivo, infoIniciais, infoTabelaprincipal):
    # Inicializa o objeto Canvas e cria o arquivo PDF
    pdf = canvas.Canvas(nome_arquivo, pagesize=letter)

    # Carrega a imagem
    # imagem = "logoCemag.png"  # Substitua pelo caminho da sua imagem
    # pdf.drawImage(ImageReader(imagem), 72, 730, width=70, height=70)

    # Define o conteúdo do cabeçalho
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(150, 750, "Relatório de Não Conformidades e Ações Corretivas")
    pdf.setFont("Helvetica", 10)
    pdf.drawString(250, 730, "Gestão da Qualidade")
    pdf.setFont("Helvetica", 10)
    
    pdf.drawString(460, 750, "Código: RQ GQ-014-000")
    pdf.drawString(460, 740, "Data de Emissão: 7/14/2023")
    pdf.drawString(460, 730, "Data da Última Revisão: N/A")
    pdf.drawString(460, 720, "Página: 1/1")

    def quebrar_texto(texto, tamanho_limite):
        palavras = texto.split(', ')
        partes = []
        parte_atual = ""
        
        for palavra in palavras:
            if not parte_atual:
                parte_atual = palavra
            elif len(parte_atual) + len(palavra) + 2 <= tamanho_limite:  # +2 para considerar a vírgula e o espaço
                parte_atual += f', {palavra}'
            else:
                partes.append(parte_atual)
                parte_atual = palavra
        
        if parte_atual:
            partes.append(parte_atual)
        
        return partes

    # Exemplo de texto longo para a célula
    texto_longo = 'Gestão da Qualidade, Gestão de Pessoas, Almoxarifado, Compras, Controle da Qualidade, Engenharia de Produtos, Manutenção, Marketing, PCP, Adm de Vendas, SESMT, Produção, Vendas Externas'

    # Tamanho limite para cada parte do texto
    tamanho_limite = 90

    # Quebrar o texto em partes menores
    partes_do_texto = quebrar_texto(texto_longo, tamanho_limite)

    texto = ''
    # Exibir as partes resultantes
    for i, parte in enumerate(partes_do_texto, start=1):
        texto += ",\n" + parte

    # Dados da tabela
    dados_tabela = [
        ["Item:", infoIniciais[0]],
        ["Log de criação:", infoIniciais[1]],
        ["Origem:", infoIniciais[2]],
        ["Setor:", infoIniciais[3]],
        # Adicione mais linhas conforme necessário
    ]

    # Configurações da tabela
    largura_colunas = [100, 430]
    altura_linha = 21

    # Cria a tabela
    table1 = Table(dados_tabela, colWidths=largura_colunas, rowHeights=altura_linha)

    # Estilo da tabela
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), 'white'),  # Cor de fundo para o cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), 'black'),  # Cor do texto do cabeçalho
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),    # Alinhamento central para todas as células
        ('BACKGROUND', (0, 1), (-1, -1), 'white'),  # Cor de fundo para as outras células
        ('GRID', (0, 0), (-1, -1), 1, 'black'),  # Adiciona as grades da tabela
    ])

    table1.setStyle(estilo_tabela)

    x, y = 50, 630
    
    # Adiciona a tabela ao PDF
    table1.wrapOn(pdf, 0, 0)
    table1.drawOn(pdf, x, y)

    # Dados da tabela
    dados_tabela = [
        ["Conjunto ou Atividade:", texto],
        # Adicione mais linhas conforme necessário
    ]

    # Configurações da tabela
    largura_colunas = [100, 430]
    altura_linha = 39

    # Cria a tabela
    table1 = Table(dados_tabela, colWidths=largura_colunas, rowHeights=altura_linha)

    # Estilo da tabela
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), 'white'),  # Cor de fundo para o cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), 'black'),  # Cor do texto do cabeçalho
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),    # Alinhamento central para todas as células
        ('BACKGROUND', (0, 1), (-1, -1), 'white'),  # Cor de fundo para as outras células
        ('GRID', (0, 0), (-1, -1), 1, 'black'),  # Adiciona as grades da tabela
    ])

    table1.setStyle(estilo_tabela)

    x, y = 50, 591
    
    # Adiciona a tabela ao PDF
    table1.wrapOn(pdf, 0, 0)
    table1.drawOn(pdf, x, y)

    # Dados da tabela
    dados_tabela = [
        ["Responsável:", infoIniciais[6]],
        ["Gravidade:", infoIniciais[7]],
        ["Prioridade:", infoIniciais[8]],
        ["Status:", infoIniciais[9]],
        ["Descrição:", " "],
        ["Arquivos:", " "],
        # Adicione mais linhas conforme necessário
    ]

    # Configurações da tabela
    largura_colunas = [100, 430]
    altura_linha = 21

    # Cria a tabela
    table1 = Table(dados_tabela, colWidths=largura_colunas, rowHeights=altura_linha)

    # Estilo da tabela
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), 'white'),  # Cor de fundo para o cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), 'black'),  # Cor do texto do cabeçalho
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),    # Alinhamento central para todas as células
        ('BACKGROUND', (0, 1), (-1, -1), 'white'),  # Cor de fundo para as outras células
        ('GRID', (0, 0), (-1, -1), 1, 'black'),  # Adiciona as grades da tabela
    ])

    table1.setStyle(estilo_tabela)

    x, y = 50, 465
    
    # Adiciona a tabela ao PDF
    table1.wrapOn(pdf, 0, 0)
    table1.drawOn(pdf, x, y)

    # Transformar o DataFrame em uma lista de listas
    dados_tabela = infoTabelaprincipal.values.tolist()

    # Adicionar o cabeçalho à lista de listas
    cabecalho = ["Setores", "Ação Corretiva", "Prazo", "Pessoas", "Status", "Avaliação"]
    dados_tabela.insert(0, cabecalho)

    lista_dados = ''  # Inicializa a variável como uma string vazia

    # Exibir a lista de listas resultante
    for linha in dados_tabela:
        lista_dados += str(linha)  # Concatena a linha e adiciona uma quebra de linha

    # Dados da tabela
    dados_tabela = [
        lista_dados

        # Adicione mais linhas conforme necessário
    ]

    # Configurações da tabela
    largura_colunas = [50, 150, 60, 150, 70, 50]
    altura_linha = 30

    # Cria a tabela
    table = Table(dados_tabela, colWidths=largura_colunas, rowHeights=altura_linha)

    # Estilo da tabela
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), 'grey'),  # Cor de fundo para o cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), 'white'),  # Cor do texto do cabeçalho
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),    # Alinhamento central para todas as células
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Fonte negrito para o cabeçalho
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),   # Espaçamento inferior para o cabeçalho
        ('BACKGROUND', (0, 1), (-1, -1), 'white'),  # Cor de fundo para as outras células
        ('GRID', (0, 0), (-1, -1), 1, 'black'),  # Adiciona as grades da tabela
    ])

    table.setStyle(estilo_tabela)

    # Calcula a altura total ocupada pela tabela
    altura_tabela = calcular_altura_tabela(dados_tabela, largura_colunas, altura_linha)

    x, y = 50, 460 - altura_tabela 

    # Adiciona a tabela ao PDF
    table.wrapOn(pdf, 0, 0)
    table.drawOn(pdf, x, y)


    # Dados da tabela
    dados_tabela = [
        ["Conclusão:", " "],

        # Adicione mais linhas conforme necessário
    ]

    # Configurações da tabela
    largura_colunas = [100, 430]
    altura_linha = 30

    # Cria a tabela
    table1 = Table(dados_tabela, colWidths=largura_colunas, rowHeights=altura_linha)

    # Estilo da tabela
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), 'white'),  # Cor de fundo para o cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), 'black'),  # Cor do texto do cabeçalho
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),    # Alinhamento central para todas as células
        ('BACKGROUND', (0, 1), (-1, -1), 'white'),  # Cor de fundo para as outras células
        ('GRID', (0, 0), (-1, -1), 1, 'black'),  # Adiciona as grades da tabela
    ])

    table1.setStyle(estilo_tabela)

    x, y = 50, 95
    
    # Adiciona a tabela ao PDF
    table1.wrapOn(pdf, 0, 0)
    table1.drawOn(pdf, x, y)

    # Adiciona os campos abaixo da tabela
    pdf.drawString(50, 40, "Avaliação:")
    pdf.drawString(50, 20, "Última atualização:")

    # Fecha o objeto Canvas
    pdf.save()

excelOriginal = pd.read_excel(r'RNCAC_1688058432.xlsx')

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

infoIniciais = [id,item,log_criacao,origem,setor,conjunto,responsavel,
                gravidade,prioridade,status,ultima_atualizacao, descricao,
                arquivos]

excelOriginal_cortado = excelOriginal.iloc[5:].reset_index(drop=True)
# excelOriginal_cortado = excelOriginal_cortado.set_axis(excelOriginal_cortado.iloc[0], axis='columns', inplace=False)
excelOriginal_cortado = excelOriginal_cortado.set_axis(excelOriginal_cortado.iloc[0], axis='columns')
excelOriginal_cortado = excelOriginal_cortado[1:].reset_index(drop=True)

excelOriginal_cortado['Prazo'] = pd.to_datetime(excelOriginal_cortado['Prazo'])
excelOriginal_cortado['Prazo'] = excelOriginal_cortado['Prazo'].dt.strftime('%d-%m-%Y')
excelOriginal_cortado = excelOriginal_cortado.fillna('')

excelOriginal_cortado = excelOriginal_cortado[['Name','Ação Imediata','Prazo','Pessoas','Status','Avaliação']]
excelOriginal_cortado_parte1 = excelOriginal_cortado.iloc[:,:5]
excelOriginal_cortado_parte2 = excelOriginal_cortado.iloc[:,7:8]

excelOriginal_cortado_final = pd.concat([excelOriginal_cortado_parte1, excelOriginal_cortado_parte2], axis=1)

infoTabelaprincipal = excelOriginal_cortado_final 

pdf_buffer = criar_cabecalho(infoIniciais, infoTabelaprincipal)


if __name__ == "__main__":
    nome_arquivo = "relatorio.pdf"
    criar_cabecalho(nome_arquivo)

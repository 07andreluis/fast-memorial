import pandas as pd
from docx import Document

def gerar_memorial_descritivo(caminho_planilha, caminho_modelo, caminho_saida):
    """
    Gera um memorial descritivo em formato de texto a partir de uma planilha de dados,
    incluindo as coordenadas de cada vértice.

    Args:
        caminho_planilha (str): Caminho para o arquivo de dados (e.g., 'dados.xlsx').
        caminho_modelo (str): Caminho para o modelo do memorial (e.g., 'modelo.docx').
        caminho_saida (str): Caminho para salvar o documento final (e.g., 'memorial_final.docx').
    """
    try:
        df = pd.read_excel(caminho_planilha)
    except FileNotFoundError:
        print(f"Erro: O arquivo de planilha não foi encontrado em {caminho_planilha}")
        return

    # Extração das informações gerais (assumindo que estão na primeira linha)
    nome_imovel = df.loc[0, 'Nome Imóvel']
    proprietario = df.loc[0, 'Proprietario']
    area = df.loc[0, 'Area']
    matricula = df.loc[0, 'Matricula']
    perimetro_total = df.loc[0, 'Perimetro Total']

    # Geração do Documento com python-docx
    doc = Document(caminho_modelo)

    # Substitui os placeholders no documento
    for p in doc.paragraphs:
        if '[NOME_IMOVEL]' in p.text:
            p.text = p.text.replace('[NOME_IMOVEL]', str(nome_imovel))
        if '[PROPRIETARIO]' in p.text:
            p.text = p.text.replace('[PROPRIETARIO]', str(proprietario))
        if '[AREA]' in p.text:
            p.text = p.text.replace('[AREA]', str(area))
        if '[MATRICULA]' in p.text:
            p.text = p.text.replace('[MATRICULA]', str(matricula))
        if '[PERIMETRO_TOTAL]' in p.text:
            p.text = p.text.replace('[PERIMETRO_TOTAL]', str(perimetro_total))

    # Adiciona a descrição do perímetro em parágrafos
    doc.add_paragraph() # Adiciona uma linha em branco para separar
    doc.add_paragraph("DESCRIÇÃO")

    # Inserindo o primeiro parágrafo de descrição (início da descrição)
    primeiro_vertice = df.loc[0, 'Vertice']
    coord_N_primeiro = df.loc[0, 'coord_N']
    coord_E_primeiro = df.loc[0, 'coord_E']
    confrontante_primeiro = df.loc[0, 'Confrontante']
    
    doc.add_paragraph(f"Inicia-se a descrição deste perímetro no vértice {primeiro_vertice}, de coordenadas N {coord_N_primeiro}m e E {coord_E_primeiro}m; situado no limite com {confrontante_primeiro};")

    # Itera sobre as linhas da planilha para criar os parágrafos de cada trecho
    for index, row in df.iterrows():
        # Lógica para criar a descrição do trecho
        azimute_atual = row['Azimute']
        distancia_atual = row['Distancia']
        confrontante_atual = row['Confrontante']
        
        # O próximo vértice, se houver
        if index + 1 < len(df):
            proximo_vertice = df.loc[index + 1, 'Vertice']
            # Obtém as coordenadas do próximo vértice
            coord_N_proximo = df.loc[index + 1, 'coord_N']
            coord_E_proximo = df.loc[index + 1, 'coord_E']

            desc_paragrafo = f"deste, segue com azimute e distância de: {azimute_atual} e {distancia_atual}m, confrontando neste trecho com {confrontante_atual} até o vértice {proximo_vertice}, de coordenadas N {coord_N_proximo}m e E {coord_E_proximo}m;"
        else:
            # Caso seja o último ponto, ele volta para o primeiro
            desc_paragrafo = f"deste, segue com azimute e distância de: {azimute_atual} e {distancia_atual}m, confrontando neste trecho com {confrontante_atual} até o vértice {primeiro_vertice}, ponto inicial da descrição deste perímetro."

        # Adiciona o parágrafo ao documento
        doc.add_paragraph(desc_paragrafo)

    # Adiciona o parágrafo final sobre o sistema de coordenadas
    doc.add_paragraph("Todas as coordenadas aqui descritas estão demarcadas e encontram-se representadas no Sistema UTM (Universal Transversa de Mercator), referenciadas ao Meridiano Central 45° W (Fuso 23), tendo como o Datum o SIRGAS-2000.")

    # Salva o documento final
    doc.save(caminho_saida)
    print(f"Memorial descritivo gerado com sucesso em '{caminho_saida}'.")

# Exemplo de uso
caminho_da_planilha = "dados_memorial.xlsx"
caminho_do_modelo = "modelo_memorial.docx"
caminho_de_saida = "memorial_final.docx"

gerar_memorial_descritivo(caminho_da_planilha, caminho_do_modelo, caminho_de_saida)
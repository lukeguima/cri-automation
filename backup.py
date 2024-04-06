from odf.opendocument import load, OpenDocumentText
from odf import text, teletype

def extrair_dados_arquivo_odt(arquivo_odt):
    doc = load(arquivo_odt)
    dados = {}

    # Lista de tópicos esperados no arquivo .odt
    topicos_esperados = [
        'LIVRO',
        'FOLHA',
        'NÚMERO DE ORDEM',
        'ANO',
        'DATA',
        'LOCALIDADE:',
        'DENOMINAÇÃO OU RUA:',
        'CONFRONTAÇÕES E CARACTERÍSTICAS:',
        'ADQUIRENTE:',
        'TRANSMITENTE:',
        'TÍTULO:',
        'FORMA DO TÍTULO:',
        'VALOR:',
        'CONDIÇÕES:',
        'AVERBAÇÕES:'
    ]

    # Variável para rastrear o índice atual na lista de tópicos
    indice_topico_atual = 0

    # Percorrer todos os parágrafos no documento .odt
    for para in doc.getElementsByType(text.P):
        # Extrair texto do parágrafo
        texto = teletype.extractText(para).strip()

        # Verificar se o texto do parágrafo corresponde ao próximo tópico esperado
        if texto.startswith(topicos_esperados[indice_topico_atual]):
            # Extrair o valor do tópico e armazená-lo nos dados
            valor_topico = texto[len(topicos_esperados[indice_topico_atual]):].strip()
            dados[topicos_esperados[indice_topico_atual]] = valor_topico

            # Mover para o próximo tópico esperado
            indice_topico_atual += 1

            # Se todos os tópicos foram encontrados, interromper a busca
            if indice_topico_atual == len(topicos_esperados):
                break

    return dados

def construir_linha_dados(dados):
    # Construir a linha de dados com os valores de LIVRO, FOLHA, NÚMERO DE ORDEM, ANO e DATA
    linha_dados = f"ro{dados['LIVRO']} Folha{dados['FOLHA']} Ordem{dados['NÚMERO DE ORDEM']} Data{dados['DATA']}"
    return linha_dados

def substituir_texto_para_marcador(texto, marcador, valor):
    # Substituir todas as ocorrências do marcador pelo valor
    return texto.replace(marcador, valor)

def inserir_dados_no_template(dados, arquivo_template, arquivo_resultado):
    # Carregar o arquivo de template
    template = load(arquivo_template)

    linha_dados = construir_linha_dados(dados)

    # Percorrer todos os parágrafos no documento do template
    for elem in template.getElementsByType(text.P):
        # Extrair texto do parágrafo
        texto = teletype.extractText(elem).strip()

        # Substituir o marcador 'INJECT' pela linha de dados concatenada
        if 'Liv' in texto:
            novo_texto = texto.replace('Liv', linha_dados)
            elem.addText(novo_texto)

        # Extrair o valor do marcador 'FORMA DO TÍTULO:' dos dados fornecidos
        forma_do_titulo = dados.get('FORMA DO TÍTULO:', '')

    # Percorrer todos os parágrafos no documento do template
    for elem in template.getElementsByType(text.P):
        # Extrair texto do parágrafo
        texto = teletype.extractText(elem).strip()

        # Verificar se o marcador é 'FORMA DO TÍTULO:' e realizar a substituição
        if 'FORMA DO TÍTULO:' in texto:
            # Substituir o marcador pelo valor extraído
            novo_texto = texto.replace('FORMA DO TÍTULO:', forma_do_titulo)
            elem.text = ''
            elem.addText(novo_texto)
        else:
            # Percorrer todos os marcadores e valores para substituição
            for marcador, valor in dados.items():
                if marcador in texto:
                    # Substituir o marcador pelo valor correspondente
                    novo_texto = texto.replace(marcador, valor)
                    # Limpar o conteúdo do parágrafo
                    elem.text = ''
                    # Adicionar o novo texto ao elemento de parágrafo
                    elem.addText(novo_texto)
                    break
    # Salvar o documento resultante
    template.save(arquivo_resultado)


if __name__ == "__main__":
    arquivo_dados = "arquivo.odt"
    arquivo_template = "template.odt"
    arquivo_resultado = "resultado.odt"

    # Extrair dados do arquivo .odt
    dados = extrair_dados_arquivo_odt(arquivo_dados)

    # Exibir os dados extraídos
    print("Dados extraídos do arquivo:")
    print(dados)

    # Inserir os dados extraídos no arquivo de template e salvar o resultado
    inserir_dados_no_template(dados, arquivo_template, arquivo_resultado)

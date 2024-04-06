from odf.opendocument import load, OpenDocumentText
from odf import text, teletype

def extrair_dados_arquivo_odt(arquivo_odt):
    doc = load(arquivo_odt)
    dados = {}

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

    indice_topico_atual = 0

    for para in doc.getElementsByType(text.P):
        texto = teletype.extractText(para).strip()

        if texto.startswith(topicos_esperados[indice_topico_atual]):
            valor_topico = texto[len(topicos_esperados[indice_topico_atual]):].strip()
            dados[topicos_esperados[indice_topico_atual]] = valor_topico

            indice_topico_atual += 1

            if indice_topico_atual == len(topicos_esperados):
                break

    return dados

def construir_linha_dados(dados):
    linha_dados = f"ro{dados['LIVRO']} Folha{dados['FOLHA']} Ordem{dados['NÚMERO DE ORDEM']} Data{dados['DATA']}"
    return linha_dados

def substituir_texto_para_marcador(texto, marcador, valor):
    return texto.replace(marcador, valor)

def inserir_dados_no_template(dados, arquivo_template, arquivo_resultado):
    template = load(arquivo_template)

    linha_dados = construir_linha_dados(dados)

    for elem in template.getElementsByType(text.P):
        texto = teletype.extractText(elem).strip()

        if 'Liv' in texto:
            novo_texto = texto.replace('Liv', linha_dados)
            elem.addText(novo_texto)

        forma_do_titulo = dados.get('FORMA DO TÍTULO:', '')

    for elem in template.getElementsByType(text.P):
        texto = teletype.extractText(elem).strip()

        if 'FORMA DO TÍTULO:' in texto:
            novo_texto = texto.replace('FORMA DO TÍTULO:', forma_do_titulo)
            elem.text = ''
            elem.addText(novo_texto)
        else:
            for marcador, valor in dados.items():
                if marcador in texto:
                    novo_texto = texto.replace(marcador, valor)
                    elem.text = ''
                    elem.addText(novo_texto)
                    break
    template.save(arquivo_resultado)


if __name__ == "__main__":
    arquivo_dados = "arquivo.odt"
    arquivo_template = "template.odt"
    arquivo_resultado = "resultado.odt"

    dados = extrair_dados_arquivo_odt(arquivo_dados)

    print("Transcrição Completa")

    inserir_dados_no_template(dados, arquivo_template, arquivo_resultado)

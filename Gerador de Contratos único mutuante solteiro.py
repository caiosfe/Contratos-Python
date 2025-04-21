import pandas as pd
from docx import Document


def fill_invitation(template_path, output_path, data):
    # Carrega o documento template
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        paragraph_text = paragraph.text  # Texto completo do parágrafo
        for key, value in data.items():
            if key in paragraph_text:  # Verifica se o placeholder está presente no texto do parágrafo
                # Realiza a substituição no texto completo do parágrafo
                paragraph_text = paragraph_text.replace(key, value)

        # Atualiza os `runs` do parágrafo para os novos valores
        if paragraph.text != paragraph_text:
            for run in paragraph.runs:
                run.text = ""  # Limpa o texto do run atual
            paragraph.runs[0].text = paragraph_text  # Define o texto completo modificado no primeiro run

    # Salva o arquivo com as substituições realizadas
    doc.save(output_path)


def generate_invitation_from_csv(csv_path, template_path):
    try:
        # Carrega o CSV usando o separador especificado
        df = pd.read_csv(csv_path, sep=';')
    except Exception as e:
        print(f"Erro ao carregar o arquivo CSV: {e}")
        return

    for idx, row in df.iterrows():
        # Cria um dicionário com os dados para substituição
        data = {
            '[unidade]': str(row['unidade']),
            '[tipo]': str(row['tipo']),
            '[metro]': str(row['metro']),
            '[metrext]': str(row['metrext']),
            '[preço]': str(row['preço']),
            '[preçoext]': str(row['preçoext']),
            '[nome1]': str(row['nome1']),
            '[nac1]': str(row['nac1']),
            '[ec1]': str(row['ec1']),
            '[prof1]': str(row['prof1']),
            '[cpf1]': str(row['cpf1']),
            '[rg1]': str(row['rg1']),
            '[end]': str(row['end']),
            '[tel1]': str(row['tel1']),
            '[email1]': str(row['email1']),
            '[data]': str(row['data'])
        }

        # Define o nome do arquivo de saída
        output_path = f'{data['[unidade]']}.BRID-Contrato de mútuo.docx'

        # Realiza a substituição no template
        fill_invitation(template_path, output_path, data)


if __name__ == '__main__':
    # Define o caminho do template e do arquivo CSV
    csv_path = 'Informações.csv'
    template_path = 'Contrato mutuante solteiro.docx'

    # Gera os documentos
    generate_invitation_from_csv(csv_path, template_path)

# ProcessCard.py
from CleanValue import clean_value

# Função para processar os cartões da Matriz de Cursos
def process_card(card, headers):
    field_values = {header: "" for header in headers}

    for field in card['fields']:
        if field['name'] in headers:
            value = field['value']
            if value is None:
                field_values[field['name']] = ""
            else:
                try:
                    field_values[field['name']] = int(value)
                except ValueError:
                    try:
                        field_values[field['name']] = float(value.replace(',', '.'))
                    except ValueError:
                        field_values[field['name']] = clean_value(field['value'])

    return field_values

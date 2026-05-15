import xmltodict
import csv


file_path = 'Report.htm'

try:
    with open(file_path, 'r', encoding='utf-8') as file:
        xml_content = file.read()
    
    data_dict = xmltodict.parse(xml_content)

    divs = data_dict['html']['body']['div']
    
    tabela = []
    

    header = ['left', 'top', 'width', 'height', 'conteudo']
    
    for item in divs:
        style_str = item.get('@style', '')
 
        style_dict = {}
   
        for part in style_str.strip(';').split(';'):
            if ':' in part:
                key, value = part.split(':', 1)
                style_dict[key.strip()] = value.strip()
        
        conteudo = ''
        if '#text' in item:
            conteudo = item['#text'].strip()
        elif isinstance(item.get('div'), dict) and '#text' in item['div']:
            conteudo = item['div']['#text'].strip()


        if 'left' in style_dict and 'top' in style_dict:
            tabela.append({
                'left': float(style_dict.get('left').replace('px', '')),
                'top': float(style_dict.get('top').replace('px', '')),
                'width': float(style_dict.get('width', '0px').replace('px', '')),
                'height': float(style_dict.get('height', '0px').replace('px', '')),
                'conteudo': conteudo
            })

    tabela_ordenada = sorted(tabela, key=lambda x: (x['top'], x['left']))
    
    csv_file_name = 'relatorio_estruturado.csv'
    
    with open(csv_file_name, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=header)
        writer.writeheader()
        
        linhas_csv = []
        if tabela_ordenada:
            linha_atual = []
            top_anterior = tabela_ordenada[0]['top']
            for item in tabela_ordenada:
 
                if abs(item['top'] - top_anterior) > 2:
                    if linha_atual:
                        linhas_csv.append(linha_atual)
                    linha_atual = []
                linha_atual.append(item['conteudo'])
                top_anterior = item['top']
            
            if linha_atual:
                linhas_csv.append(linha_atual)

        for linha in linhas_csv:

            row_dict = {header[i]: linha[i] for i in range(min(len(linha), len(header)))}
            writer.writerow(row_dict)
    
    print(f"Sucesso! O arquivo '{csv_file_name}' foi criado com os dados estruturados.")

except FileNotFoundError:
    print(f"Erro: O arquivo '{file_path}' não foi encontrado.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
from openpyxl import load_workbook, Workbook  # Para trabalhar com arquivos Excel
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment  # Para aplicar estilos nas células
from datetime import datetime  # Para manipular datas
import unicodedata
from pathlib import Path  # Para manipulação de caminhos de arquivos
from utils import (  # Estilos e funções auxiliares
    cabeçalho_fill, cabeçalho_font, enviado_fill, enviado_font,
    semtecnico_fill, atrasado_fill, validado_nao_fill, validado_sim_fill,
    cores_regionais, bordas, alinhamento,
    corrigir_acentuacao, normalizar_texto
)

# Define o caminho do script atual
caminho_script = Path(__file__).resolve()
pasta_scripts = caminho_script.parent
pasta_form1 = pasta_scripts.parent / "form1"

# Caminho do arquivo do banco e arquivos auxiliares (originais do drive)
excel_file_input = pasta_form1 / "planilhas_consumo/form1.xlsx"  
planilhas_auxiliares = {
    "belem": pasta_form1 / "planilhas_consumo/belem.xlsx",    
    "expansao": pasta_form1 / "planilhas_consumo/expansao.xlsx", 
    "grs": pasta_form1 / "planilhas_consumo/GRS.xlsx"
}

# Carrega a planilha principal
wb_input = load_workbook(excel_file_input)
ws_input = wb_input.active

# Dicionário para armazenar o status de envio por município
dados_atualizados = {}

# Percorre as linhas da planilha principal (a partir da segunda)
for row in ws_input.iter_rows(min_row=2, values_only=True):
    municipio = row[1]
    data_envio = row[2]

    # Normaliza o nome do município
    if isinstance(municipio, str):
        municipio_normalizado = normalizar_texto(corrigir_acentuacao(municipio))
    else:
        continue

    # Tenta formatar a data de envio
    if isinstance(data_envio, datetime):
        data_envio_formatada = data_envio.strftime("%d/%m/%Y")
    else:
        try:
            data_envio_formatada = datetime.strptime(data_envio, "%m/%d/%Y %I:%M:%S %p").strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            data_envio_formatada = ""

    # Atualiza ou insere os dados no dicionário
    if municipio_normalizado in dados_atualizados:
        dados_atualizados[municipio_normalizado]["datas"].append(data_envio_formatada)
        dados_atualizados[municipio_normalizado]["status"] = "Envio Duplicado"
    else:
        dados_atualizados[municipio_normalizado] = {"datas": [data_envio_formatada], "status": "Enviado"}

# Processa cada planilha auxiliar (belém, expansão, GRS)
for nome, caminho in planilhas_auxiliares.items():
    wb_aux = load_workbook(caminho)

    # Verifica se a aba existe
    if "Form 1 - Município" not in wb_aux.sheetnames:
        print(f"A aba 'Form 1 - Município' não foi encontrada em {nome}. Nenhuma modificação será feita.")
        continue

    ws_aux = wb_aux["Form 1 - Município"]

    # Cria nova planilha para salvar os dados atualizados
    novo_wb = Workbook()
    novo_ws = novo_wb.active
    novo_ws.title = nome

    # Copia e estiliza os cabeçalhos
    headers = [cell.value for cell in ws_aux[1]]
    for col_num, header in enumerate(headers, start=1):
        cell = novo_ws.cell(row=1, column=col_num, value=header)
        cell.fill = cabeçalho_fill
        cell.font = cabeçalho_font
        cell.border = bordas
        cell.alignment = alinhamento  

    # Processa as linhas da planilha auxiliar
    for row_idx, row in enumerate(ws_aux.iter_rows(min_row=2, values_only=True), start=2):
        municipio_original = row[1]
        row_data = list(row)

        # Normaliza o nome do município
        if isinstance(municipio_original, str):
            municipio_normalizado = normalizar_texto(corrigir_acentuacao(municipio_original))
            print(municipio_normalizado)
        else:
            municipio_normalizado = ""

        # Atualiza status e data de envio, se estiver no dicionário
        if municipio_normalizado in dados_atualizados:
            novas_datas = ", ".join(dados_atualizados[municipio_normalizado]["datas"])
            novo_status = dados_atualizados[municipio_normalizado]["status"]
            row_data[5] = novas_datas  # Coluna de data
            row_data[4] = novo_status  # Coluna de status
        else:
            # Caso não tenha envio:
            if row_data[4] == "Sem Técnico":
                pass
            elif row_data[4] is None:
                row_data[4] = "Atrasado"

        ### ESTILIZAÇÃO 

        # Coloração de validação (Sim/Não)
        if row_data[6] == "Não":
            novo_ws.cell(row=row_idx, column=7).fill = validado_nao_fill
        elif row_data[6] == "Sim":
            novo_ws.cell(row=row_idx, column=7).fill = validado_sim_fill

        # Coloração regional
        regional = row_data[0]
        if regional in cores_regionais:
            cor_hex = cores_regionais[regional]
            novo_ws.cell(row=row_idx, column=1).fill = PatternFill(start_color=cor_hex, end_color=cor_hex, fill_type="solid")

        # Copia os dados para a nova planilha e aplica borda/alinhamento
        for col_idx, value in enumerate(row_data, start=1):
            cell = novo_ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = bordas
            cell.alignment = alinhamento
            cell.font = Font(name='Arial', size=11)

    # Aplica cor ao status 
    for row_idx in range(2, novo_ws.max_row + 1):
        status_cell = novo_ws.cell(row=row_idx, column=5)
        status = status_cell.value
        if status == "Enviado" or status == "Envio Duplicado":
            status_cell.fill = enviado_fill
            status_cell.font = enviado_font
        elif status == "Sem Técnico":
            status_cell.fill = semtecnico_fill
            status_cell.font = enviado_font
        elif status == "Atrasado":
            status_cell.fill = atrasado_fill
            status_cell.font = enviado_font

    # Ajusta a largura das colunas com base no conteúdo
    for col in novo_ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = col[0].column_letter
        novo_ws.column_dimensions[col_letter].width = max_length + 5

    # Salva a nova planilha com nome específico
    novo_caminho = pasta_form1 / f"{nome}_atualizado_form1.xlsx"
    novo_wb.save(novo_caminho)
    print(f"{novo_caminho} gerada com sucesso")

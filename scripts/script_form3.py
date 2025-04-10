from openpyxl import load_workbook, Workbook  # Para carregar e criar arquivos Excel
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment  # Para aplicar estilos nas células
from datetime import datetime  # Para lidar com datas
import unicodedata  # Para normalização de texto (acentos etc.)
from pathlib import Path  # Para manipulação de caminhos de arquivos
from utils import (  # Importa funções utilitárias e estilos personalizados
    cabeçalho_fill, cabeçalho_font, enviado_fill, enviado_font,
    semtecnico_fill, atrasado_fill, validado_nao_fill, validado_sim_fill,
    cores_regionais, bordas, alinhamento,
    corrigir_acentuacao, normalizar_texto, normalizar_uvr
)

# Caminho da pasta onde o script está localizado
caminho_script = Path(__file__).resolve()
pasta_scripts = caminho_script.parent
pasta_form3 = pasta_scripts.parent / "form3"

# Caminho da planilha principal de entrada
excel_file_input = pasta_form3 / "planilhas_consumo/form3.xlsx" 

# Caminhos das planilhas auxiliares
planilhas_auxiliares = {
    "belem": pasta_form3 / "planilhas_consumo/belem.xlsx",
    "expansao": pasta_form3 / "planilhas_consumo/expansao.xlsx",
    "grs": pasta_form3 / "planilhas_consumo/GRS.xlsx"
}

# Carrega o arquivo principal
wb_input = load_workbook(excel_file_input)
ws_input = wb_input.active  # Seleciona a aba ativa

dados_atualizados = {}  # Dicionário para armazenar informações atualizadas

# Percorre as linhas da planilha principal (a partir da segunda)
for row in ws_input.iter_rows(min_row=2, values_only=True):
    municipio = row[5]
    uvr_nro = row[4]  
    data_envio = row[6]
    cnpj = row[2]  

    # Gera a chave normalizada (municipio + UVR)
    if isinstance(municipio, str):
        municipio_uvr_normalizado = f"{normalizar_texto(corrigir_acentuacao(municipio))}_{normalizar_uvr(uvr_nro)}"
    else:
        continue  

    # Formata a data de envio
    if isinstance(data_envio, datetime):
        data_envio_formatada = data_envio.strftime("%d/%m/%Y")
    else:
        try:
            data_envio_formatada = datetime.strptime(data_envio, "%m/%d/%Y %I:%M:%S %p").strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            data_envio_formatada = ""

    # Cria a entrada no dicionário se ainda não existir
    if municipio_uvr_normalizado not in dados_atualizados:
        dados_atualizados[municipio_uvr_normalizado] = {"cnpjs": {}, "status": "Enviado"}

    # Adiciona ou atualiza as datas de envio por CNPJ
    if cnpj not in dados_atualizados[municipio_uvr_normalizado]["cnpjs"]:
        dados_atualizados[municipio_uvr_normalizado]["cnpjs"][cnpj] = [data_envio_formatada]
    else:
        dados_atualizados[municipio_uvr_normalizado]["cnpjs"][cnpj].append(data_envio_formatada)
        dados_atualizados[municipio_uvr_normalizado]["status"] = "Envio Duplicado"  # Marca como duplicado

# Processa cada uma das planilhas auxiliares (Belém, GRS e Expansão)
for nome, caminho in planilhas_auxiliares.items():
    wb_aux = load_workbook(caminho)

    # Verifica se a aba necessária existe
    if "Form 3 - Empreendimento" not in wb_aux.sheetnames:
        print(f"A aba 'Form 3 - Empreendimento' não foi encontrada em {nome}. Nenhuma modificação será feita.")
        continue

    ws_aux = wb_aux["Form 3 - Empreendimento"]

    # Cria uma nova planilha para os dados atualizados
    novo_wb = Workbook()
    novo_ws = novo_wb.active
    novo_ws.title = nome

    # Copia o cabeçalho da planilha original e aplica estilos
    headers = [cell.value for cell in ws_aux[1]]
    for col_num, header in enumerate(headers, start=1):
        cell = novo_ws.cell(row=1, column=col_num, value=header)
        cell.fill = cabeçalho_fill
        cell.font = cabeçalho_font
        cell.border = bordas
        cell.alignment = alinhamento

    # Processa cada linha da planilha auxiliar
    for row_idx, row in enumerate(ws_aux.iter_rows(min_row=2, values_only=True), start=2):
        municipio_original = row[1]
        uvr_nro_original = row[2]  
        row_data = list(row)

        # Normaliza a chave para buscar nos dados atualizados
        if isinstance(municipio_original, str):
            municipio_uvr_normalizado = f"{normalizar_texto(corrigir_acentuacao(municipio_original))}_{normalizar_uvr(uvr_nro_original)}"
            print(municipio_uvr_normalizado, ' aux')
        else:
            municipio_uvr_normalizado = ""

        # Atualiza as colunas de status e datas, se dados foram encontrados
        if municipio_uvr_normalizado in dados_atualizados:
            novas_datas = ", ".join(
                sum(dados_atualizados[municipio_uvr_normalizado]["cnpjs"].values(), [])
            )
            novo_status = dados_atualizados[municipio_uvr_normalizado]["status"]
            row_data[5] = novas_datas  # Data de envio
            row_data[4] = novo_status  # Status
        else: 
            if row_data[4] == "Sem Técnico":
                pass
            elif row_data[4] == None:
                row_data[4] = "Atrasado"

        # Colore a célula da validação
        if row_data[6] == "Não":
            novo_ws.cell(row=row_idx, column=7).fill = validado_nao_fill
        elif row_data[6] == "Sim":
            novo_ws.cell(row=row_idx, column=7).fill = validado_sim_fill

        # Aplica cor regional
        regional = row_data[0]
        if regional in cores_regionais:
            cor_hex = cores_regionais[regional]
            novo_ws.cell(row=row_idx, column=1).fill = PatternFill(start_color=cor_hex, end_color=cor_hex, fill_type="solid")

        # Escreve os dados na nova planilha com estilos
        for col_idx, value in enumerate(row_data, start=1):
            cell = novo_ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = bordas
            cell.alignment = alinhamento
            cell.font = Font(name='Arial', size=11)

    # Aplica cores de preenchimento conforme o status
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

    # Ajusta automaticamente a largura das colunas
    for col in novo_ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = col[0].column_letter
        novo_ws.column_dimensions[col_letter].width = max_length + 5

    # Salva a nova planilha atualizada
    novo_caminho = pasta_form3 / f"{nome}_atualizado_form3.xlsx"
    novo_wb.save(novo_caminho)
    print(f"{novo_caminho} gerada com sucesso")

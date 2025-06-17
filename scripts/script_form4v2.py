# Importações necessárias
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from datetime import datetime
from pathlib import Path
import pandas as pd
from datetime import timedelta
from utils import (
    cabeçalho_fill, cabeçalho_font, enviado_fill, enviado_font,
    semtecnico_fill, atrasado_fill, validado_nao_fill, validado_sim_fill, duplicado_fill, outras_fill, atrasado2_fill,
    cores_regionais, bordas, alinhamento,
    normalizar_texto, normalizar_uvr, aplicar_estilo_status
)

# Define os caminhos dos arquivos envolvidos
caminho_script = Path(__file__).resolve()
pasta_scripts = caminho_script.parent
pasta_inputs = pasta_scripts.parent / "inputs"

# Caminho do arquivo do banco e arquivos auxiliares (originais do drive)
csv_file_input = pasta_inputs/"form4.csv"  
planilhas_auxiliares = {
    "belem": pasta_inputs / "0 - Belém" / "0 - Monitoramento Form 4.xlsx",
    "expansao": pasta_inputs / "0 - Expansão" / "0 - Monitoramento Form 4.xlsx",
    "grs": pasta_inputs / "0 - GRS II" / "0 - Monitoramento Form 4.xlsx"
}

# Carrega a planilha principal
df_input = pd.read_csv(csv_file_input, dtype=str)

# Dicionários para armazenar os dados extraídos
dados_atualizados = {}
div_por_municipio = {}
regionais_por_municipio = {}

# Converte a data de referência para o formato "MM.AA"
def converter_data_para_mes_ano(data_referencia):
    if isinstance(data_referencia, datetime):
        return data_referencia.strftime("%m.%y")
    else:
        try:
            return datetime.strptime(data_referencia, "%Y-%m-%d").strftime("%m.%y")
        except (ValueError, TypeError):
            return ""

# Remove caracteres inválidos do nome da aba
def limpar_nome_aba(nome):
    return nome.replace("/", "-").replace("\\", "-").replace(":", "-").replace("*", "-").replace("?", "-").replace("[", "").replace("]", "")


for _, row in df_input.iterrows():
    municipio = row['gm_nome']
    uvr_nro = row['guvr_numero']
    data_envio = row['data_de_envio']
    tc_uvr = row['nome_tc_uvr']  
    data_referencia = row['data_de_referencia']


    if isinstance(municipio, str):
        municipio_uvr_normalizado = f"{normalizar_texto(municipio)}_{uvr_nro}"
    else:
        continue

    mes_ano = converter_data_para_mes_ano(data_referencia)

    # Tenta formatar a data de envio
    if isinstance(data_envio, datetime):
        data_envio_formatada = data_envio.strftime("%d/%m/%Y")
    else:
        try:
            data_envio_formatada = datetime.strptime(data_envio, "%Y-%m-%d").strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            data_envio_formatada = ""

    # Agrupa dados por município + UVR + mês/ano
    chave = (municipio_uvr_normalizado, mes_ano)
    if chave in dados_atualizados:
        dados_atualizados[chave]["datas_envio"].append(data_envio_formatada)
        dados_atualizados[chave]["status"] = "Duplicado"
    else:
        dados_atualizados[chave] = {
            "datas_envio": [data_envio_formatada],
            "status": "Enviado",
            "municipio_original": municipio,
            "uvr_nro": uvr_nro,
            "mes_ano": mes_ano,
            "tc_uvr" : tc_uvr
        }

# Cria um novo workbook para cada planilha auxiliar
wb_final = {nome: Workbook() for nome in planilhas_auxiliares}
for nome in wb_final:
    wb_final[nome].remove(wb_final[nome].active)

# Processa cada planilha auxiliar
for nome, caminho in planilhas_auxiliares.items():
    wb_aux = load_workbook(caminho)

    for aba in wb_aux.sheetnames:
        # Só processa abas no formato MM.AA
        if aba.count('.') == 1 and all(x.isdigit() for x in aba.split('.')):
            mes_ano_aux = aba
            ws_aux = wb_aux[aba]

            mes_ano_limpo = limpar_nome_aba(mes_ano_aux)
            if mes_ano_limpo not in wb_final[nome].sheetnames:
                wb_final[nome].create_sheet(title=mes_ano_limpo)

            ws_final = wb_final[nome][mes_ano_limpo]

            # Copia cabeçalhos com formatação
            headers = [cell.value for cell in ws_aux[1]]
            for col_num, header in enumerate(headers, start=1):
                cell = ws_final.cell(row=1, column=col_num, value=header)
                cell.fill = cabeçalho_fill
                cell.font = cabeçalho_font
                cell.border = bordas
                cell.alignment = alinhamento

            # Calcula mês/ano atual
            hoje = datetime.today()
            mes_atual = hoje.month
            ano_atual = hoje.year

            # Função auxiliar para calcular a diferença de meses
            def diferenca_em_meses(ano_alvo, mes_alvo, ano_base, mes_base):
                return (ano_base - ano_alvo) * 12 + (mes_base - mes_alvo)

            # Processa linhas de dados
            for row_idx, row in enumerate(ws_aux.iter_rows(min_row=2, values_only=True), start=2):
                regional = row[0]
                municipio_original = row[1]
                uvr_nro_original = row[2]
                row_data = list(row)

                if not isinstance(municipio_original, str) or not municipio_original.strip():
                    continue

                municipio_uvr_normalizado = f"{normalizar_texto(municipio_original)}_{normalizar_uvr(uvr_nro_original)}"
                chave_busca = (municipio_uvr_normalizado, mes_ano_aux)
                div_por_municipio[municipio_uvr_normalizado] = nome
                regionais_por_municipio[municipio_uvr_normalizado] = regional

                situacao_atual = row_data[4]
                tem_envio_existente = bool(row_data[5]) and isinstance(row_data[5], str) and row_data[5].strip()

                # Atualiza dados conforme a planilha principal
                if chave_busca in dados_atualizados:
                    row_data[5] = ", ".join(dados_atualizados[chave_busca]["datas_envio"])
                    row_data[4] = dados_atualizados[chave_busca]["status"]
                elif not tem_envio_existente:
                    if situacao_atual not in ("UVR Sem Técnico", "Outras Ocorrências"):
                        try:
                            aba_mes, aba_ano = map(int, mes_ano_aux.split("."))
                            aba_ano += 2000
                        except:
                            aba_mes, aba_ano = None, None

                        if aba_ano and aba_mes:
                            diff = diferenca_em_meses(aba_ano, aba_mes, ano_atual, mes_atual)
                            if diff == 1:
                                row_data[4] = "Atrasado"
                            elif diff >= 2:
                                row_data[4] = "Atrasado >= 2"

                # Estiliza as linhas
                for col_idx, value in enumerate(row_data, start=1):
                    cell = ws_final.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = bordas
                    cell.alignment = alinhamento
                    cell.font = Font(name='Arial', size=11)

                # Aplica cor para célula de validação
                if row_data[6] == "Não":
                    ws_final.cell(row=row_idx, column=7).fill = validado_nao_fill
                elif row_data[6] == "Sim":
                    ws_final.cell(row=row_idx, column=7).fill = validado_sim_fill

                # Aplica cor da regional
                if regional in cores_regionais:
                    cor_hex = cores_regionais[regional]
                    ws_final.cell(row=row_idx, column=1).fill = PatternFill(start_color=cor_hex, end_color=cor_hex, fill_type="solid")

            # Aplica estilo com base no status
            for row_idx in range(2, ws_final.max_row + 1):
                status_cell = ws_final.cell(row=row_idx, column=5)
                status = status_cell.value
                aplicar_estilo_status(status_cell, status)

            # Ajusta largura das colunas
            for col in ws_final.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                col_letter = col[0].column_letter
                wb_final[nome][mes_ano_limpo].column_dimensions[col_letter].width = max_length + 5

# Cria aba "irregulares" com registros que não se encaixam nas abas mensais
for nome, wb in wb_final.items():
    aba_irregulares = wb.create_sheet("irregulares")
    
    # Cabeçalho da aba
    colunas_base = ["Regional", "Município", "UVR", "Técnico de UVR", "Situação", "Data de Envio", "Mês de referência"]
    for col_num, col_name in enumerate(colunas_base, start=1):
        cell = aba_irregulares.cell(row=1, column=col_num, value=col_name)
        cell.fill = cabeçalho_fill
        cell.font = cabeçalho_font
        cell.border = bordas
        cell.alignment = alinhamento

    linha_atual = 2
    for chave, info in dados_atualizados.items():
        municipio_uvr, mes_ano = chave
        if mes_ano not in wb.sheetnames:
            if div_por_municipio.get(municipio_uvr) == nome:
                nova_linha = [
                    regionais_por_municipio.get(municipio_uvr, ""),
                    info["municipio_original"],
                    info["uvr_nro"],
                    info["tc_uvr"],
                    info["status"],
                    ", ".join(info["datas_envio"]),
                    mes_ano
                ]
                for col_idx, valor in enumerate(nova_linha, start=1):
                    cell = aba_irregulares.cell(row=linha_atual, column=col_idx, value=valor)
                    cell.border = bordas
                    cell.alignment = alinhamento
                    cell.font = Font(name='Arial', size=11)

                # Cor da regional
                regional_cell = aba_irregulares.cell(row=linha_atual, column=1)
                regional_nome = regional_cell.value
                if regional_nome in cores_regionais:
                    cor_hex = cores_regionais[regional_nome]
                    regional_cell.fill = PatternFill(start_color=cor_hex, end_color=cor_hex, fill_type="solid")

                # Cor do status
                status_cell = aba_irregulares.cell(row=linha_atual, column=5)
                status = status_cell.value
                aplicar_estilo_status(status_cell, status)

                linha_atual += 1

    # Ajusta largura das colunas
    for col in aba_irregulares.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = col[0].column_letter
        aba_irregulares.column_dimensions[col_letter].width = max_length + 5

df_input['data_de_referencia'] = pd.to_datetime(df_input['data_de_referencia'], errors='coerce')
df_input['receita_vendas'] = pd.to_numeric(df_input['receita_vendas'], errors='coerce')
df_input['mun_uvr'] = df_input['gm_nome'].apply(normalizar_texto) + "_" + df_input['guvr_numero']

data_hoje = pd.Timestamp.today() #coleta a data atual pra poder calcular o desvio dos ultimos 6 meses
inicio_mes_atual = pd.Timestamp(year=data_hoje.year, month=data_hoje.month, day=1)
data_minima = inicio_mes_atual - pd.DateOffset(months=6)

df_filtrado = df_input.dropna(subset=['data_de_referencia', 'receita_vendas'])
df_filtrado = df_filtrado[df_filtrado['data_de_referencia'] >= data_minima]

registros_outliers = []

for mun_uvr, grupo in df_filtrado.groupby('mun_uvr'):
    grupo = grupo.sort_values('data_de_referencia')
    print(grupo)
    media = grupo['receita_vendas'].mean()
   

    for _, atual in grupo.iterrows():
        valor = atual['receita_vendas']
        if media == 0:
            continue  # evita divisão por zero

        desvio_percentual = abs((valor - media) / media) #calcula o desvio

        if valor == 0 or desvio_percentual > 0.80: #se o desvio percentual for maior que 80%
            registros_outliers.append({
                "Município": atual['gm_nome'],
                "UVR": atual['guvr_numero'],
                "Técnico UVR": atual['nome_tc_uvr'],
                "Data de Referência": atual['data_de_referencia'].strftime("%m.%Y"),
                "Receita de Vendas": valor,
                "Média 6 meses": round(media, 2),
                "Desvio (%)": f"{round(desvio_percentual * 100, 2)}%"
            })
            

for nome, wb in wb_final.items():
    aba_outliers = wb.create_sheet("outliers")
    colunas = [
        "Município",
        "UVR",
        "Técnico UVR",
        "Data de Referência",
        "Receita de Vendas",
        "Média 6 meses",
        "Desvio (%)"
    ]

    for col_num, nome_col in enumerate(colunas, start=1):
        cell = aba_outliers.cell(row=1, column=col_num, value=nome_col)
        cell.fill = cabeçalho_fill
        cell.font = cabeçalho_font
        cell.border = bordas
        cell.alignment = alinhamento

    linha = 2
    for registro in registros_outliers:
        mun_uvr_chave = f"{normalizar_texto(registro['Município'])}_{registro['UVR']}"
        if div_por_municipio.get(mun_uvr_chave) == nome:
            for col_idx, chave in enumerate(colunas, start=1):
                valor = registro[chave]
                cell = aba_outliers.cell(row=linha, column=col_idx, value=valor)
                cell.border = bordas
                cell.alignment = alinhamento
                cell.font = Font(name='Arial', size=11)
            linha += 1

    for col in aba_outliers.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = col[0].column_letter
        aba_outliers.column_dimensions[col_letter].width = max_length + 5



# Salva os novos arquivos com nome atualizado
for nome, wb in wb_final.items():
    novo_caminho = pasta_scripts.parent / "form4" / f"V2_{nome}_atualizado_form4.xlsx"
    wb.save(novo_caminho)
    print(f"{novo_caminho} gerado com sucesso")

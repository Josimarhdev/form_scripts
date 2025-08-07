import pandas as pd
import re
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment


def processar_planilhas_excel():
    """
    Processa os arquivos Excel de um caminho específico para analisar o engajamento.
    """
    # Caminho para a pasta onde os arquivos Excel estão localizados
    caminho_dos_arquivos = '/home/josimar/Área de Trabalho/pull4/outputs/'
    
    # Nomes dos arquivos
    arquivo_geral = os.path.join(caminho_dos_arquivos, 'grs_atualizado.xlsx')
    arquivo_form4 = os.path.join(caminho_dos_arquivos, 'grs_atualizado_form4.xlsx')

    try:
        # Carregar os arquivos Excel usando o caminho completo
        xls_geral = pd.ExcelFile(arquivo_geral)
        xls_form4 = pd.ExcelFile(arquivo_form4)
    except FileNotFoundError as e:
        print(f"Erro: Arquivo não encontrado.")
        print(f"Verifique se o caminho '{caminho_dos_arquivos}' e os nomes dos arquivos estão corretos.")
        return None

    # Ler as abas de formulários e monitoramento
    try:
        form1 = pd.read_excel(xls_geral, sheet_name='Form 1 - Município')
        form2 = pd.read_excel(xls_geral, sheet_name='Form 2 - UVR')
        form3 = pd.read_excel(xls_geral, sheet_name='Form 3 - Empreendimento')
        monitoramento = pd.read_excel(xls_geral, sheet_name='Monitoramento')
    except ValueError as e:
        print(f"Erro ao ler uma das abas: {e}. Verifique se os nomes das abas estão corretos nos arquivos.")
        return None

    # Limpar e obter a lista de UVRs da aba Monitoramento
    monitoramento.columns = [
        'Regional', 'Municípios', 'UVR', 'Form 1 - Município', 'Form 2 - UVR',
        'Form 3 - Empreendimento', 'Unnamed: 6', 'Regional.1', 'Municípios.1',
        'UVR.1', 'Form 1 - Município.1', 'Form 2 - UVR.1', 'Form 3 - Empreendimento.1'
    ]

    monitoramento.drop(monitoramento[monitoramento['Regional'] == 'Regional'].index, inplace=True)

    uvr_list1 = monitoramento[['Regional', 'Municípios', 'UVR']].dropna(subset=['Regional'])
    uvr_list2 = monitoramento[['Regional.1', 'Municípios.1', 'UVR.1']].dropna(subset=['Regional.1'])
    uvr_list2.columns = ['Regional', 'Municípios', 'UVR']
    all_uvrs = pd.concat([uvr_list1, uvr_list2]).drop_duplicates().reset_index(drop=True)
    all_uvrs.rename(columns={'Municípios': 'Municipio'}, inplace=True)
    
    # Processar Form 1, 2, e 3
    def check_submission(df, municipio, uvr):
        df_filtered = df[(df['Município'] == municipio) & (df['UVR'] == uvr)]
        if not df_filtered.empty:
            situacao = df_filtered['Situação'].iloc[0]
            if situacao in ['Enviado', 'Duplicado']:
                return 1
        return 0

    all_uvrs['Form 1'] = all_uvrs.apply(lambda row: check_submission(form1, row['Municipio'], row['UVR']), axis=1)
    all_uvrs['Form 2'] = all_uvrs.apply(lambda row: check_submission(form2, row['Municipio'], row['UVR']), axis=1)
    all_uvrs['Form 3'] = all_uvrs.apply(lambda row: check_submission(form3, row['Municipio'], row['UVR']), axis=1)
    
    # Processar Form 4
    current_date = datetime.now()
    current_year = current_date.year
    current_month = current_date.month

    # Total de envios esperados
    expected_2024 = 2 # Novembro a Dezembro
    if current_year < 2025:
        expected_2025 = 0
    elif current_year == 2025:
        # A contagem é até o mês anterior ao atual
        expected_2025 = current_month - 1
    else: # Anos posteriores a 2025
        expected_2025 = 12
    
    submissions_2024 = {}
    submissions_2025 = {}

    # Iterar sobre as abas do Form 4 que correspondem ao padrão MM.AA
    for sheet_name in xls_form4.sheet_names:
        if re.match(r'^\d{2}\.\d{2}$', sheet_name):
            try:
                df_month = pd.read_excel(xls_form4, sheet_name=sheet_name)
                year_suffix = sheet_name.split('.')[1]

                if year_suffix == '24' and sheet_name not in ['11.24', '12.24']:
                    continue
                
                # Só processa abas relevantes para os anos de interesse
                if year_suffix == '24':
                    counts_dict = submissions_2024
                elif year_suffix == '25':
                    counts_dict = submissions_2025
                else:
                    continue
                
                for _, row in df_month.iterrows():
                    key = (row['Município'], row['UVR'])
                    if row['Situação'] in ['Enviado', 'Duplicado']:
                        counts_dict[key] = counts_dict.get(key, 0) + 1
            except Exception as e:
                print(f"Aviso: Não foi possível ler ou processar a aba '{sheet_name}'. Erro: {e}")

    all_uvrs['Form 4 2024'] = all_uvrs.apply(lambda row: submissions_2024.get((row['Municipio'], row['UVR']), 0), axis=1)
    all_uvrs['Form 4 2025'] = all_uvrs.apply(lambda row: submissions_2025.get((row['Municipio'], row['UVR']), 0), axis=1)

    # Calcular engajamento
    total_expected = 3 + expected_2024 + expected_2025
    if total_expected > 0:
        all_uvrs['Total Submissions'] = all_uvrs['Form 1'] + all_uvrs['Form 2'] + all_uvrs['Form 3'] + all_uvrs['Form 4 2024'] + all_uvrs['Form 4 2025']
        all_uvrs['Engagement'] = (all_uvrs['Total Submissions'] / total_expected) * 100
    else:
        all_uvrs['Engagement'] = 0

    def get_engagement_level(percentage):
        if percentage > 80:
            return 'Alto'
        elif 50 <= percentage <= 80:
            return 'Médio'
        else:
            return 'Baixo'

    all_uvrs['Engagement Level'] = all_uvrs['Engagement'].apply(get_engagement_level)
    
    # Formatar colunas do Form 4 para o formato "X de Y"
    all_uvrs['Form 4 2024'] = all_uvrs.apply(lambda row: f"{row['Form 4 2024']} de {expected_2024}", axis=1)
    all_uvrs['Form 4 2025'] = all_uvrs.apply(lambda row: f"{row['Form 4 2025']} de {expected_2025}", axis=1)

    return all_uvrs

def criar_planilha_final(df):
    """
    Cria a planilha Excel final com os dados de engajamento e formatação de cores.
    """
    if df is None:
        print("Nenhum dado para processar. A planilha não foi criada.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Engajamento GRS"

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    headers = ['Regional', 'Municipio', 'UVR', 'Form 1', 'Form 2', 'Form 3', 'Form 4 2024', 'Form 4 2025', 'Nível de Engajamento']
    ws.append(headers)
    
    for _, row in df.iterrows():
        ws.append([
            row['Regional'], row['Municipio'], row['UVR'], 
            'Enviado' if row['Form 1'] == 1 else 'Ausente',
            'Enviado' if row['Form 2'] == 1 else 'Ausente',
            'Enviado' if row['Form 3'] == 1 else 'Ausente',
            row['Form 4 2024'], row['Form 4 2025'],
            row['Engagement Level']
        ])
        
        level = row['Engagement Level']
        fill = None
        if level == 'Alto':
            fill = green_fill
        elif level == 'Médio':
            fill = yellow_fill
        else: # Baixo
            fill = red_fill
            
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill
    
    output_filename = "analise_engajamento.xlsx"
    wb.save(output_filename)
    print(f"\nPlanilha '{output_filename}' gerada com sucesso!")
    print(f"O arquivo foi salvo em: {os.path.abspath(output_filename)}")

# --- Execução Principal ---
if __name__ == "__main__":
    df_final = processar_planilhas_excel()
    if df_final is not None:
        criar_planilha_final(df_final[[
            'Regional', 'Municipio', 'UVR', 'Form 1', 'Form 2', 'Form 3', 
            'Form 4 2024', 'Form 4 2025', 'Engagement Level'
        ]])
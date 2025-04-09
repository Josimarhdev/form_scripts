from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from datetime import datetime
import unicodedata

# Estilos reutilizáveis
cabeçalho_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
cabeçalho_font = Font(color="FFFFFF", bold=True, name='Arial', size=11)

enviado_fill = PatternFill(start_color="006400", end_color="006400", fill_type="solid")
enviado_font = Font(color="FFFFFF", name='Arial', size=11)
semtecnico_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
atrasado_fill = PatternFill(start_color="FF6400", end_color="FF6400", fill_type="solid")

validado_nao_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
validado_sim_fill = PatternFill(start_color="66FF66", end_color="66FF66", fill_type="solid")

cores_regionais = {
    "Gabriel": "A9C5E6",
    "Bianca": "FFFF99",
    "Valquiria": "B2FFFF",
    "Luana": "FFCCFF",
    "Larissa": "F1E0C6",
    "Paranavai": "9B59B6",
    "Ana Paula": "993399",
    "Londrina": "A9C5E6",
    "Francisco Beltrão": "B2FFFF",
    "Maringá": "FFCCFF",
    "Curitiba": "FFFF99",
    "Guarapuava": "F1E0C6"
}

bordas = Border(
    top=Side(border_style="thin", color="000000"),
    bottom=Side(border_style="thin", color="000000"),
    left=Side(border_style="thin", color="000000"),
    right=Side(border_style="thin", color="000000")
)

alinhamento = Alignment(horizontal="center", vertical="center")

# Funções reutilizáveis
def corrigir_acentuacao(texto):
    try:
        return texto.encode('latin1').decode('utf-8')
    except (UnicodeDecodeError, AttributeError):
        return texto

def normalizar_texto(texto):
    if isinstance(texto, str):
        texto = texto.strip().lower()
        texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto

def normalizar_uvr(uvr_nro):
    try:
        return str(int(uvr_nro))
    except ValueError:
        return uvr_nro



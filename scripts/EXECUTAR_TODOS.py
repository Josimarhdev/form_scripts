from openpyxl import Workbook
from pathlib import Path 

# Inicializa os workbooks globais
belem_wb = Workbook()
expansao_wb = Workbook()
grs_wb = Workbook()

# Remove aba padrão de cada workbook
for wb in [belem_wb, expansao_wb, grs_wb]:
    wb.remove(wb.active)

# Deixa as variáveis globais disponíveis para os scripts
globals().update({
    "belem_wb": belem_wb,
    "expansao_wb": expansao_wb,
    "grs_wb": grs_wb
})


exec(open("scripts/script_form1.py").read())
exec(open("scripts/script_form2.py").read())
exec(open("scripts/script_form3.py").read())
exec(open("scripts/script_form4.py").read())


pasta_scripts = Path(__file__).parent
pasta_saida = pasta_scripts.parent / "outputs"

# Cria a pasta se não existir
pasta_saida.mkdir(parents=True, exist_ok=True)


belem_wb.save(pasta_saida / "belem_atualizado.xlsx")
expansao_wb.save(pasta_saida / "expansao_atualizado.xlsx")
grs_wb.save(pasta_saida / "grs_atualizado.xlsx")

print(f"Arquivos salvos em: {pasta_saida}")

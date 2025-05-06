# Form Scripts

Este repositório contém uma coleção de formulários e scripts associados à automação da atualização dos mesmos. Os formulários estão organizados em diretórios separados, cada um representando um formulário específico.

## Estrutura do Repositório

- form1/: Contém as planilhas geradas pelo script e arquivos auxiliares (planilhas originais/dataset) do form 1.
- form2/: Contém as planilhas geradas pelo script e arquivos auxiliares (planilhas originais/dataset) do form 2.
- form3/: Contém as planilhas geradas pelo script e arquivos auxiliares (planilhas originais/dataset) do form 3.
- scripts/: Inclui os scripts utilizados para atualização dos formulários.

## Tecnologias Utilizadas

- Python: Linguagem principal utilizada para desenvolver os scripts.
- Openpyxl: Biblioteca utilizada para manipular as planilhas

## Objetivos e Regras de Verificação de Cada Formulário

### Form 1:
- **Objetivo:** Verificar os envios com base no município.
- **Regra de Verificação:** O script busca os dados extraídos do dataset e verifica se os envios foram realizados para cada município, atualizando as planilhas auxiliares (GRS, Expansão e Belém) com as informações pertinentes. O envio é registrado apenas considerando o município.

### Form 2:
- **Objetivo:** Verificar os envios com base no município/UVR.
- **Regra de Verificação:** O script coleta os dados do dataset e verifica os envios considerando tanto o município quanto a UVR associada a ele. Isso permite um controle mais detalhado dos envios, comparando os dados de cada município com a sua UVR correspondente nas planilhas auxiliares.

### Form 3:
- **Objetivo:** Verificar os envios com base no município/UVR/empreendimento.
- **Regra de Verificação:** O script realiza a verificação de envios mais detalhada, agora considerando não apenas o município e a UVR, mas também os empreendimentos relacionados. 

### Form 4:
- **Objetivo:** Verificar os envios com base no município/UVR mensalmente, com abas no formato MM.AA.
- **Regra de Verificação:** O script coleta os dados extraídos do dataset e realiza uma verificação mensal por município e UVR. O Formulário 4 possui abas divididas por mês no formato MM.AA (por exemplo, 01.25 para janeiro de 2025). O script atualiza as planilhas auxiliares mensalmente, conforme as datas de envio e os municípios/UVRs.

### Processo Geral de Atualização:
1. **Entrada de Dados:** O script recebe uma extração do dataset e três planilhas auxiliares (GRS, Expansão e Belém).
2. **Processamento dos Dados:** O script coleta os dados da extração e verifica os envios conforme as regras de cada formulário (acima descritas).
3. **Saída de Dados:** Após processar os dados, o script atualiza as planilhas auxiliares (GRS, Expansão e Belém) e gera novos arquivos com os dados atualizados.

## Como Utilizar

1. Clone este repositório para sua máquina local:
   ```
   git clone https://github.com/Josimarhdev/form_scripts.git
2. Navegue até o diretório desejado:
   ```
   cd form_scripts/scripts
3. Execute o script correspondente (exemplo para o script do form1):
   ```
   python script_form1.py 

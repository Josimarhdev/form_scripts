# Form Scripts

Este repositório contém uma coleção de formulários e scripts associados à automação da atualização dos mesmos. Os formulários estão organizados em diretórios separados, cada um representando um formulário específico.

## Estrutura do Repositório


- `form4/`: Contém as planilhas de saída da versão de testes do formulário 04 (form4v2).
- `inputs/`: Contém os arquivos de entrada dos formulários 1, 2, 3 e 4.
- `outputs/`: Contém os arquivos de saída dos formulários 1, 2, 3 e 4.
- `scripts/`: Inclui os scripts utilizados para atualização dos formulários.


## Tecnologias Utilizadas

- **Python**: Linguagem principal utilizada para desenvolver os scripts.
- **Openpyxl**: Biblioteca utilizada para manipular planilhas Excel (.xlsx).

## Objetivos e Regras de Verificação de Cada Formulário

### Form 1
- **Objetivo**: Verificar os envios com base no município.
- **Regra de Verificação**: O script busca os dados extraídos do dataset e verifica se os envios foram realizados para cada município, atualizando as planilhas auxiliares (GRS, Expansão e Belém) com as informações pertinentes. O envio é registrado apenas considerando o município.

### Form 2
- **Objetivo**: Verificar os envios com base no município/UVR.
- **Regra de Verificação**: O script coleta os dados do dataset e verifica os envios considerando tanto o município quanto a UVR associada a ele. Isso permite um controle mais detalhado dos envios, comparando os dados de cada município com sua UVR correspondente nas planilhas auxiliares.

### Form 3
- **Objetivo**: Verificar os envios com base no município/UVR/empreendimento.
- **Regra de Verificação**: O script realiza a verificação de envios mais detalhada, agora considerando não apenas o município e a UVR, mas também os empreendimentos relacionados.

### Form 4
- **Objetivo**: Verificar os envios com base no município/UVR mensalmente, com abas no formato MM.AA.
- **Regra de Verificação**: O script coleta os dados extraídos do dataset e realiza uma verificação mensal por município e UVR. O Formulário 4 possui abas divididas por mês no formato MM.AA (por exemplo, `01.25` para janeiro de 2025). O script atualiza as abas conforme as datas de envio e os municípios/UVRs.

## Processo Geral de Atualização

1. **Entrada de Dados**: O script recebe uma extração do dataset e três planilhas auxiliares (GRS, Expansão e Belém).
2. **Processamento dos Dados**: O script coleta os dados da extração e verifica os envios conforme as regras de cada formulário (descritas acima).
3. **Saída de Dados**: Após processar os dados, o script atualiza as planilhas auxiliares e gera novos arquivos com os dados atualizados.

## Instalação de Dependências

Antes de executar os scripts, instale as dependências do projeto com o seguinte comando:

```bash
pip install -r requirements.txt
```

## Como Utilizar

1. **Clone este repositório:**

```bash
git clone https://github.com/Josimarhdev/form_scripts.git
```

2. **Instale as dependências:**

```bash
pip install -r requirements.txt
```

3. **Insira os arquivos de entrada atualizados:**

Na pasta principal de `inputs/`, insira os seguintes arquivos e pastas. Esta pasta centralizará todas as fontes de dados necessárias para a execução.

- As pastas com as planilhas de consumo originais:
    - `0 - Belém`
    - `0 - Expansão`
    - `0 - GRS II`
- Os arquivos de extração do banco de dados no formato `.csv`:
    - `form1.csv`
    - `form2.csv`
    - `form3.csv`
    - `form4.csv`

A estrutura de pastas e arquivos esperada dentro de `inputs/` deve ser a seguinte:
```
inputs/
├── 0 - Belém/
│   └── ...
├── 0 - Expansão/
│   └── ... 
├── 0 - GRS II/
│   └── ... 
├── form1.csv
├── form2.csv
├── form3.csv
└── form4.csv
```

4. **Acesse o diretório dos scripts:**

```bash
cd form_scripts/scripts
```

5. **Execute o script geral:**

Exemplo:

```bash
python EXECUTAR_TODOS.py
```

O script irá processar os dados e gerar novas planilhas atualizadas na pasta de outputs.
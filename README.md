# Unificador de Colunas de Planilhas

Este projeto é uma ferramenta simples com interface gráfica (GUI) para unificar colunas em planilhas (Excel ou CSV). Ele foi projetado para pegar colunas com nomes diferentes que representam a mesma informação (como "Endereço" e "Endereço Entrega") e consolidá-las em uma única coluna padronizada.

## Como Funciona

O script lê um arquivo de planilha e, para um conjunto pré-definido de colunas, realiza as seguintes ações:

1.  **Mapeamento de Colunas:** Existe um mapeamento que define quais colunas devem ser unificadas. Por exemplo:
    - `Endereço Entrega` e `Endereço` são unificados em `Endereço Unificado`.
    - `Número Entrega`, `Número Número Entrega` e `Número` são unificados em `Número Unificado`.
    - E assim por diante para CEP e Bairro.

2.  **Lógica de Preenchimento:**
    - O script cria uma nova coluna (ex: `Endereço Unificado`).
    - Ele primeiro tenta preencher essa nova coluna com os dados da coluna "primária" (ex: `Endereço Entrega`).
    - Se o valor na coluna primária estiver vazio para uma determinada linha, ele busca em colunas "fallback" (ex: `Endereço`) e usa o primeiro valor que encontrar para preencher a lacuna.

3.  **Salvar o Resultado:** Após processar todas as colunas mapeadas, o script gera um **novo arquivo Excel (.xlsx)** com as colunas unificadas adicionadas, sem modificar o arquivo original.

## Como Usar

1.  **Execute o programa:** Rode o arquivo `unificadorcolunas.py`. Uma janela da aplicação será aberta.
2.  **Selecione o arquivo:** Clique no botão "Selecionar Arquivo..." e escolha a planilha (.xlsx ou .csv) que você deseja processar.
3.  **Inicie o processamento:** Clique no botão "Processar e Salvar".
4.  **Salve o novo arquivo:** Uma janela de diálogo aparecerá para você escolher o nome e o local onde o novo arquivo processado será salvo. Por padrão, o nome sugerido é `planilha_com_colunas_unificadas.xlsx`.

## Dependências

Para executar este script, você precisa ter as seguintes bibliotecas Python instaladas:

-   **pandas:** Para manipulação e leitura dos dados da planilha.
-   **openpyxl:** Necessário para o pandas ler e escrever arquivos no formato `.xlsx`.

Você pode instalá-las usando o pip:

```bash
pip install pandas openpyxl
```

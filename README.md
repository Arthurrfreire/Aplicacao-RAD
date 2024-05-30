## RAD - Registro Acadêmico e Desempenho

Este software fornece uma interface gráfica para registrar, calcular e salvar as médias dos alunos, além de indicar sua situação acadêmica (Aprovado ou Reprovado).

## Funcionalidades Principais

**Registro de Alunos:** Permite inserir o nome do aluno e suas duas notas.

**Cálculo de Média:** Calcula automaticamente a média das notas inseridas.

**Verificação de Situação:** Avalia se o aluno está Aprovado (média >= 7) ou Reprovado (média < 7).

**Visualização em Tabela:** Exibe os dados dos alunos em uma tabela organizada, incluindo nome, notas, média e situação.

**Carregamento de Dados:** Carrega dados de um arquivo Excel existente (planilhaAlunos.xlsx).

**Salvamento de Dados:** Salva os dados em um arquivo Excel (planilhaAlunos.xlsx).

## Como Usar

**Insira os Dados:** Digite o nome do aluno e suas duas notas nos campos correspondentes.

**Calcule a Média:** Clique no botão "Calcular Média". A média e a situação do aluno serão calculadas e exibidas na tabela.

**Visualize os Dados:** Os dados dos alunos serão listados na tabela, permitindo fácil acompanhamento do desempenho da turma.

**Salve os Dados:** Os dados são automaticamente salvos no arquivo Excel "planilhaAlunos.xlsx".

## Pré-requisitos

- Python 3.x
- Bibliotecas:

  - tkinter

  - pandas

## Instalação

1. Certifique-se de ter o Python instalado em seu sistema.

2. Instale as bibliotecas necessárias:


        pip install tkinter pandas

## Execução

1. Salve o código fornecido em um arquivo Python (por exemplo, rad.py).

2. Execute o arquivo:

       python rad.py

## Observações

- O software utiliza um arquivo Excel ("planilhaAlunos.xlsx") para armazenar os dados. Certifique-se de que este arquivo exista no mesmo diretório do script Python.

- Caso deseje utilizar um arquivo CSV, descomente as linhas relevantes no código e comente as linhas referentes ao Excel.

**Importante:** Este software é um exemplo básico e pode ser expandido para incluir mais funcionalidades, como edição de dados, filtros de pesquisa, geração de relatórios, etc.

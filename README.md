# Somatório de Vendas

Projeto de aprendizagem em **Python e openpyxl** para automatizar o registro de vendas e a atualização de fórmulas em planilhas Excel.

**Situação: em desenvolvimento.** A versão atual altera a planilha em memória; o salvamento em arquivo ainda não está ativo.

## Objetivo

Aplicar programação a uma necessidade do varejo: organizar lançamentos de vendas por vendedor e manter as fórmulas de soma atualizadas. O projeto também é um ambiente de estudo de funções, dicionários, validação de entradas, tratamento de exceções e versionamento com Git/GitHub.

## O que existe no código atual

- Identificação de vendedores a partir dos cabeçalhos da planilha.
- Conversão básica de valores numéricos, inclusive texto com vírgula decimal.
- Rejeição de entradas não numéricas, zero e valores negativos.
- Registro da primeira venda e criação de uma fórmula de soma.
- Atualização da fórmula ao acrescentar uma venda em uma coluna.
- Verificação da célula abaixo do somatório antes de ocupá-la no fluxo de nova venda.

A função principal é `registrar_venda(aba, coluna, nova_venda)`. Ela trabalha com uma coluna por chamada.

## Tecnologias

- Python
- openpyxl
- Excel como formato de entrada e saída pretendida
- Git e GitHub

As dependências estão declaradas em `requirements.txt`.

## Preparar o ambiente

Em uma pasta local do projeto, com Python instalado:

~~~powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
~~~

O programa depende de um arquivo Excel e de uma aba configurados no final de `somatorio.py`. Prepare uma planilha fictícia e ajuste essas referências antes de experimentar.

## Como interpretar a versão atual

O arquivo tem um fluxo experimental no nível do módulo: executar ou importar `somatorio.py` abre a planilha configurada e tenta realizar um lançamento. Ainda não existe menu ou interface para selecionar arquivo, vendedor e valor.

A chamada de salvamento está em um bloco inativo. Um retorno positivo da função indica uma alteração em memória, e não a confirmação de que o arquivo foi salvo. As fórmulas são escritas na planilha; o código não verifica o total recalculado pelo Excel.

Os exemplos no fim do arquivo não constituem uma suíte de testes automatizados.

## Próximos passos

- Separar a execução experimental da definição das funções.
- Configurar arquivo, aba e entradas sem editar o código a cada uso.
- Implementar e verificar o salvamento em uma cópia de teste.
- Ampliar as validações de valores monetários e da estrutura da planilha.
- Tratar lançamentos relacionados em duas colunas como uma única operação, sem duplicar a receita.
- Adicionar testes automatizados com dados fictícios.

Esses itens são objetivos de evolução; não estão apresentados como funcionalidades prontas.

## Dados de demonstração

Use planilhas fictícias para estudo e demonstrações. Planilhas reais da empresa, dados de clientes, credenciais e backups da operação não devem ser incluídos no repositório.

## Autor

**Gabriel Victor**, estudante de Engenharia de Software.

[LinkedIn](https://www.linkedin.com/in/gabriel-victor-1a041b3b9/) · [GitHub](https://github.com/gv8407940-arch)

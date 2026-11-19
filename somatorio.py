from openpyxl import load_workbook

def registrar_venda(aba, coluna, nova_venda):


 #------------1º Etapa: Verificação de validade do valor: se string ou número -------------#
    try:
        if isinstance(nova_venda, str):
            nova_venda = nova_venda.strip().replace(",", ".")

        nova_venda = float(nova_venda)

    except(ValueError,TypeError):
        print("Erro: informe um valor numérico válido")
        return False

 #-----------2º Etapa: Verificação de validade do valor: se positivo ou falso --------------#
    if nova_venda <= 0:
        print("Erro: insira um valor positivo e maior que 0")
        return False


    tem_vendas = False
    tem_somatorio = False
    for celula in aba[coluna]:

        if celula.row >= 2:
            if celula.value is not None and celula.data_type == "n":
                tem_vendas = True

            if celula.value is not None and celula.data_type == "f":
                tem_somatorio = True

 #----------3º Etapa: Alteração da célula com a fórmula --------------------#

    if tem_vendas == False and tem_somatorio == False:
        primeira_venda = 2
        linha_somatorio = 3

        #Formula de Soma
        somatorio = f"=SUM({coluna}2:{coluna}{primeira_venda})"

        #Definicao de local para primeira venda
        aba[f"{coluna}{primeira_venda}"] = nova_venda

        #Definicao de local para somatório
        aba[f"{coluna}{linha_somatorio}"] = somatorio

        return True

    for celula in aba[coluna]:
        if celula.data_type == "f":
            linha_somatorio = celula.row
            ultima_venda = celula.row - 1
            nova_linha_somatorio = linha_somatorio + 1

            print("Formula Encontrada:", celula.coordinate, celula.value)
            print("Ultima venda:", ultima_venda)

            #Formula de Soma
            somatorio = f"=SUM({coluna}2:{coluna}{linha_somatorio})"

            if aba[f"{coluna}{nova_linha_somatorio}"].value is not None:
                print("Erro: próxima célula já está ocupada!")
                return False

            #Substituir o somatório antigo pela nova venda
            aba[f"{coluna}{linha_somatorio}"] = nova_venda

            #Criar somatório na próxima linha
            aba[f"{coluna}{nova_linha_somatorio}"] = somatorio

            return True

    print("Erro: nenhum somatório encontrado nesta coluna!")
    return False


planilha = load_workbook("Fevereiro25_teste.xlsx")
aba = planilha["Out 26"]

colunas = {}

for celula in aba[1]:
    if celula.value is not None:
        colunas[celula.value.lower().strip()] = celula.column_letter


vendedor = 'edna'
vendedor_padronizado = vendedor.lower().strip()

if vendedor_padronizado in colunas:
    coluna_vendedor = colunas[vendedor_padronizado]
    registrar_venda(aba, coluna_vendedor, 500)
    print("Vendedor Encontrado!")
    print(f"Coluna {coluna_vendedor}")
else:
    print("Erro! Vendedor não encontrado")


'''
resultado = registrar_venda(aba, "A", 1299.90)

if resultado:
    planilha.save("Fevereiro25_teste.xlsx")
    print("Registro Realizado com sucesso")
else:
    print("Não foi Possível realizar o registro")


valores_teste = ["500", "1299,90", "-200", "abc", 0]

for valor in valores_teste:

    resultado = registrar_venda(aba, "A", valor)

    print(f"Valor: {valor} | Resultado: {resultado}")
    print("-------------------")
    '''

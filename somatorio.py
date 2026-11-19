from openpyxl import load_workbook

def registrar_venda(aba, coluna, nova_venda):
    
    if not isinstance(nova_venda, (int, float)):
        print("Erro: o valor da venda deve ser numérico")
        return False

    if nova_venda <= 0:
        print("Erro: insira um valor positivo")
        return False
        

    for celula in aba[coluna]: 
        if celula.data_type == "f":
            linha_somatorio = celula.row            
            ultima_venda = celula.row - 1        
            nova_linha_somatorio = linha_somatorio + 1

            print("Formula Encontrada:", celula.coordinate, celula.value)
            print("Ultima venda:", ultima_venda)
            
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

resultado = registrar_venda(aba, "A", 1299.90)


if resultado:
    planilha.save("Fevereiro25_teste.xlsx")
    print("Registro Realizado com sucesso")
else:
    print("Não foi Possível realizar o registro")



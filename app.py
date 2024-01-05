import openpyxl
import pyperclip
import pyautogui

from time import sleep

#Entra na Planilha produtos_ficticios.xlsx
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook ['Produtos']

#Copia a informação de um campo e cola no campo correspondente no site desktop
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(55,202,duration=1)
    pyautogui.hotkey('ctrl','v')

        #ctrl + k + c = comenta todo código selecionado
    
        #ctrl + k + u = descomenta todo código selecionado

    #Copia a informação de um campo e cola no campo descrição no site desktop 
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(54,284,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo categoria no site desktop
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(114,438,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo codigo_produto no site desktop
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(69,529,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo peso no site desktop
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(57,611,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo dimensões no site desktop
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(68,684,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Clica em próximo
    pyautogui.click(70,751,duration=1)   
    sleep(3)

    #Copia a informação de um campo e cola no campo preço no site desktop
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(51,234,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo quantidade em estoque no site desktop
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(52,321,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo data de validade no site desktop
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(61,398,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo cor no site desktop
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(56,490,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo tamanho no site desktop
    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    pyautogui.click(58,579,duration=1)
    if tamanho == 'Pequeno':     #   Lê a Info da Planilha
        pyautogui.click(119,617,duration=1)     #   Se for "Pequeno", clicar em uma posição
    elif tamanho == 'Médio':  #   Se for "Médio", clicar em uma posição
        pyautogui.click(109,638,duration=1)
    else:
        pyautogui.click(113,669,duration=1)     #   Se for "Grande", clicar em uma posição

    pyautogui.click(16,545,duration=1)

    # pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo material no site desktop
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(84,653,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Clica em próximo
    pyautogui.click(73,713,duration=1)   
    sleep(3)

    #Copia a informação de um campo e cola no campo fabricante no site desktop
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(67,251,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo país de origem no site desktop
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(61,328,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo observações no site desktop
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(55,413,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo código de barras no site desktop
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(58,565,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Copia a informação de um campo e cola no campo localização armazém no site desktop
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(55,653,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Clica em concluir
    pyautogui.click(71,717,duration=1)   
    sleep(3)


    pyautogui.click(467,451,duration=1)   
    sleep(3)
    pyautogui.click(305,471,duration=1)   


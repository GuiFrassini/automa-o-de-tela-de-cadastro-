import openpyxl
import pyperclip
import pyautogui
from time import sleep

#Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1293,360,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1109,443,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1216,580,duration=1)
    pyautogui.hotkey('ctrl','v')    

    codig_produto = linha[3].value
    pyperclip.copy(codig_produto)
    pyautogui.click(1186,669,duration=1)
    pyautogui.hotkey('ctrl','v') 

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1198,760,duration=1)
    pyautogui.hotkey('ctrl','v') 

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1204,849,duration=1)
    pyautogui.hotkey('ctrl','v') 

    pyautogui.click(1143,899,duration=1)
    sleep(2)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1171,388,duration=1)
    pyautogui.hotkey('ctrl','v') 

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1176,475,duration=1)
    pyautogui.hotkey('ctrl','v') 

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1187,560,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1212,648,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1205,730,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1173,759,duration=1)
        
    elif tamanho == 'Médio':
        pyautogui.click(1159,779,duration=1)
    else:
        pyautogui.click(1168,801,duration=1)

    material = linha [11].value
    pyperclip.copy(material)
    pyautogui.click(1198,814,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #botao proximo
    pyautogui.click(1142,879,duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1226,405,duration=1)
    pyautogui.hotkey('ctrl','v')    

    pais_origem =linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1229,495,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1230,590,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1224,716,duration=1)
    pyautogui.hotkey('ctrl','v')    

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1237,792,duration=1)
    pyautogui.hotkey('ctrl','v')

    #botão concluir
    pyautogui.click(1131,849,duration=1)

    #botão confirmar inclusão
    pyautogui.click(1618,163,duration=1)

    #Adicionar mais um
    pyautogui.click(1465,616,duration=1)

    
    
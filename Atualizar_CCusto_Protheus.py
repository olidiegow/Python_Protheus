import pyautogui
import time
import openpyxl

# Pesquisar o protheus no windows
pyautogui.press('win')
time.sleep(5)
pyautogui.typewrite('Protheus')
time.sleep(5)
pyautogui.press('enter')
time.sleep(5)

# Logar no sistema
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(5)

pyautogui.typewrite('DIEGO.OLIVEIRA')
pyautogui.press('tab')
pyautogui.typewrite('1220244c')
pyautogui.press('enter')

# Definir filial e ambiente
pyautogui.moveTo(x=1086, y=776)
pyautogui.click()
time.sleep(5)

# Entrar no sistema
pyautogui.moveTo(x=1093, y=806)
pyautogui.click()
time.sleep(10)

# Atualizações
pyautogui.moveTo(x=72, y=551)
pyautogui.click()
time.sleep(2)

# Cadastros
pyautogui.moveTo(x=75, y=577)
pyautogui.click()
time.sleep(10)

# rolar para baixo
pyautogui.tripleClick(x=101, y=1010)
pyautogui.tripleClick(x=101, y=1010)
pyautogui.tripleClick(x=101, y=1010)
pyautogui.tripleClick(x=101, y=1010)
time.sleep(10)

# moviment.carro
pyautogui.moveTo(x=105, y=904)
print(pyautogui.position())
pyautogui.click()
time.sleep(2)

# incluir movimentação
pyautogui.moveTo(x=78, y=189)
print(pyautogui.position())
pyautogui.click()
time.sleep(2)

planilha = openpyxl.load_workbook('')
aba = planilha.active()

for coluna in aba.iter_rows(min_row=2, values_only=True):
    placa = coluna[0]
    centro_custo = coluna[1]

    # incluir dados
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(placa)
    pyautogui.press('tab')
    pyautogui.typewrite(centro_custo)
    pyautogui.click(x=1699, y=158)
    pyautogui.press('enter')
    print(placa, 'movimentada para: ', centro_custo, '.')
    time.sleep(1)






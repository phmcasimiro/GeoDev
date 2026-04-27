
#----------------------------
# DESAFIO: Achar carros no OLX
#----------------------------

import pyautogui
import time

pyautogui.PAUSE = 1

# Pressionar a tecla do Windows
pyautogui.press('win')

# Digitar Edge
pyautogui.write('edge')

# Apertar enter para abrir o Edge
pyautogui.press('enter')
time.sleep(2) # Aguardar 2 segundos

# definição de endereço em uma variável link
link_sspdf = 'https://www.ssp.df.gov.br/dados-por-regiao-administrativa/#DF'

# Digitar o site e apertar enter
pyautogui.write(link_sspdf)
pyautogui.press('enter')
time.sleep(2) # Aguardar 3 segundos


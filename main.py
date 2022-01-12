from __future__ import print_function
import time
import acessos
import sys
from pywinauto import application
from acessos import *
import yagmail
from datetime import datetime
import pandas as pd
import gspread
import pyautogui as py



# INICIO

now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print("HORA DE INÍCIO =", current_time)

# MENSAGENS AO USER
print("Iniciando a rotina de atualização dos indicadores por automação...")
time.sleep(2)
print("Carregando a rotina...")
time.sleep(2)
print("Para esta interação é importante não mexer no computador durante o processo, para evitar erros.")
time.sleep(5)
print("Limpando DESKTOP")
py.click(1439, 889)
print("Acessando MaxGPS")
time.sleep(2)

# ABRE O MAX GPS
app = application.Application()
app.start("L:\MaxGPS\MaxGPS.exe")
time.sleep(10)

for restante in range(20, 0, -1):
    sys.stdout.write("\r")
    sys.stdout.write("{:2d} segundos restantes.".format(restante))
    sys.stdout.flush()
    time.sleep(1)
sys.stdout.write("\rINSERINDO LOGIN E SENHA!            \n")

py.click(839, 372)
py.typewrite(acessos.nome)
py.typewrite(['tab'])
py.click(839, 407)
py.typewrite(acessos.senha)
py.typewrite(['enter'])

time.sleep(1)
print("ACESSO AO SISTEMA CONCEDIDO")
time.sleep(1)
print("AGUARDE")

for restante in range(10, 0, -1):
    sys.stdout.write("\r")
    sys.stdout.write("{:2d} segundos restantes.".format(restante))
    sys.stdout.flush()
    time.sleep(1)

# --------------------------------------------------------------#
# RELATÓRIO DE VENDAS

sys.stdout.write("\rPREPARANDO RELATÓRIO DE VENDAS!            \n")
py.click(65,108)
py.sleep(5)

# PASSA AS ABAS DE ACORDO COM O MÊS ATUAL:
for contagem in range(mes_atual, 0, -1):
    py.click(45, 186)

py.sleep(3)
py.click(447,82)

print("GERANDO RELATÓRIO")

# RELATORIO DE CONTROLE DE ENTREGAS E F.E

py.click(57,182)
py.sleep(2)
py.click(378,41)
py.sleep(2)
py.click(556,90)
py.sleep(10)
py.click(363,9)
py.sleep(2)
py.hotkey('down')
py.sleep(2)
py.hotkey('enter')
py.sleep(2)

# TEMPO DE ESPERA DO RELATÓRIO
print("GERANDO RELATÓRIO: POR FAVOR AGUARDE!")

for restante in range(300, 0, -1):
    sys.stdout.write("\r")
    sys.stdout.write("{:2d} segundos restantes.".format(restante))
    sys.stdout.flush()
    time.sleep(1)

# SALVANDO RELATÓRIO DE FATURAMENTO
print("ARQUIVO SENDO SALVO - AGUARDE!")
py.hotkey('ctrl','s')
py.sleep(5)
py.typewrite("C:\\Users\Micro\\Desktop\\dados de faturamento - 2021 - REL. VENDAS-EXP-CONTROLE DE ENTREGAS.xlsx")
py.sleep(2)
py.hotkey('enter')
py.sleep(2)
py.hotkey('left')
py.sleep(2)
py.hotkey('enter')
py.sleep(5)

# SALVANDO RELATÓRIO DE F.E
print("INICIANDO RELATÓRIO F.E - AGUARDE")

py.click(736,91)

for restante in range(250, 0, -1):
    sys.stdout.write("\r")
    sys.stdout.write("{:2d} segundos restantes.".format(restante))
    sys.stdout.flush()
    time.sleep(1)

py.click(360,8)
py.sleep(2)
py.hotkey('down')
py.sleep(2)
py.hotkey('enter')
py.sleep(2)

# SALVANDO F.E
print("ARQUIVO SENDO SALVO - AGUARDE!")
py.hotkey('ctrl','s')
py.sleep(5)
py.typewrite("C:\\Users\Micro\\Desktop\\faturamento por FE.xlsx")
py.sleep(2)
py.hotkey('enter')
py.sleep(2)
py.hotkey('left')
py.sleep(2)
py.hotkey('enter')
py.sleep(5)

# FECHA O MAXGPS

py.hotkey('Alt','F4')


print("INICIANDO O GSPREAD PARA O GOOGLE DRIVE - VOCÊ JÁ PODE MEXER NO COMPUTADOR NORMALMENTE")


# INICIA GSPREAD

#LOCALIZAÇÃO DA PLANILHA

gc = gspread.service_account()
sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1KMV-oR6MbUv5fdswz6VskfqpvFdsFLJtUZMMZfpkc9Q")
# PLANILHA DO DRIVE
worksheet = sh.worksheet("FAT")

# DATA FRAME FAT
doc_fat = pd.read_excel("C:\\Users\Micro\\Desktop\\dados de faturamento - 2021 - REL. VENDAS-EXP-CONTROLE DE ENTREGAS.xlsx")
df = pd.DataFrame(doc_fat)
df['Data Entrega'] = df['Data Entrega'].astype(str)
df = df.fillna("0")
df.to_excel(r'C:\\Users\\Micro\\PycharmProjects\\interfacebotnf\\faturamento.xlsx', index= False)

# ATUALIZA O DRIVE
new_ds = worksheet.update([df.columns.values.tolist()] + df.values.tolist())




# DATA FRAME FE
doc_fat_dois = pd.read_excel("C:\\Users\Micro\\Desktop\\faturamento por FE.xlsx")
df_dois = pd.DataFrame(doc_fat_dois)
df_dois['Data'] = df_dois['Data'].astype(str)
df_dois = df_dois.fillna("0")
df_dois.to_excel(r'C:\\Users\\Micro\\PycharmProjects\\interfacebotnf\\faturamento por FE.xlsx', index= False)

# ATUALIZA O DRIVE

worksheet = sh.worksheet("FE")

new_ds_dois = worksheet.update([df_dois.columns.values.tolist()] + df_dois.values.tolist())
print("Relatório concluído com SUCESSO!")




# EMAIL COM AVISO DA ATUALIZAÇÃO

yagmail.register(username, password)
yag = yagmail.SMTP(username)

yag.send(to=['si@rhowert.com.br'],
         subject="RELATÓRIO DE FATURAMENTO - DIÁRIO",
         contents="Bom dia! Os dados  de faturamento foram atualizados de maneira automática com sucesso" + " as: " + current_time + "\n Acesse o Link para "
                                                                                               "vizualização:" + "\n "
                                                                                                                 "https://datastudio.google.com/reporting/75597fee-6497-4a0b-b4ba-4e836d858dcf/page/p_quaxvwnvoc")
# FIM

now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print("HORA DE INÍCIO =", current_time)

import sys
from botcity.core import DesktopBot
import pyautogui as py
import time
import pandas as pd
import gspread
from datetime import datetime
from prophet import Prophet
import webbrowser

#ATENÇÃO: DEPENDÊNCIAS QUE DEVEM ESTAR NA PASTA GERADA PELO PYINSTALLER: RESOURCES, CRED.JSON E PROPHET

#ATUALIZA O WEBDRIVER
# from selenium import webdriver
# from webdriver_manager.chrome import ChromeDriverManager
# driver = webdriver.Chrome(ChromeDriverManager(path = r"C:\webdriver").install())

now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print("HORA DE INÍCIO =", current_time)
print('---------------------------------------')

mes = int(now.strftime("%m")) #STR DO MÊS CONVERTIDO EM INTERGER
ano = int(now.strftime("%Y"))#STR DO ANO CONVERTIDO EM INTERGER

mes_atual = mes-1 #DEFINE A QUANTIDADE DE CLIQUES NA ABA DE MESES (MES ATUAL - JANEIRO DO ANO ATUAL)
ano_atual = ano + (2-ano) #DEFINE A QUANTIDADE DE CLIQUES NA ABA DE ANOS 202X + 1 - 202X)



class Bot(DesktopBot):
    def action(self, execution=None):

        try:
            # Instantiate a DesktopBot
            desktop_bot = DesktopBot()
            print('ROTINA DE EXTRAÇÃO DE DADOS AUTOMÁTICOS NO MAXGPS')
            print('------------------------------------------')

            # ATIVE QUANDO UTILIZAR UM EXE
            #print('Fechando possível janelas anteriores')
            #py.hotkey('windows','down')

            #self.wait(3000)

            #py.hotkey('alt','f4')


            desktop_bot.execute(
                r'\\****\MaxGPS.exe')

            self.wait(20000)
            print('Aguardando a resposta do sistema MAX GPS')
            print('------------------------------------------')


            if not self.find( "botao_inical", matching=0.97, waiting_time=10000):
                self.not_found("botao_inical")


            self.paste('USER')
            self.wait(5000)
            self.tab()
            self.paste('PASSWORD')
            self.enter()
            self.wait(10000)

            print('Iniciando o relatório de faturamento')
            print('------------------------------------------')

            print('Ajustando o range do período')
            print('------------------------------------------')



            if not self.find( "periodo", matching=0.97, waiting_time=10000):
                self.not_found("periodo")
            self.click()


            for contagem in range(mes_atual, 0, -1):
               py.click(40, 198)

            self.wait(5000)

            for i in range(ano_atual, 0, -1):
                py.click(138, 200)


            if not self.find( "faturamento", matching=0.97, waiting_time=10000):
                self.not_found("faturamento")
            self.click()


            self.wait(5000)


            if not self.find( "emissao_nf", matching=0.97, waiting_time=10000):
                self.not_found("emissao_nf")
            self.click()

            while self.find( "carregando_nf", matching=0.97, waiting_time=10000):
                self.wait(5000)
                print('carregando...')

            while self.find( "carregando_dados", matching=0.97, waiting_time=10000):
                self.wait(5000)
                print('carregando...')

            while self.find( "aguarde_atualizando", matching=0.97, waiting_time=10000):
                self.wait(5000)
                print('carregando...')

            print('Extraindo o relatório de notas emitidas')
            print('---------------------------------------')



            if not self.find( "extrair", matching=0.97, waiting_time=10000):
                self.not_found("extrair")
            self.click()
            self.wait(5000)
            self.type_down()
            self.enter()
            self.wait(30000)
            print('Salvando o relatório...')
            print('-----------------------')



            if not self.find( "salvar_como", matching=0.97, waiting_time=10000):
                self.not_found("salvar_como")
            self.click()
            time.sleep(1)
            self.paste(r'L:\Relatórios\faturamento.xlsx')
            time.sleep(1)
            self.enter()



            if not self.find( "alerta", matching=0.97, waiting_time=10000):
                self.not_found("alerta")
                print('ATENÇÃO: primeira gravação do arquivo')
                print('-----------------------')
                self.enter()
                self.wait(5000)

            else:
                print('ATENÇÃO: Este arquivo já existe e será substituído')
                print('-----------------------')
                self.type_left()
                self.enter()
                self.wait(5000)

            print('Buscando os dados de produção para o mesmo período...')
            print('-----------------------')


            if not self.find( "producao", matching=0.97, waiting_time=10000):
                self.not_found("producao")
            self.click()

            self.wait(5000)


            if not self.find( "pcp", matching=0.97, waiting_time=10000):
                self.not_found("pcp")
            self.click()

            self.wait(5000)

            print('Gerando os dados de produção...')
            print('-----------------------')


            if not self.find( "estatisticas", matching=0.97, waiting_time=10000):
                self.not_found("estatisticas")
            self.click()

            self.wait(5000)



            if not self.find( "gerador", matching=0.97, waiting_time=10000):
                self.not_found("gerador")
            self.click()
            self.type_down()
            self.enter()

            print('Aguardando a geração do relatório...')
            print('-----------------------')

            while self.find( "carregando_dad_ag", matching=0.97, waiting_time=10000):
                self.wait(5)
                print('carregando...')

            print('Transformando em excel...')
            print('-----------------------')

            if not self.find( "extrair", matching=0.97, waiting_time=10000):
                self.not_found("extrair")
            self.click()
            self.type_down()
            self.enter()

            while self.find( "exportando", matching=0.97, waiting_time=10000):
                self.wait(50000)
                print('carregando...')



            print('Salvando o arquivo...')
            print('-----------------------')

            if not self.find( "salvar_como", matching=0.97, waiting_time=10000):
                self.not_found("salvar_como")
            self.click()
            self.paste(r'L:\Relatórios\pcp.xlsx')
            self.enter()

            if not self.find( "alerta", matching=0.97, waiting_time=10000):
                self.not_found("alerta")
                print('ATENÇÃO: primeira gravação do arquivo')
                print('-----------------------')
                self.enter()
                self.wait(5000)

            else:
                print('ATENÇÃO: Este arquivo já existe e será substituído')
                print('-----------------------')
                self.type_left()
                self.enter()
                self.wait(5000)

            print('Saindo do sistema ERP...')
            print('-----------------------')

            print('Iniciando os dados do indicador de Tarefas')
            
            if not self.find( "qualidade_lateral", matching=0.97, waiting_time=10000):
                self.not_found("qualidade_lateral")
            self.click()
            
            if not self.find( "sqg", matching=0.97, waiting_time=10000):
                self.not_found("sqg")
            self.click()
            
            if not self.find( "tarefas", matching=0.97, waiting_time=10000):
                self.not_found("tarefas")
            self.click()
            
            if not self.find( "extrair", matching=0.97, waiting_time=10000):
                self.not_found("extrair")
            self.click()
            self.wait(5000)
            self.type_down()
            self.enter()
            self.wait(30000)
            print('Salvando o relatório...')
            print('-----------------------')
            
            if not self.find( "salvar_como", matching=0.97, waiting_time=10000):
                self.not_found("salvar_como")
            self.click()
            time.sleep(1)
            self.paste(r'L:\Relatórios\tarefas.xlsx')
            time.sleep(1)
            self.enter()
            print('SALVO COM SUCESSO!')
     
            
            
            


            if not self.find( "encerrar", matching=0.97, waiting_time=10000):
                self.not_found("encerrar")
            self.click()
            print('Tarefa no ERP finalizada com SUCESSO!' + '\n' + 'Iniciando FORECAST e alimentação do BI')

        except Exception:
            print('Algo deu errado, vou reiniciar...')
            Bot.main()

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()

    service_json = "service_account.json"

    # LOCALIZAÇÃO DA PLANILHA
    gs = gspread.service_account(service_json)
    sh = gs.open_by_url("https://docs.google.com/spreadsheets/d/*******")
    # PLANILHA DO DRIVE
    worksheet = sh.worksheet("estatisticas de produção1")
    val = worksheet.range('A:X')

    mes = int(now.strftime("%m"))  # STR DO MÊS CONVERTIDO EM INTERGER
    ano = int(now.strftime("%Y"))  # STR DO ANO CONVERTIDO EM INTERGER
    mes_atual = mes - 1  # DEFINE A QUANTIDADE DE CLIQUES NA ABA DE MESES (MES ATUAL - JANEIRO DO ANO ATUAL)
    ano_atual = ano + (1 - ano)  # DEFINE A QUANTIDADE DE CLIQUES NA ABA DE ANOS 202X + 1 - 202X)

    # ARQUIVO ESTATISTICAS MAX GPS AS PCP.XLSX
    nome_arquivo = pd.read_excel(r'L:\Relatórios\pcp.xlsx', engine='openpyxl')
    data_frame = pd.DataFrame(nome_arquivo)
    data_frame.columns = (data_frame.iloc[0])
    data_frame.drop(index=data_frame.index[-1], axis=0, inplace=True)
    data_frame.drop(index=data_frame.index[0], axis=0, inplace=True)
    data_frame = data_frame.fillna('')

    data_frame["Início"] = data_frame["Início"].astype(str)
    data_frame["Término"] = data_frame["Término"].astype(str)

    data_final = data_frame["Término"].values
    lista = []

    for i in data_final:
        valor = i.split('.', 1)[0]

        lista.append(valor)

    lista_df = pd.DataFrame(lista, columns=['Término'])
    lista_df.drop(index=lista_df.index[-1], axis=0, inplace=True)
    worksheet.batch_clear(['A:X'])
    time.sleep(3)

    print("Alimentando o DRIVE")
    new_ds = worksheet.update([data_frame.columns.values.tolist()] + data_frame.values.tolist())
    time.sleep(5)
    worksheet = sh.worksheet("estatisticas de produção1")
    worksheet.batch_clear(['E:E'])
    time.sleep(1)
    val_dois = worksheet.range('E:E')
    worksheet.update('E:E', [lista_df.columns.values.tolist()] + lista_df.values.tolist())

    print('Iniciando indicador de faturamento')

    faturamento = pd.read_excel(r'L:\Relatórios\faturamento.xlsx', engine='openpyxl')
    fat_df = pd.DataFrame(faturamento)
    colunas = ['Sel', 'Finalidade', 'NF', 'CFOP', 'Tipo', 'Código', 'Cliente / Fornecedor', 'Emissão', 'Sefaz',
               'Valor NF',
               'Situação NFe', 'Fatura', 'Chave', 'Protocolo']

    df_extract = fat_df
    fat_df = fat_df[fat_df['Fatura'] > 0]

    fat_df = fat_df.astype(str)
    fat_df = fat_df[fat_df['Situação NFe'] == 'Sucesso']

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print("HORA DE INÍCIO =", current_time)
    print('Iniciando a alimentação dos arquivos de PCP no Google Drive')

    service_json = "service_account.json"

    # LOCALIZAÇÃO DA PLANILHA
    gs = gspread.service_account(service_json)
    sh = gs.open_by_url("https://docs.google.com/spreadsheets/d/******")

    # PLANILHA DO DRIVE
    worksheet = sh.worksheet("faturamento")
    val = worksheet.range('A:Z')
    worksheet.batch_clear(['A:Z'])

    print("Alimentando o DRIVE")
    worksheet.update([fat_df.columns.values.tolist()] + fat_df.values.tolist(), value_input_option='USER_ENTERED')

    df_extract.drop(
        columns=['Sel', 'Finalidade', 'NF', 'CFOP', 'Tipo', 'Código', 'Cliente / Fornecedor', 'Sefaz', 'Valor NF',
                 'Situação NFe', 'Chave', 'Protocolo'], inplace=True)
    print('Preparando arquivos para forecast - PROPHET')

    fat_pro = df_extract
    fat_pro['Emissão'] = pd.to_datetime(fat_pro['Emissão'], format='%Y/%m/%d').dt.date
    fat_pro = fat_pro.groupby(['Emissão'])['Fatura'].sum()
    fat_pro = fat_pro.reset_index()
    fat_pro.columns = ['ds', 'y']
    fat_pro.to_excel('base_realizado.xlsx', index=False)
    fat_pro['cap'] = 8.5

    print('DF_PRO: ', fat_pro.head())

    m = Prophet(growth='logistic')
    m.fit(fat_pro)
    future = m.make_future_dataframe(periods=90)
    future['cap'] = 8.5
    fcst = m.predict(future)
    fig = m.plot(fcst)

    lista = []
    fcst['origem'] = 'previsão'

    fcst.drop(columns=['trend', 'cap', 'yhat_lower', 'yhat_upper', 'trend_lower', 'trend_upper', 'additive_terms',
                       'additive_terms_lower', 'additive_terms_upper', 'weekly', 'weekly_lower', 'weekly_upper',
                       'yearly',
                       'yearly_lower', 'yearly_upper', 'multiplicative_terms', 'multiplicative_terms_lower',
                       'multiplicative_terms_upper'], inplace=True)

    fcst.columns = ['ds', 'y', 'origem']

    real = pd.read_excel('base_realizado.xlsx')
    real_df = pd.DataFrame(real)
    real_df['origem'] = 'realizado'
    print(real_df.head())
    previsao = pd.concat([real_df, fcst])

    previsao = previsao[previsao['y'] > 0]
    previsao = previsao.astype(str)
    previsao.to_excel('resultados.xlsx', index=False)

    # LOCALIZAÇÃO DA PLANILHA
    gs = gspread.service_account(service_json)
    sh = gs.open_by_url("https://docs.google.com/spreadsheets/d/****")

    # PLANILHA DO DRIVE
    worksheet = sh.worksheet("prophet")
    worksheet.range('A:Z')
    worksheet.batch_clear(['A:Z'])
    print("Alimentando o DRIVE")
    worksheet.update([previsao.columns.values.tolist()] + previsao.values.tolist(), value_input_option='USER_ENTERED')

    
    now_dois = datetime.now()
    tempo_execucao = now_dois - now

    print('Iniciando indicador de faturamento')

    tarefa = pd.read_excel(r'L:\Relatórios\tarefas.xlsx', engine='openpyxl')
    tar_df = pd.DataFrame(tarefa)

    service_json = "service_account.json"

    # LOCALIZAÇÃO DA PLANILHA
    gs = gspread.service_account(service_json)
    sh = gs.open_by_url("https://docs.google.com/spreadsheets/d/****")

    # PLANILHA DO DRIVE
    worksheet = sh.worksheet("tarefas")
    val = worksheet.range('A:Z')
    worksheet.batch_clear(['A:Z'])

    print("Alimentando o DRIVE")
    worksheet.update([tar_df.columns.values.tolist()] + tar_df.values.tolist(), value_input_option='USER_ENTERED')

    print('PROCESSO FINALIZADO COM SUCESSO!' + '\n' + 'Tempo de execução do processo integral: ', str(tempo_execucao))
    
    webbrowser.open('https://datastudio.google.com/reporting/****')

    time.sleep(10)
    sys.exit()














from tkinter import filedialog
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import openpyxl
from openpyxl.drawing.image import Image
import pyperclip
import time
from selenium import webdriver
import datetime
import pyautogui as pi
import os
import json
import cv2





class Scrappy:
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_experimental_option(
            'excludeSwitches', ['enable-logging']
        )
    options.add_argument('--lang=pr-BR')
    options.add_argument('--disable-notifications')
    driver = webdriver.Chrome(options=options)
    


    listaLocCa = []
    listaGeracaoCa = []
    listaEndCa = []
    listaClimaCa = []
    
    
    listaLocPHB = []
    listaGeracaoPHB = [] 
    listaEndPHB = []
    
    
    listaLocSolis = []
    listaGeracaoSolis = []
    listaEndSolis = []
    
    listaLocGro = []
    listaGeracaoGro = []
    
    
    def iniciar(self):
        self.raspagemDeDadosCanadian()
        # self.raspagemDeDadosClima()
        self.raspagemDeDadosPHB()
        self.raspagemDeDadoSolis()
        # self.raspagemDeDadoGrowatt()
        self.criarPlanilha()
    
    def raspagemDeDadosCanadian(self):
        self.link = 'https://monitoring.csisolar.com/maintain/plant'
        self.driver.get(self.link)
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div[5]/div[3]/div/div[4]/button[1]').click()
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div[5]/div[5]/div/section/div[2]/div[2]/div[3]/div/input').send_keys('operacao@energiaarion.com.br', Keys.TAB, 'Arion2023', Keys.ENTER)
        time.sleep(10)
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div[5]/div[2]/section/div/section/div[4]/div[1]/div[1]/div[2]/a').click()
        print('\033[42m'+'\033[1m'+'\033[33m'+'Vamos começar a varredura CANADIAN☀️'+'\033[0;0m')
        value = 1
        time.sleep(5)
        for n in range(1):
            for i in range(16):
                lista_localizacao = self.driver.find_elements(By.XPATH , f'/html/body/div[1]/div[5]/div[2]/section/div/div[5]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div/table/tbody/tr[{value}]/td[2]/div[1]')
                self.listaLocCa.append(lista_localizacao[0].get_attribute('innerHTML'))
                lista_geracao = self.driver.find_element(By.XPATH , f'/html/body/div[1]/div[5]/div[2]/section/div/div[5]/div[2]/div[1]/div[1]/div[3]/div[1]/div/div[2]/div[1]/div/table/tbody/tr[{value}]/td[4]')
                self.listaGeracaoCa.append(lista_geracao.get_attribute('innerHTML'))
                lista_end = self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/div[2]/section/div/div[5]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div/table/tbody/tr[{value}]/td[2]/div[2]')
                self.listaEndCa.append(lista_end.get_attribute('title'))
                lista_clima =  self.driver.find_element(By.XPATH, f'//*[@id="mCSB_42_container"]/table/tbody/tr[{value}]/td[9]')
                self.listaClimaCa.append(lista_clima.get_attribute('innerHTML'))
                value += 1
        print('\033[42m'+'\033[1m'+'\033[33m'+'varredura CANADIAN terminou☀️'+'\033[0;0m')
        time.sleep(10)
      
    def raspagemDeDadosPHB(self):
        self.driver.set_window_size(800, 700)
        self.link = 'http://www.phbsolar.com.br/home/login' 
        self.driver.get(self.link)  
        self.driver.find_element(By.XPATH, '//*[@id="username"]').send_keys('monitoramentoarion.sfv@gmail.com', Keys.TAB, 'solar2021', Keys.ENTER)
        print('\033[42m'+'\033[1m'+'\033[33m'+'Vamos começar a varredura PHB☀️'+'\033[0;0m')
        time.sleep(10)
        value = 1
        for n in range(1):
            for i in range(9):
                lista_localizacao = self.driver.find_elements(By.XPATH , f'//*[@id="app"]/div[5]/div[2]/div/ul[{value}]/li[2]/span[1]')
                self.listaLocPHB.append(lista_localizacao[0].get_attribute('innerHTML'))
                lista_geracao = self.driver.find_element(By.XPATH , f'//*[@id="app"]/div[5]/div[2]/div/ul[{value}]/li[8]')
                self.listaGeracaoPHB.append(lista_geracao.get_attribute('innerHTML'))
                lista_end = self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[5]/div[2]/div/ul[{value}]/li[3]')
                self.listaEndPHB.append(lista_end.get_attribute('title'))
                value += 1
        print('\033[42m'+'\033[1m'+'\033[33m'+'varredura PHB terminou☀️'+'\033[0;0m')
            
    def raspagemDeDadoSolis(self):
        self.driver.set_window_size(800, 700)
        self.link = 'https://www.soliscloud.com/#/homepage' 
        self.driver.get(self.link)  
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div[1]/div[2]/div/div/div[2]/div/div[1]/div[2]/form/div[1]/div/div/div[1]/input').send_keys('solar@energiaarion.com.br')
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div[1]/div[2]/div/div/div[2]/div/div[1]/div[2]/form/div[2]/div/div/input').send_keys('Arion2021')
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div[1]/div[2]/div/div/div[2]/div/div[1]/div[2]/form/div[4]/span[1]/label/span/span').click()
        self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div[1]/div[2]/div/div/div[2]/div/div[1]/div[2]/form/div[5]/button').click()
        print('\033[42m'+'\033[1m'+'\033[33m'+'Vamos começar a varredura SOLIS☀️'+'\033[0;0m')
        time.sleep(10)
        value = 1
        for i in range(9):
            lista_localizacao = self.driver.find_elements(By.XPATH , f'//*[@id="station"]/div[1]/div[3]/div[3]/div[2]/div[2]/div[1]/div[3]/table/tbody/tr[{value}]/td[2]/div/div/div/div[1]/span/span[3]')
            self.listaLocSolis.append(lista_localizacao[0].get_attribute('innerHTML'))
            lista_geracao = self.driver.find_element(By.XPATH , f'/html/body/div[1]/div/div[4]/div[1]/div[3]/div[3]/div[2]/div[2]/div[1]/div[3]/table/tbody/tr[{value}]/td[4]/div')
            self.listaGeracaoSolis.append(lista_geracao.get_attribute('innerHTML'))
            lista_end = self.driver.find_element(By.XPATH, f'/html/body/div[1]/div/div[4]/div[1]/div[3]/div[3]/div[2]/div[2]/div[1]/div[3]/table/tbody/tr[{value}]/td[2]/div/div/div/div[2]/span/span[3]')
            self.listaEndSolis.append(lista_end.get_attribute('innerHTML'))
            value += 1
        print('\033[42m'+'\033[1m'+'\033[33m'+'varredura SOLIS terminou☀️'+'\033[0;0m')
    
    def raspagemDeDadosClima(self):
        pi.PAUSE = 5
        print('\033[42m'+'\033[1m'+'\033[33m'+'Vamos começar a varredura nos climas☀️'+'\033[0;0m')
        self.linkTemp = 'https://www.google.com.br/?hl=pt-BR'
        self.driver.get(self.linkTemp)
        for end in self.listaEndCa:
            element = ','
            endTratado = end.index(element)
            print(endTratado)
            self.driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(f'{end[0: endTratado]} clima hoje clima tempo', Keys.ENTER)
            self.driver.find_element(By.XPATH, '//*[@id="APjFqb"]').clear()
            
            time.sleep(5)
            # print('\033[32m'+'tempo ----------------- 33,33% '+'\033[0;0m')
            # time.sleep(3)
            # self.driver.find_element(By.XPATH, '//*[@id="rso"]/div[1]/div/div/div[1]/div/div/span/a/h3').click()
            # self.driver.find_element(By.XPATH, '//*[@id="fixedNavigation"]/ul/li[1]/a').click()
            # graus = self.driver.find_element(By.XPATH, '//*[@id="mainContent"]/div[5]/div[4]/div[1]/div[2]/div[1]/div/div[2]/span')
            # tempo = self.driver.find_element(By.XPATH, '//*[@id="mainContent"]/div[5]/div[4]/div[1]/div[2]/div[1]/div/div[3]/span[1]')
            # os. system('cls')
            # self.listaClimaCa.append(tempo.get_attribute('innerHTML'))
            # os. system('cls')
            # self.listaClimaCa.append(graus.get_attribute('innerHTML'))
            # os. system('cls')
            # print(self.listaClimaCa) 
        print('\033[42m'+'\033[1m'+'\033[33m'+'A varredura terminou☀️'+'\033[0;0m')
    
    

    def criarPlanilha(self):
        listaTratada =  list(map(lambda x: x.replace('.', ','), self.listaGeracaoCa))
        listaTratada1 = list(map(lambda x: x.replace('.', ',').replace('kWh', ''), self.listaGeracaoPHB))
        listaTratada2 = list(map(lambda x: x.replace('.', ',').replace('kWh', ''), self.listaGeracaoSolis))
        # for clima in self.listaClimaCa:
        #     if clima == 'Ensolarado':
        #         print(clima)
               
        #         climaTratado = clima
        #     else:
        #         climaTratado = self.listaClimaCa
        #     print(climaTratado)
        self.link = 'https://www.google.com/search?q=data+e+hora+agora&sca_esv=579945731&rlz=1C1RXQR_pt-PTBR1081BR1081&sxsrf=AM9HkKmNoO_METo2UWS2xrUqtwxmqPb41g%3A1699312863267&ei=33RJZZf4D7XY5OUPnq-j4AU&ved=0ahUKEwjX487cwbCCAxU1LLkGHZ7XCFwQ4dUDCBA&uact=5&oq=data+e+hora+agora&gs_lp=Egxnd3Mtd2l6LXNlcnAiEWRhdGEgZSBob3JhIGFnb3JhMgUQABiABDIGEAAYFhgeMgYQABgWGB4yBhAAGBYYHjIGEAAYFhgeMgYQABgWGB4yBhAAGBYYHjIGEAAYFhgeMgYQABgWGB4yBhAAGBYYHkjEG1DkAVjMGHABeAGQAQCYAd8BoAHMBqoBBTAuNS4xuAEDyAEA-AEBwgIKEAAYRxjWBBiwA8ICChAAGIoFGLADGEPCAgsQABiABBixAxiDAcICChAAGBYYHhgPGAriAwQYACBBiAYBkAYK&sclient=gws-wiz-serp' 
        self.driver.get(self.link)
        data =  datetime.date.today()
        self.dataStr = data.strftime("%d/%m/%Y")
        hora = self.driver.find_element(By.XPATH, '//*[@id="rso"]/div[1]/div/div/div/div/div/div[1]').get_attribute('innerHTML')
        localizacao = self.listaLocCa
        planilha = openpyxl.Workbook()
        geracao = planilha['Sheet']
        geracao.title = 'Usinas Criativa Energia'
        geracao['A1'] = 'Nome'
        geracao['B1'] = 'Endereço'
        geracao['C1'] = 'Geração Diaria'
        geracao['D1'] = 'Clima'
        geracao['E1'] = 'Data/Hora'
        index = 2
        
        for localizacao, geracaoDiaria, end in zip(
            self.listaLocCa, listaTratada, self.listaEndCa):
            time.sleep(1)
            geracao.cell(column=1, row=index, value=localizacao)
            geracao.cell(column=2, row=index,  value=end)
            element = ","
            GerTratada = geracaoDiaria.index(element)
            print(geracaoDiaria[0: GerTratada])
            geracao.cell(column=3, row=index).value= fr"=VALOR({geracaoDiaria[0:GerTratada]}"
            geracao.cell(column=4, row=index).value=r'=IMAGEM("https://static.semsportal.com/static/Images/icon-weather-new/101.png";"CLIMA")'
            geracao.cell(column=5, row=index, value=f'{self.dataStr}, {hora}')
            index += 1
        index = 18                    
        print('\033[32m'+'Planilha ----------- 33,33% '+'\033[0;0m') 
        
        for localizacao1, geracaoDiaria1, end in zip(self.listaLocPHB, listaTratada1, self.listaEndPHB):
            geracao.cell(column=1, row=index, value=localizacao1)
            geracao.cell(column=2, row=index,  value=end)
            geracao.cell(column=3, row=index, value=geracaoDiaria1)
            geracao.cell(column=5, row=index, value=f'{self.dataStr}, {hora}')
            index +=1   
        index = 27
        
        
        print('\033[32m'+'Planilha ----------------- 33,33% '+'\033[0;0m') 
        time.sleep(3)
        
        
        for localizacao2, geracaoDiaria2, end in zip(self.listaLocSolis, listaTratada2, self.listaEndSolis): 
            geracao.cell(column=1, row=index, value=localizacao2)
            geracao.cell(column=2, row=index,  value=end)
            geracao.cell(column=3, row=index, value=geracaoDiaria2)
            geracao.cell(column=5, row=index, value=f'{self.dataStr}, {hora}')
            index += 1 
            
        
        planilha.save('Criativa.xlsx')
        print('\033[32m'+'Planilha criada com sucesso☀️🚩------------------------------------------100%😜'+'\033[0;0m') 
        pi.alert("O PROCESSO DE ATUALIZAÇÃO TERMINOU ☀️🚩😜!! A PLANILHA ESTÁ NA MESMA PASTA DO AQUIVO DO CODIGO 😜🧧✅")

    
                

start = Scrappy()
start.iniciar()
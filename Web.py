import smtplib
import time
from email.message import EmailMessage

import openpyxl as openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

class CaptadorBoletos():

    def get_options(self):
        self.options = Options()
        self.options.add_argument("--window-size=1920,1080")
        self.options.add_argument("--start-maximized")
        self.options.add_argument('--headless')
        return self.options

    navegador = webdriver.Chrome(options=get_options(self=Options))

    def atenticacao(self):
        self.usuario = input("Digite o usuário(RA ou CPF): ")
        self.senha = input("Digite a senha: ")

    def captura_boleto(self):
        print("Estabelecendo conexão com o site...")
        self.navegador.get("https://www.usf.edu.br/")
        self.navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(self.usuario)
        self.navegador.find_element(By.XPATH, '//*[@id="senha"]').send_keys(self.senha)
        self.navegador.find_element(By.XPATH, '/html/body/div[4]/div[1]/div[3]/form/div[4]/div/input').click()
        self.navegador.get('https://www.usf.edu.br/apps/portalaluno2/boleto')

    def captura_tabela(self):
        print("Capturando tabela de boletos...")
        item = 1
        qtd = self.navegador.find_elements(By.TAG_NAME, 'table')
        self.lista_boleto = []
        try:
            for i in range(len(qtd)):
                element = self.navegador.find_elements(By.XPATH, f'//*[@id="accordionBoleto"]/div/div[{item}]')
                #html_content = div.get_attribute('outerHTML')
                #soup = BeautifulSoup(html_content, 'html.parser')
                #table = soup.find(name='table')
                #df = p.read_html(str(table))[0].drop(columns=[4,5])
                self.lista_boleto.append(element[0].text)
                item += 1
        except:
            print("Tabela Capturada")

    def cria_excel(self):
        index = 2
        planilha = openpyxl.Workbook()
        boletos = planilha['Sheet']
        boletos.title = 'Lista de boletos'
        boletos['A1'] = 'Todos os boletos'
        for lista in self.lista_boleto:
            boletos.cell(column=1, row=index, value=lista)
            index+=1
        planilha.save("planilha_Boletos.xlsx")
        print("Planilha criada")

    def envia_email(self, email, senha):
        destinatario = input("Digite o e-mail que irá receber a planilha: ")
        print('Enviando e-mail...')
        msg = EmailMessage()
        msg['Subject'] = 'Planilha de boletos USF'
        msg['From'] = email
        msg['To'] = destinatario
        msg.set_content('Olá, Sua planilha com os boletos chegou.')
        arquivos = ["planilha_Boletos.xlsx"]
        for arquivo in arquivos:
            with open(arquivo, 'rb') as arq:
                dados = arq.read()
                nome_arquivo = arq.name
            msg.add_attachment(dados, maintype='application', subtype='octet-stream', filename=nome_arquivo)
        server = smtplib.SMTP('smtp.outlook.com', 587)
        server.ehlo()
        server.starttls()
        server.login(email, senha, initial_response_ok=True)
        server.send_message(msg)
        print('E-mail enviado')
        server.quit()












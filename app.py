import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

# Configuração do Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Executa o Chrome em modo headless
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# URL alvo
url = "https://www.novaliderinformatica.com.br/computadores-gamers"
driver.get(url)

# Aguarda a página carregar
time.sleep(5)

# Coleta os títulos e preços
titulos = driver.find_elements(By.XPATH, "//a[@class='nome-produto']")
precos = driver.find_elements(By.XPATH, "//strong[@class='preco-promocional']")

# Criando a planilha
workbook = openpyxl.Workbook()
sheet_produtos = workbook.active
sheet_produtos.title = 'Produtos'
sheet_produtos['A1'] = 'Produto'
sheet_produtos['B1'] = 'Preço'

# Inserir os títulos e preços na planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])

# Salva a planilha
workbook.save('produtos_gamers.xlsx')

# Fecha o navegador
driver.quit()

print("Planilha 'produtos_gamers.xlsx' criada com sucesso!")
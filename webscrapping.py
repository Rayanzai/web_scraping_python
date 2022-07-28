from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd


#Abrindo o Navegador
navegador = webdriver.Chrome()


###########PROCURANDO A COTAÇÃO DO DOLAR
navegador.get("https://www.google.com/")
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dolar')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)
cotacao_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)


###########PROCURANDO A COTAÇÃO DO EURO
navegador.get('https://www.google.com/')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

###########PROCURANDO A COTAÇÃO DO OURO
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element('xpath', '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',', '.')
print(cotacao_ouro)

#####FECHANDO NAVEGADOR
navegador.quit()


#ABRINDO PANDAS E LEITURA DE PLANILHA
df = pd.read_excel('Produtos.xlsx')

# .loc[linha, coluna]
df.loc[df['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar)
df.loc[df['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro)
df.loc[df['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_ouro)

#print(df)

###SUBSTITUINDO OS VALORES DA PLANILHA, POIS AS FORMULAS DA PLANILHA NÃO FUNCIONARAM 
df['Preço de Compra'] = df['Preço Original'] * df['Cotação']
df['Preço de Venda'] = df['Preço de Compra'] * df['Margem']

print(df)

#######EXPORTA O ARQUIVO PARA UMA NOVA TABELA EXCEL E RETIRA OS INDEX DO PANDAS
df.to_excel('Novos Produtos.xlsx', index=False)



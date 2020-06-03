from selenium import webdriver
import xlsxwriter as xw


# from bs4 import BeautifulSoup

# f = open(r"C:\Users\Hugo\Downloads\casbantigo\casbantigo\index.html", encoding="utf8")  
# soup = BeautifulSoup(f)
# print (soup)
driver = webdriver.Firefox()
workbook = xw.Workbook("Eventos.xlsx")
worksheet = workbook.add_worksheet("Eventos")
print ("Criando Planilha...")

def getConteudo():
    row = 0
    print("Colhendo o conteúdo...")
    for i in range(1,21):
        try: 
            # SELECIONA O EVENTO
            driver.find_element_by_xpath("""/html/body/div[5]/div[1]/div[1]/div/div[3]/div[%s]/div/a""" % i).click()

            # SELECIONA CLASSE DO CONTEUDO E TITULO
            title = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[3]/h1""").text
            evento_info = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[3]/div""").text
            date = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[3]/p""").text
            conteudo = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[4]""").text

            worksheet.write(row, 0, title)
            worksheet.write(row, 1, evento_info)
            worksheet.write(row, 2, date)
            worksheet.write(row, 3, conteudo)


            # print("Progresso: " + str(round((i/21)*100)) + "%")
            print (row)
            row = row + 1
            driver.back()
        except:
            error_url = driver.current_url()
            driver.back()

            print("Erro encontrado na página: " + str(error_url))
            print("Cotninuando processo...")
            continue

driver.get("file:///C:/Users/Hugo/Downloads/casbantigo/casbantigo/agenda-e-eventos/eventos-passados/index.html")
getConteudo()
workbook.close()
print ("Fim!")
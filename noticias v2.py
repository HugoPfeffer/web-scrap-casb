from selenium import webdriver
from selenium.common.exceptions import TimeoutException
import xlsxwriter as xw

workbook = xw.Workbook("scrap2.xlsx")
worksheet = workbook.add_worksheet("Noticias")
# worksheet_error_page = workbook.add_worksheet("Erros de página")
# worksheet_error_content = workbook.add_worksheet("Erros de conteudo")
print ("Começando Scrap...")

driver = webdriver.Firefox()

global row
row = 0

def scrapPageError():
    global row
    url_error = driver.current_url
    worksheet.write(row, 3, url_error)
    row += 1
    driver.back()

def scrapContentError():
    global row
    url_error = driver.current_url
    worksheet.write(row, 3, url_error)
    row += 1
    driver.back()

# for i in range(1,20,2):
def getConteudo(x):
    driver.set_page_load_timeout(10)
    try:
        # SELECIONA NOTICIA
        driver.find_element_by_xpath("""//*[@id="body"]/div/div[3]/div[%s]/h3/a""" % x).click()
        

        # SELECIONA CLASSE DO CONTEUDO E TITULO
        header_element = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[3]""")
        title = header_element.find_element_by_tag_name("h1").text
        date = header_element.find_element_by_tag_name("p").text

        post_element = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[4]""")
        conteudo = post_element.text
        driver.back()
        return title, date, conteudo
    except: 
        driver.execute_script("window.stop();")
        

def writePlanilha(z):
    global row
    worksheet.write(row, 0, getConteudo(z)[0])
    worksheet.write(row, 1, getConteudo(z)[1])
    worksheet.write(row, 2, getConteudo(z)[2])
    row += 1


# for i in range(13):
def pageNext(x):
    driver.set_page_load_timeout(10)
    try:
        driver.get("file:///C:/Users/Hugo/Downloads/casbantigo/casbantigo/noticias/index_ccm_paging_p_b400=%s.html" % (x))
    except TimeoutException:
        driver.execute_script("window.stop();")
        driver.back()







pageNext(10)
for z in range(1,20,2):
    try:
        getConteudo(z)
        writePlanilha(z)
    except:
        scrapContentError()
        print("Erro ao extrair conteudo")

# for i in range(1,13):
#     try:
#         pageNext(i)
#         print("Progresso: " + str(round(((i/13)*100), 2)) + "%")
#         for z in range(1,20,2):
#             try:
#                 getConteudo(z)
#                 writePlanilha(z)
#             except:
#                 scrapContentError()
#                 print("Erro ao extrair conteudo")
#     except: 
#         print("Erro ao entrar na página")
#         scrapPageError()
#         continue



print("Scrap Finalizado!")
workbook.close()

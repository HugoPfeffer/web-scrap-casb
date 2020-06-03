from selenium import webdriver
import xlsxwriter as xw


# from bs4 import BeautifulSoup

# f = open(r"C:\Users\Hugo\Downloads\casbantigo\casbantigo\index.html", encoding="utf8")  
# soup = BeautifulSoup(f)
# print (soup)

workbook = xw.Workbook("scrap.xlsx")
worksheet = workbook.add_worksheet("Noticias")
worksheet_error_page = workbook.add_worksheet("Erros de página")
worksheet_error_content = workbook.add_worksheet("Erros de conteudo")
print ("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")

driver = webdriver.Firefox()
# driver.get(r"file:///C:\Users\Hugo\Downloads\casbantigo\casbantigo\index.html")

# ABRE PÁGINA DE NOTICIAS
# driver.find_element_by_xpath("""//*[@id="innerconteudo"]/div[1]/div[1]/div[7]/a""").click()


def getConteudo(x):
    '''Abre a página de noticias a partir do index.html seleciona a noticia, pega o conteudo e coloca em um xlsx''' 
    try:  
        # for i in range(1,20,2):
        for i in range(1,12,2):
            try:
                # SELECIONA NOTICIA
                driver.find_element_by_xpath("""//*[@id="body"]/div/div[3]/div[%s]/h3/a""" % i).click()

                # SELECIONA CLASSE DO CONTEUDO E TITULO
                header_element = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[3]""")
                title = header_element.find_element_by_tag_name("h1").text
                date = header_element.find_element_by_tag_name("p").text

                post_element = driver.find_element_by_xpath("""/html/body/div[5]/div/div[1]/div/div[4]""")
                conteudo = post_element.text

                worksheet.write(row, 0, title)
                worksheet.write(row, 1, date)
                worksheet.write(row, 2, conteudo)


                print (title)
                print (date)
                print (conteudo)
                # print (type(title.text))
                row += 1
                driver.back()
            # except:
            #     url_error = driver.current_url
            #     driver.back()
            #     worksheet_error_content.write(row_error, 0, url_error)
            #     row_error += 1
            #     print ("erro ao carregar conteudo ")
            #     driver.back()
            #     continue
    except NameError:
        row = 0
        row_error = 0
        pass
    
    return row, row_error


def setError():
    url_error = driver.current_url
    driver.back()
    worksheet_error_content.write(row_error, 0, url_error)
    row_error += 1
    print ("erro ao carregar conteudo ")
    driver.back()

# PAGINAÇÃO
row_error = 0
# for i in range(13):
x = getConteudo(0)
for i in range(3):
    try: 
        driver.get("file:///C:/Users/Hugo/Downloads/casbantigo/casbantigo/noticias/index_ccm_paging_p_b400=%s.html" % (i+1))
        url = driver.current_url
        x[0] = getConteudo(x)
    except:
        url_error = driver.current_url
        driver.back()
        worksheet_error_page.write(x[1], 0, url_error)
        x[1] += 1
        print ("erro ao carregar página ")

        continue
    

# HEADERS DA PLANILHA 
# worksheet.write (0, 0, "Título")
# worksheet.write (0, 2, "Enviado por")
# worksheet.write (0, 1, "Conteúdo")



workbook.close()

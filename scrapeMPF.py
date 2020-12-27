from selenium import webdriver
import xlsxwriter
from datetime import datetime

now = (datetime.now()).strftime("%d-%m-%Y_%H-%M_")


PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

workbook = xlsxwriter.Workbook(now + 'mpf.xlsx')
worksheet = workbook.add_worksheet("MPF") 
worksheet2 = workbook.add_worksheet("RJ") 
worksheet3 = workbook.add_worksheet("AC") 
worksheet4 = workbook.add_worksheet("AL") 
worksheet5 = workbook.add_worksheet("AP") 
worksheet6 = workbook.add_worksheet("AM") 
worksheet7 = workbook.add_worksheet("BA") 
worksheet8 = workbook.add_worksheet("CE") 
worksheet9 = workbook.add_worksheet("ES") 
worksheet10 = workbook.add_worksheet("GO") 
worksheet11 = workbook.add_worksheet("MA") 
worksheet12 = workbook.add_worksheet("MT") 
worksheet13 = workbook.add_worksheet("MS") 
worksheet14 = workbook.add_worksheet("MG") 
worksheet15 = workbook.add_worksheet("PA") 
worksheet16 = workbook.add_worksheet("PB") 
worksheet17 = workbook.add_worksheet("PR") 
worksheet18 = workbook.add_worksheet("PE") 
worksheet19 = workbook.add_worksheet("PI") 
worksheet20 = workbook.add_worksheet("RN") 
worksheet21 = workbook.add_worksheet("RS") 
worksheet22 = workbook.add_worksheet("RO") 
worksheet23 = workbook.add_worksheet("RR") 
worksheet24 = workbook.add_worksheet("SC") 
worksheet25 = workbook.add_worksheet("SP") 
worksheet26 = workbook.add_worksheet("SE") 
worksheet27 = workbook.add_worksheet("TO") 
worksheet28 = workbook.add_worksheet("DF") 


#Open the website - MPF
driver.get("http://www.mpf.mp.br/sala-de-imprensa/noticias")

#Take articles from the first page
articles = driver.find_elements_by_tag_name("article")
row = 0
col = 0
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet.write(row, col,     data.text)
        worksheet.write(row, col + 1, header.text)
        worksheet.write(row, col + 2, link.get_attribute("href"))
        row += 1      
       
#Go to the second page 
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the second page
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet.write(row, col,     data.text)
        worksheet.write(row, col + 1, header.text)
        worksheet.write(row, col + 2, link.get_attribute("href"))
        row += 1      
           
#Go to the third page
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the third page
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet.write(row, col,     data.text)
        worksheet.write(row, col + 1, header.text)
        worksheet.write(row, col + 2, link.get_attribute("href"))
        row += 1      

#Open the website - RJ
driver.get("http://www.mpf.mp.br/rj/sala-de-imprensa")

#Take articles from the first page - RJ
articles = driver.find_elements_by_tag_name("article")
row = 0
col = 0
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet2.write(row, col,     data.text)
        worksheet2.write(row, col + 1, header.text)
        worksheet2.write(row, col + 2, link.get_attribute("href"))
        row += 1      
       
#Go to the second page - RJ
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the second page - RJ
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet2.write(row, col,     data.text)
        worksheet2.write(row, col + 1, header.text)
        worksheet2.write(row, col + 2, link.get_attribute("href"))
        row += 1      
           
#Go to the third page - RJ
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the third page - RJ
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet2.write(row, col,     data.text)
        worksheet2.write(row, col + 1, header.text)
        worksheet2.write(row, col + 2, link.get_attribute("href"))
        row += 1      

#Open the website - AC
driver.get("http://www.mpf.mp.br/ac/sala-de-imprensa")

#Take articles from the first page - AC
articles = driver.find_elements_by_tag_name("article")
row = 0
col = 0
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet3.write(row, col,     data.text)
        worksheet3.write(row, col + 1, header.text)
        worksheet3.write(row, col + 2, link.get_attribute("href"))
        row += 1      
       
#Go to the second page - AC
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the second page - AC
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet3.write(row, col,     data.text)
        worksheet3.write(row, col + 1, header.text)
        worksheet3.write(row, col + 2, link.get_attribute("href"))
        row += 1      
           
#Go to the third page - AC
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the third page - AC
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet3.write(row, col,     data.text)
        worksheet3.write(row, col + 1, header.text)
        worksheet3.write(row, col + 2, link.get_attribute("href"))
        row += 1              

#Open the website - AL
driver.get("http://www.mpf.mp.br/al/sala-de-imprensa")

#Take articles from the first page - AL
articles = driver.find_elements_by_tag_name("article")
row = 0
col = 0
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet4.write(row, col,     data.text)
        worksheet4.write(row, col + 1, header.text)
        worksheet4.write(row, col + 2, link.get_attribute("href"))
        row += 1      
       
#Go to the second page - AL
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the second page - AL
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet4.write(row, col,     data.text)
        worksheet4.write(row, col + 1, header.text)
        worksheet4.write(row, col + 2, link.get_attribute("href"))
        row += 1      
           
#Go to the third page - AC
search = driver.find_element_by_class_name("next")
search.click()    

#Take articles from the third page - AC
articles = driver.find_elements_by_tag_name("article")
for article in articles:
        data = article.find_element_by_class_name("data")
        header = article.find_element_by_tag_name("h2")
        link = article.find_element_by_tag_name("a")
        worksheet4.write(row, col,     data.text)
        worksheet4.write(row, col + 1, header.text)
        worksheet4.write(row, col + 2, link.get_attribute("href"))
        row += 1              

workbook.close()      

driver.close()

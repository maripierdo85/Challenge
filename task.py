"""Template robot with Python."""
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from datetime import datetime
import time
browser_lib = Selenium()
browser_lib2=Selenium()
lib = Files()
table = Tables()
dwPath = "output"
def open_the_website(url):
    #download_preferences = {"download.default_directory": dwPath}
    #browser_lib.set_download_directory(directory=dwPath)
    browser_lib.open_available_browser(url)
def close_the_website():
    browser_lib.close_browser()
def click_button(xpath):
    dive_element = "xpath:%s" % (xpath)
    browser_lib.find_element(dive_element).click()
def agency_totals():
    lista = browser_lib.find_elements("xpath://div[@id='agency-tiles-widget']")
    for symbol in lista:
        element_text = symbol.text
        formatted = element_text.split("\n")
    return formatted
def create_worksheet(nameW):
    lib.create_worksheet(nameW)
def write_excel_worksheet(path, nameW, result):
    lista = []
    excel = lib.open_workbook(path)
    date = datetime.today().strftime('%Y-%m-%d %H:%M')
    for i in range(len(result)):
        if result[i] == "view":
            agencies = result[i-3]
            amount = result[i-1]
            lista.append([date, agencies, amount])
    headers = ['Datetime', 'Agency', 'Amount']
    tablaExcel = table.create_table(data=lista, columns=(headers))
    lib.append_rows_to_worksheet(tablaExcel, nameW, headers)
    lib.save_workbook(path)
def close_excel_file(path):
    lib = Files()
    lib.close_workbook(path)
def get_max_pag():
    pages = "xpath://*[@id='investments-table-object_paginate']/span/a"
    listaPAgs = browser_lib.find_elements(pages)
    time.sleep(30)
    span = "/span/a[%s]" % (len(listaPAgs))
    element = "xpath://*[@id='investments-table-object_paginate']%s" % (span)
    time.sleep(30)
    txtMaxPag = browser_lib.find_element(element).text
    return txtMaxPag
def get_headers():
    div = "/div[3]/div[1]/div/table/thead/tr[2]/th"
    element = "xpath://*[@id='investments-table-object_wrapper']%s" % (div)
    listaEncabezado = browser_lib.find_elements(element)
    encabezado = []
    for symbol in listaEncabezado:
        element_text = symbol.text
        encabezado.append(element_text)
    result = [len(listaEncabezado), encabezado]
    return result
def individual_investment(path):
    element1 = "xpath://*[@id='investments-table-object']/tbody/tr/td/a"
    listaFilasLinks = browser_lib.find_elements(element1)
    element2 = "xpath://*[@id='investments-table-object']/tbody/tr"
    listaFilas = browser_lib.find_elements(element2)
    lenHeaders = get_headers()[0]
    headers = get_headers()[1]
    filaGeneral = []
    for f in range(len(listaFilas)):
        filas = []
        for c in range(lenHeaders):
            el3 = "/tbody/tr[%s]/td[%s]" % (f+1, c+1)
            element4 = "xpath://*[@id='investments-table-object']%s" % (el3)
            fila = browser_lib.find_element(element4).text
            filas.append(fila)
        filaGeneral.append(filas)
    time.sleep(1)
    tablaExcel = table.create_table(data=filaGeneral, columns=headers)
    time.sleep(1)
    lib.open_workbook(path)
    time.sleep(1)
    lib.append_rows_to_worksheet(tablaExcel, "Investments", headers)
    time.sleep(1)
    lib.save_workbook(path)
    time.sleep(1)
    filasLinks = []
    if len(listaFilasLinks)>0:
        for h in listaFilasLinks:
            href = h.get_attribute('href')
            time.sleep(5)
            #browser_lib2.op
            open_the_website(href)
            time.sleep(15)
            element5 = "xpath://*[@id='business-case-pdf']/a"
            browser_lib.find_element(element5).click()
            time.sleep(20)      
            browser_lib.close_browser()
            time.sleep(20)  
    time.sleep(20)
    click_button("//*[@id='investments-table-object_next']")
def minimal_task():
    try:
        path = "output/amounts.xlsx"      
        open_the_website("https://itdashboard.gov/")
        browser_lib.set_browser_implicit_wait(15)
        click_button("//*[@id='node-23']/div/div/div/div/div/div/div/a")
        browser_lib.set_browser_implicit_wait(15)
        totals = agency_totals()
        write_excel_worksheet(path, 'Agencies', totals)
        browser_lib.set_browser_implicit_wait(15)
        div = "//*[@id='agency-tiles-widget']/div/div[1]/div[1]/div/div/div/div[2]/a"
        click_button(div)
        #time.sleep(5)
        browser_lib.set_browser_implicit_wait(15)
        max_pag = get_max_pag()
        for i in range(int(max_pag)):
            print(individual_investment(path))
            time.sleep(10)
        time.sleep(10)
    finally:
        close_the_website()
if __name__ == "__main__":
    minimal_task()
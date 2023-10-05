from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import openpyxl as opx

def getDataToArchive(path, colQtd, ixtable = 0, colName = 2, colUn = 3, colCod = 1):
    ##* Função reutilzada de versões anteriores
    ##! Refatorar código
    catologo = opx.load_workbook(filename=path)
    tabelaEscolha = catologo[catologo.sheetnames[ixtable]]
    lista = []
    for i in range(2, tabelaEscolha.max_row):
        line = {
                "codCliente": tabelaEscolha.cell(row=i, column=colCod).value,
                "nomeCliente": tabelaEscolha.cell(row=i, column=colName).value,
                "quantidade": tabelaEscolha.cell(row=i, column=colQtd).value,
                "unidade": tabelaEscolha.cell(row=i, column=colUn).value
            }
        if tabelaEscolha.cell(row=i, column=2).value != None:
            lista.append(line)
        
    return lista

def ChangeCatalogInSystem(browser):
    select = Select(browser.find_element(By.XPATH, "/html/body/div[@id='div_principal']/div[@id='div_principal']/div[@id='conteudo_template']/span[@id='conteudo']/form[@id='cadastro_itens']/div/div/div/div[1]/select"))
    select.select_by_value('500')
    
    sleep(3)
    
    navigator = browser.find_element(By.XPATH, "/html/body/div[@id='div_principal']/div[@id='div_principal']/div/span/form/div/div/div/div[1]/span[2]")
    navigator = navigator.find_elements(By.TAG_NAME, 'a')
    
    for i in range(len(navigator)):
        if i != 0:
            navigator[i].click();
            sleep(4)
            
        rows = browser.find_elements(By.XPATH, "/html/body/div[@id='div_principal']/div[@id='div_principal']/div[@id='conteudo_template']/span[@id='conteudo']/form/div/div/div/div/table/tbody/tr")
        
        for row in rows:
            cels = row.find_elements(By.TAG_NAME, 'td')

            if cels[5].text.isdigit():
                print(cels[2].text, cels[5].text)
                # item = analyzeItem(cels[2].text, int(cels[5].text))
                # if item != False:
                #     ix = 0
                #     while True:
                #         try:
                #             btsEdit = cels[13].find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'a')
                #             btsEdit[0].click()
                            
                #             inputQtd = cels[5].find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'div')[1].find_element(By.TAG_NAME, 'span').find_elements(By.TAG_NAME, 'input')
                #             inputQtd[0].click()
                #             inputQtd[0].clear()
                #             inputQtd[0].send_keys(item)

                                
                #             inputQtd = cels[8].find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'div')[1].find_element(By.TAG_NAME, 'span').find_elements(By.TAG_NAME, 'input')
                #             inputQtd[0].click()
                #             inputQtd[0].clear()
                #             inputQtd[0].send_keys(item)
                            
                #             btsEdit[1].click()
                #             break
                        
                #         except:
                #             ix += 1
                #             if ix == 5:
                #                 break
                #             else:
                #                 sleep(1)
            
            
        navigator = browser.find_element(By.XPATH, "/html/body/div[@id='div_principal']/div[@id='div_principal']/div/span/form/div/div/div/div[1]/span[2]")
        navigator = navigator.find_elements(By.TAG_NAME, 'a')
    
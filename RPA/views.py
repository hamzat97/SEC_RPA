from selenium.webdriver.support.ui import Select
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.by import By       
from django.http import HttpResponse 
from django.shortcuts import render           
from selenium import webdriver          
from bs4 import BeautifulSoup
from datetime import datetime
from pandas import DataFrame
from datetime import date 
import pandas as pd
import time

def UI(request):
    if request.method == 'POST':
        try:    
            BFM = date.today() + relativedelta(months=-5)
            Company = request.POST['company']
            Chrome_options = webdriver.ChromeOptions()
            #Chrome_options.add_argument('--headless=new')
            chromedriver_path = '/usr/bin/chromedriver'
            driver = webdriver.Chrome(executable_path=chromedriver_path, options=Chrome_options)
            driver.maximize_window()
            driver.get('https://www.sec.gov/edgar/search/')
            time.sleep(4)
            driver.find_element(By.ID, 'entity-short-form').send_keys(Company)
            time.sleep(4)
            driver.find_element(By.ID, 'search').click()
            time.sleep(4)
            driver.find_element(By.ID, 'keywords').clear()
            driver.find_element(By.ID, 'keywords').send_keys('13F')
            time.sleep(4)
            driver.find_element(By.ID, 'entity-full-form').send_keys(Company)
            time.sleep(4)
            select = Select(driver.find_element(By.ID, 'date-range-select'))
            select.select_by_value('custom')
            time.sleep(4)
            driver.find_element(By.ID, 'date-from').click()
            time.sleep(4)
            select = Select(driver.find_element(By.CLASS_NAME, 'ui-datepicker-month'))
            select.select_by_value(str(int(str(BFM).split('-')[1])-1))    
            time.sleep(4)
            select = Select(driver.find_element(By.CLASS_NAME, 'ui-datepicker-year'))
            select.select_by_value(str(BFM).split('-')[0]) 
            time.sleep(4)
            for d in driver.find_elements(By.CLASS_NAME, 'ui-state-default'):
                if d.text == str(int(str(BFM).split('-')[2])):
                    d.click()
                    break    
            time.sleep(4)
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            time.sleep(4)
            if len(driver.find_elements(By.CLASS_NAME, 'preview-file')) > 0:
                IFF = []
                IRF = []
                IFEP = []
                for (ff, rf, en) in zip(driver.find_elements(By.CLASS_NAME, 'preview-file'), driver.find_elements(By.CLASS_NAME, 'enddate')[1:], driver.find_elements(By.CLASS_NAME, 'entity-name')[1:]):
                    if '13F-HR ' in ff.text:
                        IFF.append(ff.text)
                        IRF.append(rf.text)
                        IFEP.append(en.text)     
                if len(IFF) > 0:
                    FF = []
                    RF = []
                    FEP = [] 
                    for (i, j) in zip(IFEP, range(len(IFEP))):
                        if IFEP.count(i) > 1:
                            if FEP.count(i) == 0:
                                FF.append(IFF[j])
                                FEP.append(IFEP[j])
                                DL = []
                                for k in range(len(IRF)):
                                    if IFEP[k] == i and IFEP[k] != '':
                                        DL.append(datetime(int(str(IRF[k]).split('-')[0]), int(str(IRF[k]).split('-')[1]), int(str(IRF[k]).split('-')[2])))
                                RF.append(str(max(DL)).split(' ')[0])         
                        else:
                            FF.append(IFF[j])
                            RF.append(IRF[j])
                            FEP.append(IFEP[j])
                    IMC = []
                    COL1 = []
                    COL4 = []
                    COL5 = [] 
                    for (ff, rf, en) in zip(driver.find_elements(By.CLASS_NAME, 'preview-file'), driver.find_elements(By.CLASS_NAME, 'enddate')[1:], driver.find_elements(By.CLASS_NAME, 'entity-name')[1:]):
                        if '13F-HR ' in ff.text:
                            if rf.text in RF and en.text in FEP:  
                                ff.click()
                                time.sleep(4)
                                driver.find_element(By.ID, 'open-submission').click()
                                time.sleep(4)
                                driver.switch_to.window(driver.window_handles[-1])
                                driver.find_element(By.XPATH, '//*[@id="formDiv"]/div/table/tbody/tr[4]/td[3]/a').click()
                                time.sleep(4)
                                page = driver.page_source
                                soup = BeautifulSoup(page, 'lxml')
                                subsoup = soup.find('table', {'summary': 'Form 13F-NT Header Information'})
                                for r in subsoup.find_all('tr')[3:]:
                                    COL1.append(r.find_all('td', {'class': 'FormData'})[0].text) 
                                    COL4.append(r.find_all('td', {'class': 'FormDataR'})[0].text) 
                                    COL5.append(r.find_all('td', {'class': 'FormDataR'})[1].text) 
                                    IMC.append(Company)
                                driver.close()  
                                driver.switch_to.window(driver.window_handles[0])  
                    time.sleep(1)    
                    df = DataFrame({'INVESTMENT MANAGEMENT COMPANY': IMC, 'SECURITY (ISSUER) NAME': COL1, 'VALUE OF THE POSITION': COL4, 'SHARES OF THE POSITION': COL5})
                    response = HttpResponse(content_type='application/xlsx')
                    response['Content-Disposition'] = f'attachment; filename='+str(Company).replace(' ','_')+'.xlsx'
                    with pd.ExcelWriter(response) as writer:
                        df.to_excel(writer, sheet_name=str(Company).replace(' ','_'))
                    return response
                else:
                    return render(request, "index.html", {'message': 'No results for this company!'})
            else:
                return render(request, "index.html", {'message': 'No results for this company!'})        
        except Exception as e:
            print(e)
            return render(request, "index.html", {'message': 'Something bad happened!'})
    else:    
        return render(request, "index.html")

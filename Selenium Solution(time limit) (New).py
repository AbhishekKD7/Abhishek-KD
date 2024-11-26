from selenium import webdriver
from selenium.webdriver.common.by import By
import html
import time
import requests
from bs4 import BeautifulSoup
from threading import Thread,BoundedSemaphore
import os
import shutil
import traceback
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException,StaleElementReferenceException,TimeoutException,ElementClickInterceptedException
import sys
import trace
import re
import pprint
class thread_with_trace(Thread): 
  def __init__(self, *args, **keywords): 
    Thread.__init__(self, *args, **keywords) 
    self.killed = False
  def start(self): 
    self.__run_backup = self.run 
    self.run = self.__run       
    Thread.start(self) 

  def __run(self): 
    sys.settrace(self.globaltrace) 
    self.__run_backup() 
    self.run = self.__run_backup 
  def globaltrace(self, frame, event, arg): 
    if event == 'call': 
      return self.localtrace 
    else: 
      return None

  def localtrace(self, frame, event, arg): 
    if self.killed: 
      if event == 'line': 
        raise SystemExit() 
    return self.localtrace 
  def kill(self): 
    self.killed = True

    
def states(s,l,RTO1):
    
    while True:
        try:
            # loopterminator[s] = loopterminator[s] + 1
            
            if loopterminator[s] == 1:
              limiter.acquire()
            f = s.split("(")[0]
            f = f.strip()

            if loopterminator[s] == 5:
                print('limit reached for state : '+s)
                statepath = os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f)
                if os.path.isdir(statepath):
                    loopterminator[s] = 1
                    limiter.release()
                    shutil.rmtree(os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f))

                return
            preferences = {"download.default_directory": os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f),  "download.directory_upgrade": True,
                "download.prompt_for_download": False, }
            
            options = webdriver.ChromeOptions()

            """ "safebrowsing.enabled": "false" """

            options.add_experimental_option("prefs", preferences)

            # options.headless = True
            options.add_argument('--headless=new')
            
            driver = webdriver.Chrome(options=options , executable_path= r"C:\Users\USER\Downloads\chromedriver-win32\chromedriver-win32\chromedriver.exe")

            driver.get("https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml")
            driver.maximize_window()

            driver.implicitly_wait(50)
            print('Running for state:',s)
            wait = WebDriverWait(driver, 10 , ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,ElementClickInterceptedException))
            driver.find_element(By.ID, "xaxisVar_label").click()
            time.sleep(1)
            driver.find_element(By.ID, "xaxisVar_6").click()
            time.sleep(2)
            driver.find_element(By.ID, "yaxisVar_label").click()
            time.sleep(1)
            driver.find_element(By.ID, "yaxisVar_4").click()
            time.sleep(1)
            # driver.find_element(By.ID, "selectedYear_label").click()
            # time.sleep(1)
            # driver.find_element(By.ID, "selectedYear_1").click()
            # time.sleep(2)
            driver.find_element(By.ID, "j_idt67").click()
            time.sleep(2)
            driver.find_element(By.ID, "filterLayout-toggler").click()
            time.sleep(2)
            driver.find_element(By.XPATH, '//label[text()="TWO WHEELER(NT)"]').click()
            time.sleep(1)
            # driver.find_element(By.XPATH, '//label[text()="ELECTRIC(BOV)"]').click()
            # time.sleep(1)
            driver.find_element(By.XPATH, '//label[text()="PETROL"]').click()
            time.sleep(1)
            driver.find_element(By.XPATH, '//label[text()="PETROL/ETHANOL"]').click()
            time.sleep(1)            
            mainrefreshbutton = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@ID="j_idt72"]')))
            mainrefreshbutton.click()
            time.sleep(1)
            wait.until(EC.presence_of_element_located((By.XPATH, '//div[@ID="combTablePnl"][@aria-busy="false"]')))

            wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@ID="j_idt72"]')))
            driver.find_element(By.ID, "j_idt37").click()
            time.sleep(1)
            stateselector = wait.until(EC.element_to_be_clickable((By.XPATH, '//li[text()= "'+s+'" ]')))
            
            driver.execute_script("arguments[0].click();", stateselector)
            time.sleep(1)
            
            RTO =  []
            if RTO1 == []:

                
                

                gotitRTO = wait.until(EC.presence_of_element_located((By.ID,"selectedRto_items")))
                
                # gotitRTOlist = gotitRTO.find_elements(By.TAG_NAME,"li")
                gotitRTOlist = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME,"li")))

                HTML = driver.page_source

                soup1 = BeautifulSoup(HTML, 'html.parser')
                RTOSOUP = soup1.find('ul', id="selectedRto_items")

                for i in RTOSOUP.find_all('li'):
        
                  RTO.append(i.text)

                print(len(RTO)," RTO's for state : ",s)    
                # print(RTO)
            else :
                RTO = RTO1

                 
            break
        except Exception as e: 
            print(e)
            print(traceback.format_exc())
            print('above exception for state : '+ s)
            # limiter.release()
            loopterminator[s] = loopterminator[s] + 1
            driver.quit()
            continue

    def rtotest(driver,j,l,f,s1):
        while True:
            try:
                
                wait2 = WebDriverWait(driver, 10 , ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,TimeoutException,ElementClickInterceptedException))
                
                if s1 == 1:
                    time.sleep(5)
                file_path = os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f)+r'\reportTable.xlsx'
                if os.path.isfile(file_path):
                    os.remove(file_path)
                                
                driver.find_element(By.XPATH, '//label[@id = "selectedRto_label"]').click()
                time.sleep(0.5)

                wait2.until(EC.presence_of_element_located((By.XPATH, '//li[text()= "'+j+'" ]'))).click()
                time.sleep(0.5)
                checkrto = wait2.until(EC.visibility_of_element_located((By.ID,"selectedRto_label")))
                # print(j,' ',checkrto.text)
                O = j.rindex('(')
                M = j[0:O]
                M = (M[M.rindex(' '):O])
                M = M.strip()
                if(M in checkrto.text):

                    mainrefreshbutton2 = driver.find_element(By.XPATH, '//button[@ID="j_idt72"]')
                    # driver.execute_script("arguments[0].click();", mainrefreshbutton2)
                    while True:
                        wait3 = WebDriverWait(driver, 3 , ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,TimeoutException,ElementClickInterceptedException))
                        
                        # time.sleep(0.5)
                        try :
                          wait3.until(EC.presence_of_element_located((By.XPATH, '//div[@ID="combTablePnl"][@aria-busy="false"]')))
                          mainrefreshbutton2.click()
                          wait3.until(EC.presence_of_element_located((By.XPATH, '//div[@ID="combTablePnl"][@aria-busy="true"]')))
                          break
                        except TimeoutException:  
                          print('Timeout for RTO ' + j + ' of state: '+ s)
                          pass
   
                    wait2.until(EC.presence_of_element_located((By.XPATH, '//div[@ID="combTablePnl"][@aria-busy="false"]')))

                    time.sleep(0.5)
                    downloadbutton = wait2.until(EC.element_to_be_clickable((By.ID , "groupingTable:xls")))

                    downloadbutton.click()
                else:
                    print('rto label is incorrect for RTO : '+ j+' state : '+ s)
                    break  

                

                filepresent = False

                while filepresent == False:
                    filepath2 = os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f)+r'\reportTable.xlsx'
                    if os.path.isfile(filepath2):
                        os.rename(os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f)+r'\reportTable.xlsx',os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f,M)+'.xlsx')
                        filepresent = True
                        # print('rto : '+ j + ' renamed for state : '+s)
                    time.sleep(1)    
                break

            except Exception as e: 
                print(e)
                print(traceback.format_exc())
                print('above exception for RTO ' + j + ' of state: '+ s)
                break
        
             
    for j in RTO:
        s1 = 0
        max_execution_time = 120
        p = thread_with_trace(target=rtotest, args=(driver,j,l,f,s1))
        p.start()
        p.join(max_execution_time)
        if p.is_alive():
            s1 = 1
            print("Exceeded Execution Time for RTO : ",j," in state ",s)
            p.kill()
            time.sleep(5)
            break
        
        
        
       
    print("state ",s," got out of rto loop")   
    
    missingrto = []
    rtoloadedlist = []

    statepath = os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f)
    time.sleep(5)
    if os.path.isdir(statepath):
        
        rtofilenames = os.scandir(os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f))
        for rtoname in rtofilenames:
          rtoloadedlist.append(rtoname.name.replace('.xlsx',''))
        rtoclean = []  
        for k in RTO:
            H = k.rindex('(')
            M = k[0:H]
            M = (M[M.rindex(' '):H])
            M = M.strip()
            if M not in rtoloadedlist:
                missingrto.append(k)

        for N in rtoloadedlist:
            if ('reportTable' in N):
                print('reporttable name disprency for: ',s)
                print(rtoloadedlist)
                # limiter.release()
                loopterminator[s] = loopterminator[s] + 1
                if len(rtoloadedlist)>len(RTO):
                   print('extra reporttable file disprency for: ',s)
                   file_path_reporttable = os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f)+r'\reportTable.xlsx'
                   if os.path.isfile(file_path_reporttable):
                     os.remove(file_path_reporttable)
                   limiter.release()
                   driver.quit() 
                   return "State "+s+ " completed"
                   
                driver.quit() 
                shutil.rmtree(os.path.join(r'C:\Users\USER\Documents\Vaahan',l,f))
  
                states(s,l,[])
                return
            
        if len(missingrto) != 0:
            print('Going to missing rto: ',s)
            # limiter.release()
            loopterminator[s] = loopterminator[s] + 1
            print('loopterminator value of '+s +' :',loopterminator[s])
            driver.quit() 
            
            states(s,l,missingrto)
            return
        if len(missingrto) == 0:
            loadedstatelist.append(f)
    else   :
        # limiter.release()
        print('no folder')
        loopterminator[s] = loopterminator[s] + 1
        print('loopterminator value of '+s +' :',loopterminator[s])
        
        driver.quit()
        states(s,l,[])

    limiter.release()
    driver.quit()
    return "State "+s+ " completed"

if __name__ == "__main__":
    URL = 'https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml'
    
    print('Program start at : '+ str(time.time()))

    reqs = requests.get(URL )#,verify= r'C:\Users\USER\Documents\Vaahan\cert\cacertcopy.pem')
    soup = BeautifulSoup(reqs.text, 'html.parser')
    state = []
    statesorted = []
    global loopterminator 
    loopterminator = {}
    global loadedstatelist
    loadedstatelist = []

    parentstatecode = soup.find('select', id="j_idt37_input")

    for i in parentstatecode.find_all('option'):
        
        state.append(i.text)
        loopterminator.update({i.text : 1})

    l = 'Vahan' + str(round(time.time()))
    
    substrings = [int(x) for string in state for x in re.findall(r'\d+', string)]
    substrings.sort(reverse=True)
    statesorted = [string for x in substrings for string in state if int(re.findall(r'\d+', string)[0]) == x]
    limiter = BoundedSemaphore(1)
    thread = []
    
    statesorted = [i for n, i in enumerate(statesorted) if i not in statesorted[:n]] 
    for s in statesorted:
        
        t = thread_with_trace(target=states, args=(s,l,[]))
        thread.append(t)
        t.start()
       
    
    for o in thread:
       o.join(timeout = 3000) 
       if o.is_alive():
            
            
            o.kill()
            

       print('terminating state thread at : '+ str(time.time()))

    print('This shouldnt start till all states are over')
    
    loadedstates = []
    
    missingstate = []
    loadedstates2 =  os.scandir(os.path.join(r'C:\Users\USER\Documents\Vaahan',l))
    for loadedstates1 in loadedstates2:
        loadedstates.append(loadedstates1.name)

    for s in statesorted:
        f = s.split("(")[0]
        f = f.strip()
        
        if f not in loadedstates:
          missingstate.append(s)
         
    thread2 = []        
    for k in missingstate:
        print('missing state start at : '+ str(time.time()))
        t = Thread(target=states, args=(k,l,[]))
        t.start()        
        thread2.append(t)

    for v in thread2   :
        v.join()
        print('missing state completed :'+ str(time.time()))

    print('program finishing at:'+ str(time.time()))
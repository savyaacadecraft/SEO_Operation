import pandas as pd
import requests
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException

from os import mkdir, path, listdir
from time import sleep

def getValidDomain(Data):
    valid_domain = [".com", ".org", ".edu", ".in", ".edu.au", ".org.uk", ".co", ".net", ".info", ".co.uk", ".co", ".ca", ".com.eu", ".buzz", ".eu", ".io", ".se", ".biz", ".to", ".fm"]
    
    for url in Data:
        appropriate = ""
        for i in valid_domain:
            if i in url[1]:
                appropriate = i if len(i) > len(appropriate) else appropriate
        if len(appropriate): 
            url[1] = url[1].split(appropriate)[0] + appropriate

def getAge(engine, i) -> str:
    # global Age_dict
    # # Age is i[5]
    
    engine.get(f"https://www.whatsmydns.net/domain-age?q={i}")
    sleep(2)
    
    try:
        age = WebDriverWait(engine, 30).until(EC.visibility_of_element_located((By.ID, "years"))).text
    except Exception as E:
        age = ""

    return age

def getLocation(engine, i) -> str:
    # global Country_dict
    # Location is i[6]
    
    engine.get(f"https://www.whatsmydns.net/whois?q={i}")
    sleep(2)
    try:
        
        details = (WebDriverWait(engine, 30).until(EC.visibility_of_element_located((By.ID, "whois"))).text).split("\n")
        for d in details:
            if "Country" in d:
                country = d.split(" ")[-1]
            else:
                country = ""

    except Exception as E:
        country = ""

    return country

def getDRandOrganicTraffic(engine, i) -> str:
    
    engine.get(f"https://www.semrush.com/analytics/overview/?q={i}&protocol=https&searchType=domain")
    sleep(5)

    # DR is i[3]
    
    try:
        DR = WebDriverWait(engine, 10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/main/div/div/div[2]/div[2]/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/a/span"))).text
    except Exception as E:
        DR = ""
    

    # Organic Traffic is i[4]
    
    try:
        Organic_Traffic = WebDriverWait(engine, 10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/main/div/div/div[2]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[1]/a/span"))).text
    except Exception as E:
        Organic_Traffic = ""
    

    return (DR, Organic_Traffic)
        
def websiteExist(URL):

    try:
        requests.get(URL, timeout=5)
        return True
    except Exception as E:
        return False

if __name__ == "__main__":
    

    filename = 'Poonam_Editorial Acadecraft.xlsx'
    df = pd.read_excel(filename)

    cols = df.columns.tolist()
    df = df.values.tolist()
    
    
    getValidDomain(df)
    for i in range(len(df)):
        df[i][0] = df[i][0].strftime('%Y-%m-%d')
        df[i][1] = str(df[i][1])
        df[i][2] = str(df[i][2])
        df[i][3] = str(df[i][3])
        df[i][4] = str(df[i][4])
        df[i][5] = str(df[i][5])
        df[i][6] = str(df[i][6])
        df[i][7] = str(df[i][7])
        df[i][8] = str(df[i][8])
        df[i][9] = str(df[i][9])
        df[i][10] = str(df[i][10])
        df[i][11] = str(df[i][11])
        df[i][12] = str(df[i][12])
        df[i][13] = str(df[i][13])
        df[i][14] = str(df[i][14])

        
        df[i] = "|".join(df[i])
        df[i].replace("\n","")
    

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9999")

    service = Service(ChromeDriverManager().install())

    engine = webdriver.Chrome(service=service, options=options)
    # engine.maximize_window()
    print("Operation Started-----")

    mainData = list()
    try:
        for data in df:
            i = data.split("|")

            if i[3] == "nan" or i[4] == "nan":
                i[3], i[4] = getDRandOrganicTraffic(engine, i[1])
            
            # if i[5] == "nan" or (not len(i[5])):
            #     print("i")
            #     try:
            #         WebDriverWait(engine, 2).until(EC.visibility_of_element_located(By.XPATH, '//h2[text()="Checking if the site connection is secure"]'))
            #         input("Enter: ")
            #     except Exception as E:
            #         pass

            #     i[5] = getAge(engine, i[1])
            

            # if i[6] == "nan" or (not len(i[6])):
            #     print("o")
            #     try:
            #         WebDriverWait(engine, 2).until(EC.visibility_of_element_located(By.XPATH, '//h2[text()="Checking if the site connection is secure"]'))
            #         input("Enter: ")
            #     except Exception as E:
            #         pass

            #     i[6] = getLocation(engine, i[1])
            
            
            
            mainData.append(i)
    except Exception as E:
        mainData = pd.DataFrame(mainData, columns=cols)
        mainData.to_excel("Output_Data_1.xlsx", index=False)
    
    except KeyboardInterrupt as KE:
        mainData = pd.DataFrame(mainData, columns=cols)
        mainData.to_excel("Output_Data_1.xlsx", index=False)

    mainData = pd.DataFrame(mainData, columns=cols)
    mainData.to_excel("Output_Data_1.xlsx", index=False)

    engine.quit()
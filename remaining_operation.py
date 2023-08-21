import pandas as pd
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



if __name__ == "__main__":

    Data_file = "Output_Data.xlsx"
    df = pd.read_excel(Data_file)

    cols = df.columns.tolist()
    df = df.values.tolist()


    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9999")

    service = Service(ChromeDriverManager().install())

    engine = webdriver.Chrome(service=service, options=options)
    # engine.maximize_window()
    print("Operation Started-----")

    mainData = list()
    for i in df:
        if not len(i[4]):
            i[4] = getAge(engine, i[1])
        
        if not len(i[5]):
            i[5] = getLocation(engine, i[1])

    df = pd.DataFrame(df, columns=cols)
    df.to_excel("Output_Data_1.xlsx", index=False)

    engine.quit()
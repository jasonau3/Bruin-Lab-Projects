from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import pandas as pd
from datetime import date
from getLogIn import getPassword, getUsername   # hidden file

# open the website
def setUp(url):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(url)
    print(driver.title)
    return driver

# log in
def logIn(email, password):
    emailInput = driver.find_element(by=By.XPATH, value="//*[@id=\"MainContent_txtEmail\"]")
    emailInput.send_keys(email)

    pwdInput = driver.find_element(by=By.XPATH, value="//*[@id=\"MainContent_txtPassword\"]")
    pwdInput.send_keys(password)

    loginButton = driver.find_element(by=By.XPATH, value="//*[@id=\"MainContent_Login\"]")
    loginButton.send_keys(Keys.RETURN)

# get ids from excel file
def get_all_ids(idFile):
    # set up
    workbook = load_workbook(filename=idFile)
    sheet = workbook.active

    # get ids
    all_ids = []
    for col_A in sheet["B"]:
        if col_A.value:
            all_ids.append(col_A.value)

    # first id is "ID", so remove that
    all_ids.pop(0)

    # finally return the list
    return(all_ids)

# scrap for id parameter
def getInfo(id):
    # each id has its own url in the format (ex "r=R300")
    driver.get(f"https://www.epicore.ualberta.ca/IsletCore/Default?r={id}")

    # get basic data
    ageValue = float(driver.find_element(by = By.XPATH, value="//*[@id=\"MainContent_PanelDonorInfo\"]/div/table/tbody/tr[1]/td[4]").text)
    bmiValue = float(driver.find_element(by = By.XPATH, value="//*[@id=\"MainContent_PanelDonorInfo\"]/div/table/tbody/tr[1]/td[6]").text)
    genderValue = driver.find_element(by = By.XPATH, value="//*[@id=\"MainContent_PanelDonorInfo\"]/div/table/tbody/tr[2]/td[4]").text[0] # get first letter
    HbA1cRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"MainContent_PanelDonorInfo\"]/div/table/tbody/tr[2]/td[6]").text
    HbA1cValue = "NA" if HbA1cRaw == "no data" or HbA1cRaw == "" else float(HbA1cRaw)
    diabetesRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"MainContent_PanelDonorInfo\"]/div/table/tbody/tr[2]/td[8]/ul/li").text
    diabetesValue = "N" if diabetesRaw == "None" else "Y"

    # get GSIS data
    functionAndFeedbackTab = driver.find_element(by = By.XPATH, value="//*[@id=\"info\"]/div/div[1]/ul/li[3]/a").click()
    # data is in a iframe, so we need to use switchTo
    iframe = driver.find_element(By.CSS_SELECTOR, "#quality > div > iframe")
    driver.switch_to.frame(iframe)

    # 1 mM/10 mM Paired Data
    oneVsTen_oneMMRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[3]/td[2]").text 
    oneVsTen_oneMMvalue = "NA" if oneVsTen_oneMMRaw == "" else float(oneVsTen_oneMMRaw)
    oneVsTen_tenMMRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[5]/td[2]").text
    oneVsTen_tenMMvalue = "NA" if oneVsTen_tenMMRaw == "" else float(oneVsTen_tenMMRaw)
    oneVsTen_simIndexRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[7]/td[2]").text
    oneVsTen_simIndexValue = "NA" if oneVsTen_simIndexRaw == "" else float(oneVsTen_simIndexRaw)

    # 1 mM/16.7 mM Paired Data
    oneVsSixteen_oneMMRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[10]/td[2]").text
    oneVsSixteen_oneMMvalue = "NA" if oneVsSixteen_oneMMRaw == "" else float(oneVsSixteen_oneMMRaw)
    oneVsSixteen_sixteenMMRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[12]/td[2]").text
    oneVsSixteen_sixteenMMvalue = "NA" if oneVsSixteen_sixteenMMRaw == "" else float(oneVsSixteen_sixteenMMRaw)
    oneVsSixteen_simIndexRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[14]/td[2]").text
    oneVsSixteen_simIndexValue = "NA" if oneVsSixteen_simIndexRaw == "" else float(oneVsSixteen_simIndexRaw)
    
    # 2.8 mM/16.7 mM Paired Data
    twoVsSixteen_twoMMRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[17]/td[2]").text
    twoVsSixteen_twoMMvalue = "NA" if twoVsSixteen_twoMMRaw == "" else float(twoVsSixteen_twoMMRaw)
    twoVsSixteen_sixteenMMRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[19]/td[2]").text
    twoVsSixteen_sixteenMMvalue = "NA" if twoVsSixteen_sixteenMMRaw == "" else float(twoVsSixteen_sixteenMMRaw)
    twoVsSixteen_simIndexRaw = driver.find_element(by = By.XPATH, value="//*[@id=\"form1\"]/div[3]/table/tbody/tr[21]/td[2]").text
    twoVsSixteen_simIndexValue = "NA" if twoVsSixteen_simIndexRaw=="" else float(twoVsSixteen_simIndexRaw)

    # escape iframe
    driver.switch_to.default_content()

    # open isolation info
    isolationTab = driver.find_element(by = By.CSS_SELECTOR, value="#info > div > div.col-md-4 > ul > li:nth-child(2) > a").click()
    # data is in a iframe, so we need to use switchTo
    iframe = driver.find_element(By.CSS_SELECTOR, "#isolation > div > iframe")
    driver.switch_to.frame(iframe)

    # get isolation data
    coldIschemiaTime = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#form1 > div:nth-child(11) > table > tbody > tr:nth-child(2) > td:nth-child(2)"))).text
    coldIschemiaTime = "NA" if coldIschemiaTime == "" else float(coldIschemiaTime) 
    purity = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#form1 > div:nth-child(11) > table > tbody > tr:nth-child(24) > td:nth-child(2)"))).text
    purity = "NA" if purity == "" else float(purity)
    trapped = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#form1 > div:nth-child(11) > table > tbody > tr:nth-child(25) > td:nth-child(2)"))).text
    trapped = "NA" if trapped == "" else float(trapped)

    # escape iframe
    driver.switch_to.default_content()

    # open sample tab
    isolationTab = driver.find_element(by = By.CSS_SELECTOR, value="#info > div > div.col-md-4 > ul > li:nth-child(6) > a").click()
    # data is in a iframe, so we need to use switchTo
    iframe = driver.find_element(By.CSS_SELECTOR, "#sample > div > iframe")
    driver.switch_to.frame(iframe)

    # get sample data
    preDisCultureTime = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#form1 > div:nth-child(11) > table > tbody > tr:nth-child(3) > td:nth-child(2)"))).text
    preDisCultureTime = "NA" if preDisCultureTime == "" else float(preDisCultureTime) 
    
    # use pandas to make a dataframe
    idData = pd.DataFrame({
        'Id': id,
        'Sex': genderValue, 
        'Age': ageValue, 
        'BMI': bmiValue, 
        '%Hba1c': HbA1cValue, 
        'Diabetes Status': diabetesValue, 
        '1mM/10mM 1mM Value (pg/ml)': oneVsTen_oneMMvalue,
        '1mM/10mM 10mM Value (pg/ml)': oneVsTen_tenMMvalue,
        '1mM/10mM Stimulation index': oneVsTen_simIndexValue,
        '1mM/16.7mM 1mM Value (pg/ml)': oneVsSixteen_oneMMvalue,
        '1mM/16.7mM 16.7mM Value (pg/ml)': oneVsSixteen_sixteenMMvalue,
        '1mM/16.7mM Stimulation index': oneVsSixteen_simIndexValue,
        '2.8mM/16.7mM 1mM Value (pg/ml)': twoVsSixteen_twoMMvalue,
        '2.8mM/16.7mM 16.7mM Value (pg/ml)': twoVsSixteen_sixteenMMvalue,
        '2.8mM/16.7mM Stimulation index': twoVsSixteen_simIndexValue,
        'Cold ischemia time (h)': coldIschemiaTime,
        'Purity (%)': purity,
        'Trapped (%)': trapped,
        'Pre-distribution culture time (hours)': preDisCultureTime
        }, index =[0])
    return idData


# Program set up
driver=setUp("https://www.epicore.ualberta.ca/IsletCore/Login")
logIn(email=getUsername(), password=getPassword())

# Get ids
idFile = "../ADI Data/JA_2022-05-06_ADI ISLET DATA.xlsx"
allIds = get_all_ids(idFile)

# Get id info and export it
exportData = pd.DataFrame()
for id in allIds:
    idData = getInfo(id)
    exportData = pd.concat([exportData, idData]).reset_index(drop = True)
    print(idData)
fileName = "JA_" + str(date.today()) + "_ADI ISLET DATA.xlsx"
exportData.to_excel(fileName, sheet_name='Islet Data')
print("Successfully exported to the excel file: " + fileName)
driver.quit()

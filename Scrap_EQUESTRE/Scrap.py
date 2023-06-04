from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import pandas as pd
import pyautogui

Nom = []
col1 = "Titre"
Adresse = []
col2 = "Adresse"
Numero = []
col3 = "Numéro"
Email = []
col4 = "Email"
Lien = []
col5 = "URL spécifiée"

Error = 1514
Site = f"https://www.xxx-xxx.com/recherche?perPage=5&page={Error}"

# Ouvrir le fichier chrome.exe

options = webdriver.ChromeOptions()

driver = webdriver.Chrome("/Users/wglint/Downloads/chromedriver", options=options)
driver.maximize_window()
driver.get(Site)

# Cliquer sur le bouton "ANNNUAIRE
sleep(2)

try:
    popup = driver.find_element_by_xpath("/html/body/div[5]/div/div/div[2]/button[2]")
    popup.click()
    sleep(0.5)
    cookie = driver.find_element_by_xpath("/html/body/div[1]/div/div/footer/section/div/div/div[2]/button[2]")
    cookie.click()
    sleep(0.5)
except:
    print("Pas de popup")


pyautogui.keyDown('command')
for i in range(6):
    pyautogui.press('+')
    sleep(1)
pyautogui.keyUp('command')


sleep(5)
# Commence à prendre les informations
Nom_Int = 0
Page = 1
while True:
    try:
        print("GO")
        for i in range(1,6):
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[1]/div/div/div/header/section/div/div/div/div/h1")))
            sleep(1)
            for k in range(4):
                body = driver.find_element_by_xpath("/html/body")
                body.send_keys(Keys.ARROW_DOWN)
                sleep(0.2)

            while True:
                try:
                    fen = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                f'/html/body/div[1]/div/div/div/div[2]/div/div[3]/div[2]/div/div[{i}]')))
                    Nom_inter = driver.find_elements_by_xpath(f'//*[@id="showInfo"]/div/div[3]/div[2]/div/div[{i}]/div/div[1]/div[2]/a/h2/strong')[0].text
                    if (len(Nom_inter) < 5):
                        print(Nom_inter)
                        Nom.append("ERREUR")
                    else :
                        Nom.append(Nom_inter)
                        print(Nom_inter)
                    fen.click()
                    break
                except Exception as e:
                    print(e)
                    body.send_keys(Keys.ARROW_DOWN)
                    sleep(0.5)
                    continue


            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[1]/div/div/footer")))
            sleep(0.7)

            for k in range(7):
                body = driver.find_element_by_xpath("/html/body")
                body.send_keys(Keys.ARROW_DOWN)
                sleep(0.1)

            Link = 1
            try:
                Adresse.append(driver.find_elements_by_xpath('//*[@id="contact"]/div/div[2]/div[1]/span')[0].text)
                Link += 1
            except:
                Adresse.append("Donnée Manquante")

            for i in range(1,5):
                try:
                    num = driver.find_elements_by_xpath(f'//*[@id="contact"]/div/div[2]/div[{i}]/span')[0].text
                    if (num[0] == "+") or (num[0] == "0"):
                        Numero.append(num)
                        Link += 1
                        break
                    elif (i == 4):
                        Numero.append("Donnée manquante")
                        break
                except:
                    if (i == 4):
                       Numero.append("Donnée Manquante")
                       break
                    continue

            for i in range(1,5):
                try:
                    if (driver.find_elements_by_xpath(f'//*[@id="contact"]/div/div[2]/div[{i}]/div')[0].text == "Affichez l'email"):
                        Cache = driver.find_element_by_xpath(f'//*[@id="contact"]/div/div[2]/div[{i}]/div')
                        Cache.click()
                        Email.append(driver.find_elements_by_xpath(f'//*[@id="contact"]/div/div[2]/div[{i}]/div')[0].text)
                        Link += 1
                        break
                    elif (i == 4):
                        Email.append("Donnée manquante")
                        break
                except:
                    if (i == 4):
                       Email.append("Donnée Manquante")
                       break
                    continue

            try:
                Lien.append(driver.find_elements_by_xpath(f'//*[@id="contact"]/div/div[2]/div[{Link}]/div')[0].text)
            except:
                Lien.append("Donnée Manquante")

            driver.back()
            sleep(1.5)

        if (Page == 1):
            suiv = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[2]/div/div[5]/div/ul/li[6]/a")
            suiv.click()
            Page = 0
            sleep(3)
        else :
            suiv = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[2]/div/div[5]/div/ul/li[7]/a")
            suiv.click()
            sleep(3)

        Error += 1
        print(len(Nom))
        print(len(Adresse))
        print(len(Numero))
        print(len(Email))
        print(len(Lien))
        Data = pd.DataFrame(
            {col1: Nom, col2: Adresse, col3: Numero, col4: Email, col5: Lien})
        Data.to_excel('SCRAP.xlsx', index=False)
    except Exception as e:
        driver.close()
        print(e)
        sleep(10)
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome("/Users/wglint/Downloads/chromedriver", options=options)
        driver.maximize_window()
        driver.get(f"https://www.cheval-reference.com/recherche?perPage=5&page={Error + 1}")
        sleep(10)

        pyautogui.keyDown('command')
        for i in range(8):
            pyautogui.press('+')
            sleep(1)
        pyautogui.keyUp('command')

        try:
            popup = driver.find_element_by_xpath("/html/body/div[5]/div/div/div[2]/button[2]")
            popup.click()
            sleep(0.5)
            cookie = driver.find_element_by_xpath("/html/body/div[1]/div/div/footer/section/div/div/div[2]/button[2]")
            cookie.click()
            sleep(0.5)
        except:
            print("Pas de popup")

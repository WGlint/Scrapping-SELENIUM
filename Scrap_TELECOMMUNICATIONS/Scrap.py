from selenium import webdriver
from time import sleep
import pandas as pd

def Press_cloture(X):
    global Cloture_C
    try:
        Button_cloture = driver.find_element_by_xpath('/html/body/div[{}]/div[2]/div[1]/ul/li[6]/a'.format(X))
        Button_clotures = driver.find_elements_by_xpath('/html/body/div[{}]/div[2]/div[1]/ul/li[6]/a'.format(X))
        Brut = Button_clotures[0].text
        Net = []
        for i in range(0, 7):
            Net.append(Brut[i])
        Net_final = ''.join(Net)
        if "Clôture" == Net_final:
            Button_cloture.click();
            sleep(0.3)
            Cloture_C = 6
        else:
            Button_cloture = driver.find_element_by_xpath('/html/body/div[{}]/div[2]/div[1]/ul/li[7]/a'.format(X))
            Button_cloture.click();
            sleep(0.3)
            Cloture_C = 7
    except:
        pass

def Press_button(X):
    try:
        Button = driver.find_element_by_xpath('/html/body/div[{}]/div[3]/div/button[1]'.format(X))
        Button.click();
        sleep(4)
    except:
        pass


def Press_exit(X):
    try:
        Exit = driver.find_element_by_xpath('/html/body/div[{}]/div[3]/div/button/span'.format(X))
        Exit.click();
        sleep(1)
    except:
        pass


def Press_Rapport(X):
    global Cloture_C
    try:
        Button_Rapport = driver.find_element_by_xpath(
            '/html/body/div[{}]/div[2]/div[1]/div/div[{}]/div[5]/div/div[1]'.format(X, Cloture_C))
        Button_Rapport.click();
        sleep(0.5)
    except:
        pass


def Press_Raccordement(X):
    global Cloture_C
    try:
        Button_Raccordement = driver.find_element_by_xpath(
            '/html/body/div[{}]/div[2]/div[1]/div/div[{}]/div[5]/div/div[2]/div/div[4]/div/div[1]/div'.format(X,
                                                                                                              Cloture_C))
        Button_Raccordement.click();
        sleep(0.5)
    except:
        pass


def Add_Batiment(X):
    global Cloture_C
    try:
        Batiment_xpath = driver.find_elements_by_xpath(
            '/html/body/div[{}]/div[2]/div[1]/div/div[{}]/div[5]/div/div[2]/div/div[4]/div/div[2]/div[2]/div[2]/div'.format(
                X, Cloture_C))
        Batiment.append(Batiment_xpath[0].text)
    except:
        pass


def Add_Activité(X):
    global Cloture_C
    try:
        RECO_xpath = driver.find_elements_by_xpath(
            '/html/body/div[{}]/div[2]/div[1]/div/div[{}]/div[5]/div/div[2]/div/div[4]/div/div[2]/div[1]/div[2]/div'.format(
                X, Cloture_C))
        RECO.append(RECO_xpath[0].text)
    except:
        pass


# -----------------------------------------------------------------------------------
# Liste EXCEL
Nom_et_Prenom = []
Numéro_as = []
Date_J = []
Durée = []
Statut = []
Erreur = []
Type = []
Marque = []
Batiment = []
Activité = []
RECO = []

# -----------------------------------------------------------------------------------
# Constante :
# n => Sample iteration for
n = 0
c = 0
Erreur_clique = 0
# S_x => Les iterations pour les semaines (date)
S_2 = 1;
S_3 = 1;
S_4 = 1;
S_5 = 1;
S_6 = 1
# Le temps de pause 
Time = 2
Cloture_C = 0
# Liste d'information programme :
Id_List = [4922, 5856, 6066, 4228, 5103, 4926, 6178, 3110, 5175, 5503, 5185, 5465, 4457]
# -----------------------------------------------------------------------------------

# Dictionnaire
Id = {6066: "xxx", 4922: "xxx", 4228: "xxx", 5103: "xxx",
      5856: "xxx", 6178: "xxx", 4926: "xxx", 3110: "xxx",
      5175: "xxx", 5503: "xxx", 5185: "xxx", 5465: "xxx",
      4457: "xxx"}
Mois_Dict = {"Janvier": 1, "Fevrier": 2, "Mars": 3, "Avril": 4, "Mai": 5, "Juin": 6, "Juillet": 7, "Août": 8,
             "Septembre": 9, "Octobre": 10, "Novembre": 11, "Decembre": 12}
Jour_Dict = {"Janvier": 31, "Fevrier": 28, "Mars": 31, "Avril": 30, "Mai": 31, "Juin": 30, "Juillet": 31, "Août": 31,
             "Septembre": 30, "Octobre": 31, "Novembre": 30, "Decembre": 31}
Dict_Jour = {1: "Lundi", 2: "Mardi", 3: "Mercredi", 4: "Jeudi", 5: "Vendredi", 6: "Samedi", 7: "Dimanche"}

# Mois et année à demander
Année = 2022
Mois = "Juillet"

# Execution Plugin "Chromedriver.exe" et lancement du site
driver = webdriver.Chrome("/Users/wglint/Downloads/chromedriver")
driver.get("https://xxx.xx.fr/xxx/xxx/")

# Entrer les logins CIRCET
sleep(5)
Login = "xxx"
Password = "xxx"
Login_enter = driver.find_element_by_xpath("/html/body/div[4]/div[1]/form/input[1]")
Login_enter.send_keys(Login)
Password_enter = driver.find_element_by_xpath("/html/body/div[4]/div[1]/form/input[2]")
Password_enter.send_keys(Password)
ENTER = driver.find_element_by_xpath("/html/body/div[4]/div[1]/form/input[4]")
ENTER.click()

# Entrer dans le planning
Planning = driver.find_element_by_id("menu_planning")
Planning.click()

# Etape lancement du programme
sleep(5)
Mois_Planning = driver.find_elements_by_xpath('//*[@id="datepicker"]/div/div/div/span[1]')
sleep(0.5)
Année_planning = driver.find_elements_by_xpath('//*[@id="datepicker"]/div/div/div/span[2]')
sleep(1)

Iteration_Mois = Mois_Dict[Mois_Planning[0].text] - Mois_Dict[Mois]
Iteration_Année = int(Année_planning[0].text) - Année

try:
    for Jour in range(0, Jour_Dict[Mois]):
        if Iteration_Année == 0:
            for i in range(0, Iteration_Mois):
                Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/div/a[1]/span')
                Button.click()
                sleep(0.1)
        else:
            for i in range(0, Iteration_Mois + 12 * Iteration_Année):
                Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/div/a[1]/span')
                Button.click()
                sleep(0.1)

        sleep(Time)
        if n == 0:
            Nom_Jour = driver.find_elements_by_xpath('//*[@id="datepicker"]/div/table/tbody/tr[1]/td[7]/a')
            Nomber_Jour = 8 - int(Nom_Jour[0].text)
            Date = Nomber_Jour
            n += 1

        # Premier Semaine
        if Date <= 7:
            Button = driver.find_element_by_xpath(
                '//*[@id="datepicker"]/div/table/tbody/tr[1]/td[{}]'.format(Nomber_Jour))
            Button.click()
            Date += 1
            Nomber_Jour = Nomber_Jour + 1
            sleep(Time)

        # Deuxième Semaine
        elif 8 <= Date <= 14:
            Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/table/tbody/tr[2]/td[{}]'.format(S_2))
            Button.click()
            S_2 += 1
            Date += 1
            sleep(Time)

        # Trosième Semaine
        elif 15 <= Date <= 21:
            Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/table/tbody/tr[3]/td[{}]'.format(S_3))
            Button.click()
            S_3 += 1
            Date += 1
            sleep(Time)

        # Quatrième Semaine
        elif 22 <= Date <= 28:
            Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/table/tbody/tr[4]/td[{}]'.format(S_4))
            Button.click()
            S_4 += 1
            Date += 1
            sleep(Time)

        # Cinquième Semaine
        elif 29 <= Date <= 35:
            Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/table/tbody/tr[5]/td[{}]'.format(S_5))
            Button.click()
            S_5 += 1
            Date += 1
            sleep(Time)

        # Sixième Semaine
        elif 36 <= Date <= 42:
            Button = driver.find_element_by_xpath('//*[@id="datepicker"]/div/table/tbody/tr[6]/td[{}]'.format(S_6))
            Button.click()
            S_6 += 1
            Date += 1
            sleep(Time)

        # Début du traitement CENTRE_EST et BOURGOGNE
        for n in range(0, 2):
            Id_valide = []
            Inter_valide = []
            Opération = -1
            Inter = driver.find_elements_by_xpath('//*[@id="inter"]/b');
            Inter_text = Inter[0].text
            Intervention = Inter_text[0]
            if n == 1:
                CENTRE_EST = driver.find_element_by_xpath('//*[@id="region_id"]/option[3]')
                CENTRE_EST.click()
            else:
                BOURGOGNE = driver.find_element_by_xpath('//*[@id="region_id"]/option[1]')
                BOURGOGNE.click()
            sleep(Time)

            # Création d'une liste des ID par technicien
            for ID in Id_List:
                try:
                    Mission = driver.find_element_by_xpath('//*[@id="g{}"]/div[1]'.format(ID))
                    Id_valide.append(ID);
                    sleep(0.3)
                except:
                    pass

            # Création d'une liste pour le nombre d'opération par jour par technicien
            for ID_valide in Id_valide:
                Compteur = 0
                for i in range(1, 100):
                    try:
                        Mission = driver.find_element_by_xpath('//*[@id="g{}"]/div[{}]'.format(ID_valide, i));
                        sleep(0.2)
                        Compteur = Compteur + 1
                    except:
                        Inter_valide.append(Compteur)
                        break

            sleep(1)
            for ID in range(1, len(Id_valide) + 1):
                Opération += 1
                for Intervention_final in range(1, Inter_valide[Opération] + 1):
                    Nom_et_Prenom.append(Id[Id_valide[Opération]])

                    while True:
                        try:
                            sleep(2)
                            Mission_final = driver.find_element_by_xpath(
                                '/html/body/div[4]/div[1]/div[8]/div/table/tbody/tr[{}]/td[2]/div[{}]/div/span'.format(
                                    ID, Intervention_final))
                            Mission_final.click();
                            sleep(2)
                            break
                        except:
                            print("")
                            print("")
                            print("--- ERREUR SUR LA MISSION SELECTIONNE ---")
                            print(Numéro_as[-1])
                            continue

                    for i in range(5, 30):
                        Press_button(i)

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            AS = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[1]')
                            if Boucle_While == 100:
                                Numéro_as.append("ERREUR")
                                print("")
                                print("Try first click...")
                                print("")
                                break
                            Numéro_as.append(int(AS[0].text));
                            sleep(0.2)
                            break
                        except:
                            continue

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            Date_xpath = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[8]')
                            if Boucle_While == 15:
                                Date_J.append("ERREUR")
                                break
                            Date_new = Date_xpath[0].text
                            Date_J.append(Date_new);
                            sleep(0.2)
                            break
                        except:
                            continue

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            Temps = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[10]')
                            if Boucle_While == 15:
                                Durée.append("ERREUR")
                                break
                            Durée.append(int(Temps[0].text));
                            sleep(0.2)
                            break
                        except:
                            continue

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            Activité_FTT = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[4]')
                            if Boucle_While == 15:
                                Activité.append("ERREUR")
                                break
                            Activité.append(Activité_FTT[0].text);
                            sleep(0.2)
                            break
                        except:
                            continue

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            Statut_xpath = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[5]/b')
                            if Boucle_While == 15:
                                Statut.append("ERREUR")
                                break
                            Statut.append(Statut_xpath[0].text);
                            sleep(0.2)
                            break
                        except:
                            continue

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            Type_xpath = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[7]')
                            if Boucle_While == 15:
                                Type.append("ERREUR")
                                break
                            Type.append(Type_xpath[0].text);
                            sleep(0.2)
                            break
                        except:
                            continue

                    Boucle_While = 0
                    while True:
                        try:
                            Boucle_While += 1;
                            sleep(0.2)
                            Marque_xpath = driver.find_elements_by_xpath('//*[@id="tabs-1"]/fieldset[1]/span[11]')
                            if Boucle_While == 15:
                                Marque.append("ERREUR")
                                break
                            Marque.append(Marque_xpath[0].text);
                            sleep(0.2)
                            break
                        except:
                            continue

                    if ("RACC" == Type[c]) and ("CLOTURE TERMINEE" == Statut[c]):
                        for i in range(10, 30):
                            Press_cloture(i)
                        for i in range(10, 30):
                            Press_Rapport(i)
                        for i in range(10, 30):
                            Press_Raccordement(i)
                        for i in range(10, 30):
                            Add_Batiment(i)
                        for i in range(10, 30):
                            Add_Activité(i)
                    else:
                        Batiment.append("NONE")
                        RECO.append("NONE")

                    for i in range(1, 60):
                        Press_exit(i)
                    c += 1
                    print("")
                    print(Activité[-1])
                    print("")
                    sleep(2)

except:
    if len(Nom_et_Prenom) != len(Numéro_as):
        Nom_et_Prenom.pop()

    print("")
    print("")
    print("")
    print("!!!!!!!!!!!!!!!!!!!!!!!!!! ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print("")
    print("")
    print("")

    col_1 = 'Technician'
    col_2 = 'Numéro AS'
    col_3 = 'Date'
    col_4 = 'Durée'
    col_5 = 'Statut'
    col_6 = 'Erreur'
    col_7 = 'Type'
    col_8 = 'Marque'
    col_9 = 'Batiment'
    col_10 = 'Activité'
    col_11 = 'Reconnexion'

    Data = pd.DataFrame(
        {col_1: Nom_et_Prenom, col_2: Numéro_as, col_3: Date_J, col_4: Durée, col_5: Statut, col_7: Type, col_8: Marque,
         col_9: Batiment, col_10: Activité, col_11: RECO})
    Data.to_excel('{}_{}_OSIRIS.xlsx'.format(Mois, Année), index=False)

col_1 = 'Technicien'
col_2 = 'Numéro AS'
col_3 = 'Date'
col_4 = 'Durée'
col_5 = 'Statut'
col_6 = 'Erreur'
col_7 = 'Type'
col_8 = 'Marque'
col_9 = 'Batiment'
col_10 = 'Activité'
col_11 = 'Reconnexion'

Data = pd.DataFrame(
    {col_1: Nom_et_Prenom, col_2: Numéro_as, col_3: Date_J, col_4: Durée, col_5: Statut, col_7: Type, col_8: Marque,
     col_9: Batiment, col_10: Activité, col_11: RECO})
Data.to_excel('{}_{}_OSIRIS.xlsx'.format(Mois, Année), index=False)

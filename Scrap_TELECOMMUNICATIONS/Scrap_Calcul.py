import pandas as pd

A = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='A', index_col=0)
E = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='E', index_col=0)
F = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='F', index_col=0)
G = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='G', index_col=0)
H = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='H', index_col=0)
I = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='I', index_col=0)
J = pd.read_excel('Juillet_2022_xxx.xlsx', usecols='J', index_col=0)

Nom = A.index
Statut = E.index
Type = F.index
Marque = G.index
Batiment = H.index
Activite = I.index
Reconnexion = J.index

# Constant :
Cloture = 0
Somme = 0

SAV_FTTB = 0
SAV_FTTH = 0

Presta_FTTB = 0
Presta_FTTH = 0

RECO = 0

IMMEUBLE = 0
PAVILLON = 0

RACC_FTTB = 0
RACC_FTTH = 0

Lounes = 0
FTTB = 0
FTTH = 0

SFR_PRO_Sav = 0
SFR_PRO_Rac = 0

for i in range(len(Type)):
    if 'CLOTURE TERMINEE' == Statut[i]:
        Cloture += 1

        if ('DJEBBARI Lounes' == Nom[i]) and ('SAV' == Type[i]):
            Lounes += 1

        if 'FTTB' == Activite[i]:
            FTTB += 1
            if ('RACC' == Type[i]) or ('RACC PREMIUM' == Type[i]):
                RACC_FTTB += 1
                Somme += 55
                continue
            elif 'SAV' == Type[i]:
                SAV_FTTB += 1
                Somme += 26
                continue
            elif 'PRESTA COMPL' == Type[i]:
                Presta_FTTB += 1
                Somme += 55
                continue
        elif 'FTTH' == Activite[i]:
            FTTH += 1
            if 'SAV' == Type[i]:
                SAV_FTTH += 1
                if 'SFR PRO' == Marque[i]:
                    SFR_PRO_Sav += 1
                Somme += 26
                continue
            if 'Reconnexion' == Reconnexion[i]:
                RECO += 1
                Somme += 50
                continue
            if ('RACC' == Type[i]) or ('RACC PREMIUM' == Type[i]):
                RACC_FTTH += 1
                if 'SFR PRO' == Marque[i]:
                    SFR_PRO_Rac += 1
                    Somme += 190
                    continue
                if 'Pavillon' == Batiment[i]:
                    PAVILLON += 1
                    Somme += 140
                    continue
                else :
                    IMMEUBLE += 1
                    Somme += 90
                    continue
            if 'PRESTA COMPL' == Type[i]:
                Presta_FTTH += 1
                Somme += 55
                continue

print("\n")
print("La somme totale est de : {} €".format(Somme - Lounes*2))
print("Le nombre d'opération total est de : {}".format(len(Statut)))
print("Soit : {} Cloture terminée.".format(Cloture))
print("------------------------------------------------------")
print("Détail des opérations : FTTB")
print("Nombres d'opérations en cloture terminée : {} FTTB".format(FTTB))
print("Raccordement : {} / {} € TOTAL".format(RACC_FTTB, RACC_FTTB*55))
print("SAV : {} / {} € TOTAL".format(SAV_FTTB, SAV_FTTB*26))
print("Prestation : {} / {} € TOTAL".format(Presta_FTTB, Presta_FTTB*55))
print("------------------------------------------------------")
print("Détail des opérations : FTTH")
print("Nombres d'opérations en cloture terminée : {} FTTH".format(FTTH))
print("Raccordement : {} avec {} IMMEUBLE / {} PAVILLON / {} SFR PRO.".format(RACC_FTTH, IMMEUBLE, PAVILLON, SFR_PRO_Rac))
print("Raccordement au TOTAL => {} €".format(IMMEUBLE*90 + PAVILLON*140 + SFR_PRO_Rac*190))
print("SAV : {} avec {} SFR PRO".format(SAV_FTTH, SFR_PRO_Sav))
print("SAV au TOTAL => {} €".format(SAV_FTTH*26))
print("Prestation : {} / {} € TOTAL".format(Presta_FTTH, Presta_FTTH*55))
print("Reconnexion : {} / {} € TOTAL".format(RECO, RECO*50))
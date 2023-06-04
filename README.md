# Scraping in python - SELENIUM

I've done two scrapping projects in python :  
- Take information from a site on equestrian centers 
- Make automated invoices for a large French telecommunications company.

For this two projet, you need to download chrome.exe in SELENIUM documentations.

## Scrapping projet 1 : Equestrian centers

With this one, i get all Title, adress, number, email and URL from all equestrian in french :

```python
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
```

Only need to set the good URL of the website and run the projet :

```python
Error = 1514
Site = f"https://www.xxx-xxx.com/recherche?perPage=5&page={Error}"
```

After, with panda, i get all information in a .xlsx for see result in EXCEL.

## Scrapping projet 2 : Telecommunications company

In this one, i create 2 files :

- `Scrap_Calcul.py`
- `Scrap.py`

`Scrap.py` it's for scrap all information and `Scrap_Calcul.py` it's for see al result get in preview file.

This one work like Equestrian centers, i get all information with selenium and after put all in a .xlsx file with pandas.
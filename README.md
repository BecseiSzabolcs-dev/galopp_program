# 🐎Galopp program

Ez a Python alkalmazás egy grafikus felhasználói felülettel (GUI) rendelkező eszköz, amelyet a Galopp hivatalos PDF-programjainak feldolgozására, a kivonatolt adatok szerkesztésére, valamint CSV fájlok és PowerPoint prezentációk generálására terveztem.

⚙️ Funkciók és Képességek
PDF Betöltése: Képes betölteni egy PDF fájlt, és automatikusan kivonatolni belőle a versennyel kapcsolatos adatokat (pl. futamok címei, lovasok/hajtók, lovak, időpontok, vélemények).

# Futtatás
A dist/ allat az ugeto.exe allat található fájlal indítható.


## Adatszerkesztő Felület (GUI): 
A kivonatolt adatok két fülön szerkeszthetők:

![types](https://github.com/user-attachments/assets/547864ed-05db-4801-91a9-00379586652e)

## Load PDF (PDF betöltése): 
Ezen keresztűl lehet megnyitni az adatokat

![load](https://github.com/user-attachments/assets/ff4a2663-3cca-440a-8050-00044a461b5b)

## Titles (Címek):
Itt találhatók a futamokra vonatkozó fő adatok (pl. Azonosító, Cím, Táv, Időpont, Start típusa, Vélemény).

![titles](https://github.com/user-attachments/assets/6d420d4f-31a4-44c3-8f73-a46663ff95ff)

## Drivers (Hajtók): 
Itt találhatók a futam résztvevőinek adatai (pl. Lószám, Lónév, Táv, Hajtó neve, Futam száma, Futott-e státusz).

![drivers](https://github.com/user-attachments/assets/09fabedf-6a0b-4ada-a7ff-96d443e5e463)

## Adatszerkesztés: 
sorok hozzáadása és törlése.

![sor hozzá adása](https://github.com/user-attachments/assets/7d3fbdb1-c9b2-402f-8481-4f224db11dd2)

Cellák módosítása dupala kattintással elérhetővéválik.

![change](https://github.com/user-attachments/assets/cabc76a8-6158-4587-9cc2-65a92344f093)

Cellák módosítása szelekciót követően hozzá add hatunk a cellához bármely gomb lenyomásával, illetve nyillakkal mozoghatun a táblázaton.


![select](https://github.com/user-attachments/assets/00b5c4e4-7e13-48db-81ef-cb4e67e5dcba)



## Keresés: 
Mindkét táblában van keresési funkció az adatok gyors szűrésére.

![search](https://github.com/user-attachments/assets/74b73f30-7fe8-4f24-8f7e-0933a6220be9)

## Adatok Mentése CSV-be: 
A szerkesztett adatok exportálhatók két külön CSV fájlba (titles_data.csv és drivers_data.csv), amelyek a csv/ mappában jönnek létre UTF-8 kódolással.

![csve](https://github.com/user-attachments/assets/f085d6eb-ae9c-4150-b034-8ccb9bef9b21)

## PPT Készítése: 
Egyéni formázású PowerPoint prezentációkat generál a szerkesztett adatokból (futamonként egy .pptx fájlt) a ppt/ mappában. A prezentációk tartalmazzák a futam adatait, a résztvevőket, és diát az eredményeknek/osztalékoknak.

![Capture](https://github.com/user-attachments/assets/cc2bf90e-15da-4e5d-a5b0-a0edfb4da3f9)


# PPT elkészítésének Módszere:
1. Betöltjük a PDF-et a "Load PDF" gombra kattintva.
2. Egy alapos szemle után elemntjük a csv-t a "Save Data to CSV" gombra katintva
3. Majd le nyomjuk a "Make PPT" gombot

Ezután ha minden jól ment akkor kapunk egy üzenetet:

![succes](https://github.com/user-attachments/assets/b4d57064-be32-47de-a5c9-4f58d2cf566f)

## könyvtár szerkezett:
### A program futatása után megjelenik két mappa:
1. ppt/
  Ez a mappa tartalamzza a pptx fájlokat.
2. csv/
  Ez a mappa tartalamzza a csv fájlokat.
   
```
ugeto_program/
├── ppt/
|   ├── I. futam.pptx
|   ├── II. futam.pptx
|   ├── III. futam.pptx
|   ├── ... .pptx
├── csv/
|   ├── drivers_data.csv
|   └── titles_data.csv
├── add macro.xlsm
├── clock.jpeg
└── ugeto.exe
```







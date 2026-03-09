# PDF Engine

Automatikus **PDF → Excel feldolgozó rendszer**.

A program különböző típusú PDF dokumentumokból kinyeri a szükséges adatokat és Excel fájlba rendezi őket előre definiált sablonok alapján.

A feldolgozás teljesen automatizált:

1. Indítsd el a programot
2. Húzd be a PDF-et a megfelelő Input mappába
3. Az Excel automatikusan létrejön az Output mappában

---

# Támogatott PDF típusok

A motor jelenleg három különböző PDF struktúrát kezel.

---

## 1. EK lista

Rendelési PDF-ek feldolgozása.

Kinyert adatok:

- Cikkszám
- Megnevezés
- EK
- Áregység

PDF bemenet:

Input_EK

Kimenet:

VEVOSZAM EK.xlsx

---

## 2. Sajátkészlet lista

Ugyanabból a PDF struktúrából dolgozik mint az EK lista, de más adatokat generál.

Kinyert adatok:

- Cikkszám
- Megnevezés
- Mennyiség

PDF bemenet:

Input

A program automatikusan felismeri ha nem GE típusú PDF.

---

## 3. GE Blanket Release

GE beszállítói Blanket Release dokumentumok feldolgozása.

Kinyert adatok:

- GE cikkszám
- RECA cikkszám
- Mennyiség
- GE ár

PDF bemenet:

Input

Kimenet:

RELEASE_NUMBER.xlsx

---

# Mappastruktúra

A projekt az alábbi struktúrát használja:

PDF_ENGINE
│
├─ engine.py  
├─ indit.bat  
│
├─ Python  
│   └─ python.exe  
│
├─ Input  
│
├─ Input_EK  
│
├─ Output  
│
├─ lookup.xlsx  
│
└─ templates  
&nbsp;&nbsp;&nbsp;&nbsp;├─ EK_template.xlsx  
&nbsp;&nbsp;&nbsp;&nbsp;├─ GE_template.xlsx  
&nbsp;&nbsp;&nbsp;&nbsp;└─ SAJATKESZLET_template.xlsx  

---

# Telepítés

A projekt **portable Python környezetet használ**, ezért nem szükséges Python telepítése a rendszerre.

Ha a Python mappa hiányzik, töltsd le az embeddable Python verziót:

https://www.python.org/downloads/windows/

Az **Windows embeddable package** fájlt csomagold ki ide:

PDF_ENGINE/Python/

---

# Indítás

A program indítása:

indit.bat

Dupla kattintással elindítható.

A program ezután figyeli az Input mappákat.

---

# Használat

## EK lista feldolgozás

PDF bemásolása:

Input_EK

---

## GE vagy Sajátkészlet feldolgozás

PDF bemásolása:

Input

---

A feldolgozott Excel automatikusan ide kerül:

Output

---

# Lookup fájl

A `lookup.xlsx` fájl tartalmazza a cikkszám megfeleltetéseket.

---

## EK és Sajátkészlet

Lookup sheet:

Sheet1

Oszlopok:

B – SAP cikkszám  
D – Megnevezés  
E – Áregység  
H – EK  

---

## GE

Lookup sheet:

Sheet2

Oszlopok:

GE cikkszám  
RECA cikkszám  

---

# Excel Template rendszer

A program **template alapú Excel generálást használ**.

Ez azt jelenti, hogy az Excel struktúráját a sablon fájlok határozzák meg.

Template mappa:

templates/

Fájlok:

- EK_template.xlsx
- GE_template.xlsx
- SAJATKESZLET_template.xlsx

Ha módosítod az oszlopokat a template-ben, a program automatikusan ehhez igazítja a kimenetet.

---

# Automatikus PDF felismerés

A motor automatikusan felismeri a dokumentum típusát.

GE felismerés:

- Item
- EACH
- Net Unit Price

Ha ezek szerepelnek a dokumentumban, a motor GE feldolgozást használ.

---

# Főbb jellemzők

✔ Automatikus PDF felismerés  
✔ Template alapú Excel generálás  
✔ Lookup alapú cikkszám megfeleltetés  
✔ Portable Python futtatás  
✔ Drag & Drop használat  
✔ Automatikus mappa figyelés  

---

# Használt könyvtárak

A projekt Python könyvtárakat használ:

- pdfplumber
- openpyxl
- watchdog

---

# Követelmények

- Windows operációs rendszer
- Excel kompatibilis környezet
- Python embeddable package (portable)


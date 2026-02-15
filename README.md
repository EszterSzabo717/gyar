# Forgalom riport automatizálás Pythonban
forgalomkimutatas
A script több telephely forgalmi adataiból készít havi kimutatást, kezeli az árlista változásokat (külső árlistából), és generál összesítő, napi bontás és telephelyenkénti részletes lapokat formázott Excelben.

## Funkciók

- több Excel fájl feldolgozása
- árlista változások kezelése dátum alapján
- bevétel automatikus számítása
- havi szűrés
- riport generálás:
  - Összesítő lap
  - Napi bontás
  - Telephelyenkénti részletes lap
- automatikus formázás (Ft, dátum, oszlopszélesség)

## Használat

1. Helyezd a forgalmi fájlokat a script mappájába
2. Add meg az időszakot (YYYYMM)
3. A program elkészíti a riportot Excelben

## Technológiák

- Python
- pandas
- openpyxl
- tkinter

## Cél

Valós üzleti riportkészítés automatizálása manuális Excel-munka kiváltására.

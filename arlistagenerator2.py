import pandas as pd
import random
from datetime import date, timedelta

# ===== dátumok generálása 2026.01.01 - 2026.07.01 között =====
datumok = []
d = date(2026, 1, 1)
while d <= date(2026, 7, 1):
    datumok.append(d)
    # kb havonta 1-2 változás
    d += timedelta(days=random.choice([14, 20, 30]))

# ===== termékek =====
termekek = [f"Termék{i}" for i in range(1, 13)]

# ===== induló árak =====
arak = {termek: random.randint(300, 2000) for termek in termekek}

sorok = []

for datum in datumok:
    sor = {"Dátum": datum}
    for termek in termekek:
        # néha változzon az ár
        if random.random() < 0.35:
            valtozas = random.randint(-150, 150)
            uj_ar = max(100, arak[termek] + valtozas)
            arak[termek] = uj_ar
        sor[termek] = arak[termek]
    sorok.append(sor)

df = pd.DataFrame(sorok)
df = df.sort_values("Dátum")

# ===== mentés =====
utvonal = r"C:\Users\szabo\Desktop\pythonscripts\gyar\arlista.xlsx"
df.to_excel(utvonal, index=False)

print("Új arlista.xlsx elkészült!")

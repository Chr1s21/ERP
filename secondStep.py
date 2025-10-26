import pandas as pd

# 1️⃣ Datei laden
df = pd.read_excel("2Rohdaten_nurVolumenplanung.xlsx")

# 2️⃣ Fehlende Werte zählen
fehlende_werte = df.isna().sum()

# 3️⃣ Ausgabe im Terminal schön anzeigen
print("🧩 Fehlende Werte pro Spalte:\n")
print(fehlende_werte)

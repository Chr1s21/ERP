import pandas as pd

# Datei laden
df = pd.read_excel("2Rohdaten_nurVolumenplanung.xlsx")

# Zeilen löschen, wo der Produktname fehlt
df = df.dropna(subset=["Produktname"])

# Ergebnis speichern
df.to_excel("3Rohdaten_ohneLeereProduktnamen.xlsx", index=False)

print("✅ Fertig! Alle Zeilen ohne Produktname wurden gelöscht.")
print("Neue Größe der Tabelle:", df.shape)

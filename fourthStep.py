import pandas as pd

# Datei laden
df = pd.read_excel("3Rohdaten_ohneLeereProduktnamen.xlsx")

# Gruppieren und nach tatsächlicher Liefermenge sortieren
artikel_liefermengen = (
    df.groupby(["Artikelnummer", "Produktname"])["Tatsächliche Liefermenge"]
    .sum()
    .reset_index()
    .sort_values("Tatsächliche Liefermenge", ascending=False)
)

# Ergebnis speichern (keine Ausgabe im Terminal)
artikel_liefermengen.to_excel("4Artikel_Liefermengen_sortiert.xlsx", index=False)

print("✅ Datei 'Artikel_Liefermengen_sortiert.xlsx' wurde erfolgreich erstellt!")

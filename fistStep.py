import pandas as pd

# 1️⃣ Datei einlesen
datei = "1Rohdaten.xlsx"  # ggf. anpassen
df = pd.read_excel(datei)

# 2️⃣ Relevante Spalten für Volumenplanung (inkl. Baumarktartikel)
relevante_spalten = [
    "Baumarktartikel",   # Produktname / Materialname
    "matnr",             # Artikelnummer
    "modulgruppen",      # Artikelgruppe
    "Baumarkt",          # Kunde
    "bedmo",             # Bedarfsmonat
    "versmo",            # Versandmonat
    "progmo",            # Prognosemonat 1
    "progmo2",           # Prognosemonat 2
    "bedmo_mg",          # Tatsächliche Liefermenge
    "prog_mg1",          # Prognosemenge 1
    "prog_mg2",          # Prognosemenge 2
    "verbauquote",       # Anteil tatsächlich verbraucht
    "ct_kapa",           # Kapazität
    "ct_auslastung",     # Auslastung
    "ct_volds",          # Produktionsvolumen
    "diff_faktorjahr_wpp1",  # Abweichung Vorjahr
    "vol_gesamt_lab_mg"  # Volumen gesamt
]

# 3️⃣ Nur vorhandene Spalten übernehmen
vorhandene_spalten = [s for s in relevante_spalten if s in df.columns]
df_relevant = df[vorhandene_spalten]

# 4️⃣ Verständliche Spaltennamen vergeben
neue_namen = {
    "Baumarktartikel": "Produktname",
    "matnr": "Artikelnummer",
    "modulgruppen": "Artikelgruppe",
    "Baumarkt": "Kunde",
    "bedmo": "Bedarfsmonat",
    "versmo": "Versandmonat",
    "progmo": "Prognosemonat 1",
    "progmo2": "Prognosemonat 2",
    "bedmo_mg": "Tatsächliche Liefermenge",
    "prog_mg1": "Prognosemenge 1",
    "prog_mg2": "Prognosemenge 2",
    "verbauquote": "Verbrauchsanteil",
    "ct_kapa": "Kapazität",
    "ct_auslastung": "Auslastung",
    "ct_volds": "Produktionsvolumen",
    "diff_faktorjahr_wpp1": "Abweichung Vorjahr",
    "vol_gesamt_lab_mg": "Volumen gesamt"
}

df_relevant = df_relevant.rename(columns=neue_namen)

# 5️⃣ Neue Datei speichern
df_relevant.to_excel("Rohdaten_nurVolumenplanung.xlsx", index=False)

print("✅ Fertig! Die Datei 'Rohdaten_nurVolumenplanung.xlsx' enthält nun auch die Spalte 'Produktname' und alle wichtigen Felder.")

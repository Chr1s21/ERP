import pandas as pd

# 1️⃣ Originaldatei einlesen
datei = "Rohdaten.xlsx"   # ← hier ggf. anpassen, falls der Name anders ist
df = pd.read_excel(datei)

# 2️⃣ Wörterbuch für neue Spaltennamen (alte → neue Namen)
neue_namen = {
    "matnr": "Artikelnummer",
    "kundnr": "Kundennummer",
    "vkbel": "Verkaufsbeleg",
    "kundabl": "Kundenauftrag",
    "wavor_bme": "Bestelleinheit",
    "bedkw": "Bedarfs-Kalenderwoche",
    "bedmo": "Bedarfs-Monat",
    "verskw": "Versand-Kalenderwoche",
    "versmo": "Versand-Monat",
    "wavor_bstlmg": "Bestellmenge gesamt",
    "wavor_anteilpromonat": "Monatlicher Anteil an Bestellung",
    "wavor_bstlmengemonat": "Bestellmenge pro Monat",
    "wavor_bstlmgjahr": "Bestellmenge pro Jahr",
    "cc_bez": "Region",
    "vbap_bstlmg": "Positionsbestellmenge",
    "bstlmgeh": "Mengeneinheit",
    "Baumarkt": "Kunde",
    "gew_bto": "Gewicht pro Teil",
    "geweh": "Gewichtseinheit",
    "gew_bto_kg": "Gewicht in kg",
    "lft_land": "Lieferland",
    "lft_ort": "Lieferort",
    "ltm_zin_lt_kategorie": "Produktkategorie (intern)",
    "zin_lt_1_menge": "Menge pro Einheit",
    "ltm_zout_lt_kategorie": "Produktkategorie (extern)",
    "lt_1_bruttogew_in_kg": "Bruttogewicht in kg",
    "lt_1_laenge": "Länge (mm)",
    "lt_1_breite": "Breite (mm)",
    "lt_1_hoehe": "Höhe (mm)",
    "lt_1_me": "Maßeinheit",
    "lt_1_flaeche": "Fläche (m²)",
    "lt_1_feh": "Flächeneinheit",
    "lt_1_volumen": "Volumen (m³)",
    "lt_1_veh": "Volumeneinheit",
    "lt_1_menge_in_lt": "Menge in Lieferteilen",
    "bedmo_mg": "Tatsächliche Liefermenge",
    "progmo": "Prognosemonat 1",
    "prog_mg1": "Prognosemenge 1",
    "progmo2": "Prognosemonat 2",
    "prog_mg2": "Prognosemenge 2",
    "diff_faktorjahr_wpp1": "Abweichungsfaktor Vorjahr 1",
    "diff_faktorjahr_wpp2": "Abweichungsfaktor Vorjahr 2",
    "ct_kapa": "Kapazität",
    "ct_auslastung": "Auslastung",
    "ct_volds": "Produktionsvolumen",
    "verbauquote": "Verbrauchsanteil",
    "prog_vol1": "Prognose-Volumen 1",
    "prog_vol2": "Prognose-Volumen 2"
}

# 3️⃣ Spalten umbenennen
df = df.rename(columns=neue_namen)

# 4️⃣ Neue Datei speichern
df.to_excel("rohdatenMitNamen.xlsx", index=False)

print("✅ Datei 'rohdatenMitNamen.xlsx' wurde erfolgreich erstellt!")


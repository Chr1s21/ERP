import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import seaborn as sns

# --- KONFIGURATION ---
INPUT_FILE_LIVE = "final.xlsx" 
OUTPUT_DIR = "./output/final"
OUTPUT_FILE_EXCEL = "Live_Forecast_Result.xlsx"

# Erstelle Ausgabeordner
os.makedirs(OUTPUT_DIR, exist_ok=True)
sns.set_theme(style="whitegrid")

# ---------------------------------------------------------
# 1. HILFSFUNKTIONEN
# ---------------------------------------------------------

def clean_keys(df, col_kunde='Kunde', col_monat='Monat'):
    """Bereinigt Schlüssel für sauberen Merge."""
    # Monat zu Int (Fehler abfangen)
    df[col_monat] = pd.to_numeric(df[col_monat], errors='coerce').fillna(0).astype(int)
    
    # Kunde (Werk) zu String ohne ".0" am Ende
    if col_kunde in df.columns:
        # Erst in String wandeln
        s = df[col_kunde].astype(str)
        # ".0" entfernen (falls Excel Zahlen als Floats lädt)
        s = s.str.replace(r'\.0$', '', regex=True)
        # Whitespace weg und Großbuchstaben
        # HIER WAR DER FEHLER: Es muss .str.upper() heißen
        df[col_kunde] = s.str.strip().str.upper()
        
    return df

def calculate_factor(row):
    """Berechnet den Faktor pro Zeile."""
    ziel = row['Ziel_Summe']
    ist = row['Bottom_Up_Summe']
    
    if ist == 0: return 0.0
    if ziel == 0: return 1.0
    
    return ziel / ist

# ---------------------------------------------------------
# 2. DATEN LADEN & DEMO-PLAN GENERIEREN
# ---------------------------------------------------------

def load_and_prep_data():
    print("Step 1: Lade Live-Daten...")
    
    if not os.path.exists(INPUT_FILE_LIVE):
        print(f"❌ Fehler: Datei '{INPUT_FILE_LIVE}' nicht gefunden.")
        return pd.DataFrame(), pd.DataFrame()

    # 1. LIVE DATEN LESEN
    try:
        # Prio 1: Excel
        df_raw = pd.read_excel(INPUT_FILE_LIVE)
        print(f"   ✅ Excel geladen: {len(df_raw)} Zeilen.")
    except Exception:
        print("   ⚠️  Excel-Modus fehlgeschlagen, versuche CSV-Fallback...")
        try:
            # Prio 2: CSV (falls Endung falsch oder Format anders)
            df_raw = pd.read_csv(INPUT_FILE_LIVE, sep=None, engine='python', on_bad_lines='skip')
            print(f"   ✅ CSV geladen: {len(df_raw)} Zeilen.")
        except Exception as e:
            print(f"❌ Kritischer Fehler beim Laden: {e}")
            return pd.DataFrame(), pd.DataFrame()

    # 2. SPALTEN MAPPING (werk -> Kunde)
    col_mapping = {
        'matnr': 'Artikel',
        'werk': 'Kunde',       # Hier greifen wir auf die Spalte 'werk' zu
        'progmo': 'Monat',
        'prog_mg1': 'Menge'
    }
    
    # Prüfen ob Spalten da sind
    missing = [c for c in col_mapping.keys() if c not in df_raw.columns]
    if missing:
        print(f"❌ Fehler: Spalten fehlen in Datei: {missing}")
        print(f"   Vorhandene Spalten: {list(df_raw.columns)}")
        # Notfall-Versuch: Falls Spalten groß geschrieben sind (WERK statt werk)
        df_raw.columns = [c.lower() for c in df_raw.columns]
        missing_lower = [c for c in col_mapping.keys() if c not in df_raw.columns]
        if not missing_lower:
            print("   (Habe Spaltennamen in Kleinbuchstaben konvertiert und gefunden!)")
        else:
            return pd.DataFrame(), pd.DataFrame()

    # Daten extrahieren (Jahr 1)
    p1 = df_raw[list(col_mapping.keys())].copy()
    p1 = p1.rename(columns=col_mapping)
    p1['Gruppe'] = 'Live-Daten' 

    # Jahr 2 (optional)
    if 'progmo2' in df_raw.columns and 'prog_mg2' in df_raw.columns:
        p2 = df_raw[['matnr', 'werk', 'progmo2', 'prog_mg2']].copy()
        p2.columns = ['Artikel', 'Kunde', 'Monat', 'Menge']
        p2['Gruppe'] = 'Live-Daten'
        df_forecast = pd.concat([p1, p2], ignore_index=True)
    else:
        df_forecast = p1

    # Bereinigen
    df_forecast = df_forecast.dropna(subset=['Monat', 'Menge'])
    df_forecast = df_forecast[df_forecast['Menge'] > 0]
    df_forecast = clean_keys(df_forecast)
    
    print(f"   ✅ Prognose erstellt: {len(df_forecast)} Datensätze.")
    if not df_forecast.empty:
        print(f"      Beispiel Werk: {df_forecast['Kunde'].iloc[0]}")

    # -------------------------------------------------------
    # DEMO-PLAN GENERIEREN
    # -------------------------------------------------------
    print("\n   ✨ GENERIERE DEMO-PLAN (basierend auf Werks-Daten)...")
    
    df_plan_dummy = df_forecast.groupby(['Kunde', 'Monat'])['Menge'].sum().reset_index()
    df_plan_dummy = df_plan_dummy.rename(columns={'Menge': 'Ziel_Summe'})
    
    # Simulation
    np.random.seed(42) 
    random_noise = np.random.uniform(0.9, 1.25, size=len(df_plan_dummy)) 
    df_plan_dummy['Ziel_Summe'] = (df_plan_dummy['Ziel_Summe'] * random_noise).astype(int)
    
    print(f"   ✅ Plan simuliert für {len(df_plan_dummy)} Werk/Monat-Kombinationen.")

    return df_forecast, df_plan_dummy

# ---------------------------------------------------------
# 3. GLÄTTUNG (RECONCILIATION)
# ---------------------------------------------------------

def run_reconciliation(df_forecast, df_plan):
    print("\nStep 2: Führe Abgleich durch...")
    
    # Aggregation
    bu_agg = df_forecast.groupby(['Kunde', 'Monat'])['Menge'].sum().reset_index()
    bu_agg = bu_agg.rename(columns={'Menge': 'Bottom_Up_Summe'})
    
    # Merge
    merged = pd.merge(bu_agg, df_plan, on=['Kunde', 'Monat'], how='inner')
    
    if merged.empty:
        print("❌ FEHLER: Keine Matches gefunden! Kunde/Monat passen nicht zusammen.")
        return pd.DataFrame()

    # Faktor
    merged['Faktor'] = merged.apply(calculate_factor, axis=1)
    
    avg_factor = merged['Faktor'].mean()
    print(f"   📊 Statistik: Ø Faktor = {avg_factor:.2f}")
    
    # Anwenden
    df_final = pd.merge(df_forecast, merged[['Kunde', 'Monat', 'Faktor']], on=['Kunde', 'Monat'], how='left')
    
    df_final['Faktor'] = df_final['Faktor'].fillna(1.0)
    df_final['Menge_Geglaettet'] = (df_final['Menge'] * df_final['Faktor']).round(0).astype(int)
    
    return df_final

# ---------------------------------------------------------
# 4. MAIN
# ---------------------------------------------------------

def main():
    df_forecast, df_plan = load_and_prep_data()
    if df_forecast.empty: return

    df_final = run_reconciliation(df_forecast, df_plan)
    if df_final.empty: return

    # Speichern
    out_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE_EXCEL)
    cols = ['Artikel', 'Kunde', 'Monat', 'Menge', 'Faktor', 'Menge_Geglaettet']
    
    try:
        df_final[cols].to_excel(out_path, index=False)
        print(f"\n✅ FERTIG! Datei gespeichert: {out_path}")
    except Exception as e:
        print(f"❌ Fehler beim Speichern: {e}")
    
    # Plot
    try:
        plot_data = df_final.groupby('Monat')[['Menge', 'Menge_Geglaettet']].sum().reset_index()
        plot_data = plot_data.sort_values('Monat')
        plot_data['Monat'] = plot_data['Monat'].astype(str)
        
        plt.figure(figsize=(12, 6))
        plt.plot(plot_data['Monat'], plot_data['Menge'], label='Original (Werk)', marker='o', linestyle='--')
        plt.plot(plot_data['Monat'], plot_data['Menge_Geglaettet'], label='Geglättet (Ziel)', marker='x')
        plt.title("Ergebnis: Abgleich Forecast vs. Ziel")
        plt.legend()
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(os.path.join(OUTPUT_DIR, "Demo_Chart.png"))
        print("   📈 Chart gespeichert.")
    except:
        pass

if __name__ == "__main__":
    main()
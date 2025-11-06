import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from sklearn.preprocessing import MinMaxScaler # Wichtig für den Trend-Vergleich

# Setzt einen sauberen "Seaborn"-Stil für alle Plots
sns.set_theme(style="whitegrid")

# --- Schritt 1: Bottom-up PROGNOSE (Stück) laden & aggregieren ---

def load_and_aggregate_prognosis(raw_data_path="rohdaten.xlsx", prog_col='prog_mg1', date_col='progmo'):
    """
    Lädt die Rohdaten und aggregiert die Prognose-Stückzahlen (prog_col)
    anhand der korrekten Datumspalte (date_col).
    """
    print(f"Lade Prognose: '{prog_col}' basierend auf Datum '{date_col}'...")
    try:
        df_raw = pd.read_excel(raw_data_path)
        # Erstelle Datumsobjekt aus der korrekten Prognosespalte (progmo oder progmo2)
        df_raw['bedmo_date'] = pd.to_datetime(df_raw[date_col], format='%Y%m')
    except Exception as e:
        print(f"FEHLER beim Laden der Rohdaten ('{raw_data_path}'): {e}")
        return None
        
    # Aggregation: Summiere die Prognose-Stückzahlen
    agg_definition = {
        prog_col: "sum"
    }
    
    df_prognosis = (
        df_raw.groupby(["Baumarkt", "bedmo_date"])
        .agg(agg_definition)
        .reset_index()
        .rename(columns={prog_col: "Prognose_Stueck"})
    )
    
    # Filtere leere Prognosen/Datumseinträge heraus
    df_prognosis = df_prognosis[df_prognosis['Prognose_Stueck'] > 0]
    
    print(f"Prognose (Stück) erfolgreich aggregiert: {len(df_prognosis)} Zeilen.")
    return df_prognosis

# --- Schritt 2: Vertriebsplan (Euro) laden & transformieren (NEU) ---

def load_sales_plan(plan_filepath="BAUMARKTPROGRAMM.xlsx"):
    """
    Lädt den komplexen Vertriebsplan (Kreuztabelle) und
    transformiert ihn in ein "langes" Format (pro Baumarkt/Monat).
    """
    print(f"Lade Vertriebsplan (Euro) aus '{plan_filepath}'...")
    try:
        # Lade CSV mit Multi-Header (Zeile 1 = Jahr, Zeile 2 = Monat)
        # Index_col=0 setzt 'Baumarkt' als Index
        df_plan = pd.read_excel(plan_filepath)
        
        # 1. Bereinige Spalten: Entferne 'Ergebnis', 'Baureihe', etc.
        cols_to_keep = [col for col in df_plan.columns if 'Ergebnis' not in col[1] and 'Baureihe' not in col[0] and 'Unnamed' not in col[0]]
        df_plan = df_plan[cols_to_keep]

        # 2. "Entpivoten" (Stacken):
        # Stapel zuerst die Jahre (level=0)
        df_stacked = df_plan.stack(level=0)
        # Stapel dann die Monate (standardmäßig level=1)
        df_stacked = df_stacked.stack()

        # 3. In sauberen DataFrame umwandeln
        df_long = df_stacked.reset_index()
        df_long.columns = ['Baumarkt', 'Jahr', 'Monat', 'Plan_Umsatz']

        # 4. Daten bereinigen:
        # Entferne Tausender-Trennzeichen (Punkte) und wandle in Zahl um
        df_long['Plan_Umsatz'] = df_long['Plan_Umsatz'].replace(r'\.', '', regex=True).astype(float)
        
        # Monate in Zahlen umwandeln
        month_map = {
            'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAI': 5, 'JUN': 6,
            'JUL': 7, 'AUG': 8, 'SEP': 9, 'OKT': 10, 'NOV': 11, 'DEZ': 12
        }
        df_long['Monat_Num'] = df_long['Monat'].map(month_map)

        # 5. Finale Datumspalte 'bedmo_date' erstellen
        df_long['bedmo_date'] = pd.to_datetime(
            df_long['Jahr'].astype(str) + '-' + df_long['Monat_Num'].astype(str) + '-01'
        )
        
        # Nur relevante Spalten behalten
        df_plan_agg = df_long[['Baumarkt', 'bedmo_date', 'Plan_Umsatz']]
        
    except Exception as e:
        print(f"FEHLER beim Laden des Vertriebsplans: {e}")
        print("Stellen Sie sicher, dass die Datei 'BAUMARKTPROGRAMM 2025-09(Kreuztabelle).csv' im selben Ordner liegt.")
        return None
        
    print(f"Vertriebsplan (Euro) erfolgreich geladen und transformiert: {len(df_plan_agg)} Zeilen.")
    return df_plan_agg

# --- Schritt 3, 4 & 5: Vergleich, Abweichung & Visualisierung ---

def compare_and_visualize_trends(df_prognosis, df_plan, output_dir, output_suffix, top_n=5):
    """
    Führt die beiden Datensätze zusammen, normalisiert die Trends (Stück vs. Euro),
    berechnet die Abweichungen und visualisiert die auffälligsten.
    output_suffix wird an Dateinamen angehängt (z.B. '_prog1' oder '_prog2')
    """
    print(f"Vergleiche Trends für {output_suffix}...")
    
    # 1. Zusammenführen der beiden Pläne (Aufgabe 1)
    df_vergleich = pd.merge(
        df_prognosis,
        df_plan,
        on=['Baumarkt', 'bedmo_date'],
        how='inner' # 'inner' = nur Zeiträume, wo BEIDE Daten haben
    )
    df_vergleich = df_vergleich.fillna(0)
    
    if df_vergleich.empty:
        print(f"FEHLER für {output_suffix}: Kein Überlapp zwischen Prognose und Vertriebsplan gefunden.")
        print("Bitte prüfen Sie die 'Baumarkt'-Namen und Zeiträume in beiden Dateien.")
        return

    # 2. Normalisieren (Die "Brücke" bauen)
    scaler = MinMaxScaler()
    all_scaled_dfs = []
    
    for baumarkt in df_vergleich['Baumarkt'].unique():
        df_gruppe = df_vergleich[df_vergleich['Baumarkt'] == baumarkt].copy()
        
        # Skalierer anwenden auf die Spalten
        df_gruppe[['Prognose_Trend', 'Plan_Trend']] = \
            scaler.fit_transform(df_gruppe[['Prognose_Stueck', 'Plan_Umsatz']])
            
        all_scaled_dfs.append(df_gruppe)
        
    df_vergleich = pd.concat(all_scaled_dfs)
    
    # 3. Abweichung der Trends berechnen (Aufgabe 2)
    df_vergleich['Trend_Abweichung'] = \
        df_vergleich['Prognose_Trend'] - df_vergleich['Plan_Trend']
        
    print(f"Trend-Vergleich {output_suffix} abgeschlossen.")
    df_vergleich.to_excel(os.path.join(output_dir, f"Abweichungsanalyse_Trends{output_suffix}.xlsx"), index=False)
    
    # 4. Visualisierung auffälliger Abweichungen (Aufgabe 3)
    print(f"Erstelle Visualisierungen der Top {top_n} Abweichungen für {output_suffix}...")

    # Finde die Baumärkte mit der größten durchschnittlichen Abweichung
    df_agg_abweichung = df_vergleich.groupby('Baumarkt')['Trend_Abweichung'].apply(lambda x: x.abs().mean()).reset_index()
    df_agg_abweichung = df_agg_abweichung.sort_values(by='Trend_Abweichung', ascending=False)
    
    # Plotte die Top N Abweichler
    for baumarkt in df_agg_abweichung.head(top_n)['Baumarkt']:
        
        df_plot = df_vergleich[df_vergleich['Baumarkt'] == baumarkt]
        
        df_melted = df_plot.melt(
            id_vars=['Baumarkt', 'bedmo_date'],
            value_vars=['Prognose_Trend', 'Plan_Trend'],
            var_name='Datenquelle',
            value_name='Normalisierter_Trend (0-1)'
        )
        
        df_melted['Datenquelle'] = df_melted['Datenquelle'].replace({
            'Prognose_Trend': f'Bottom-up Prognose (Stück-Trend {output_suffix})',
            'Plan_Trend': 'Vertriebsplan (Euro-Trend)'
        })

        plt.figure(figsize=(15, 7))
        sns.lineplot(
            data=df_melted,
            x='bedmo_date',
            y='Normalisierter_Trend (0-1)',
            hue='Datenquelle',
            style='Datenquelle',
            markers=True,
            linewidth=2.5
        )
        
        
        plt.title(f"Auffällige Abweichung {output_suffix}: Trend-Vergleich für Baumarkt {baumarkt}", fontsize=16)
        plt.ylabel("Normalisierte Skala (0 = Min, 1 = Max)")
        plt.xlabel("Monat")
        plt.legend()
        plt.grid(True, which='both', linestyle='--', linewidth=0.5)
        
        plt.savefig(os.path.join(output_dir, f"5_Trendvergleich{output_suffix}_{baumarkt}.png"))
        plt.close()

    print(f"Visualisierungen für {output_suffix} in '{output_dir}' gespeichert.")

# --- Haupt-Logik ---

def main():
    # Output-Verzeichnisse erstellen
    output_dir = "./output"
    plot_dir = "./output/plots/2"
    os.makedirs(plot_dir, exist_ok=True)
    
    # -----------------------------------------------------------------
    # Schritt A: Vertriebsplan laden (wird für beide Analysen benötigt)
    # -----------------------------------------------------------------
    df_plan = load_sales_plan(
        plan_filepath="BAUMARKTPROGRAMM.xlsx"
    )
    if df_plan is None: 
        print("Analyse wird beendet, da Vertriebsplan nicht geladen werden konnte.")
        return

    # -----------------------------------------------------------------
    # Analyse 1: Vergleich mit PROG_MG1 (Jahr +1)
    # -----------------------------------------------------------------
    print("\n--- STARTE ANALYSE FÜR PROG_MG1 (JAHR 1) ---")
    df_prognosis_1 = load_and_aggregate_prognosis(
        raw_data_path="rohdaten.xlsx",
        prog_col='prog_mg1',
        date_col='progmo'
    )
    
    if df_prognosis_1 is not None:
        compare_and_visualize_trends(
            df_prognosis_1, 
            df_plan, 
            plot_dir, 
            output_suffix="_prog1", 
            top_n=5
        )
    else:
        print("Überspringe Vergleich für prog_mg1, da Daten nicht geladen werden konnten.")

    # -----------------------------------------------------------------
    # Analyse 2: Vergleich mit PROG_MG2 (Jahr +2)
    # -----------------------------------------------------------------
    print("\n--- STARTE ANALYSE FÜR PROG_MG2 (JAHR 2) ---")
    df_prognosis_2 = load_and_aggregate_prognosis(
        raw_data_path="rohdaten.xlsx",
        prog_col='prog_mg2',
        date_col='progmo2'
    )
    
    if df_prognosis_2 is not None:
        compare_and_visualize_trends(
            df_prognosis_2, 
            df_plan, 
            plot_dir, 
            output_suffix="_prog2", 
            top_n=5
        )
    else:
        print("Überspringe Vergleich für prog_mg2, da Daten nicht geladen werden konnten.")

if __name__ == "__main__":
    main()


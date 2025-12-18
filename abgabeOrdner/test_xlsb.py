import pandas as pd


def test_xlsb():
    """Testet ob xlsb-Dateien geöffnet werden können."""

    # Test 1: Prüfe ob pyxlsb installiert ist
    print("Test 1: Prüfe pyxlsb Installation...")
    try:
        import pyxlsb

        print("   ✅ pyxlsb ist installiert")
    except ImportError:
        print("   ❌ pyxlsb ist NICHT installiert")
        print("   → Installiere mit: pip install pyxlsb")
        return False

    # Test 2: Versuche eine xlsb-Datei zu öffnen
    print("\nTest 2: Versuche xlsb-Datei zu öffnen...")

    # Hier den Pfad zur xlsb-Datei eintragen
    xlsb_datei = "dieEchtenDaten.xlsb"  # ← anpassen

    try:
        df = pd.read_excel(xlsb_datei, engine="pyxlsb")
        print(f"   ✅ Datei erfolgreich geöffnet!")
        print(f"   📊 Zeilen: {len(df)}, Spalten: {len(df.columns)}")
        print(f"   📋 Spalten: {list(df.columns[:5])}...")
        return True
    except FileNotFoundError:
        print(f"   ⚠️ Datei '{xlsb_datei}' nicht gefunden")
        print("   → Passe den Dateipfad an")
        return False
    except Exception as e:
        print(f"   ❌ Fehler: {e}")
        return False


if __name__ == "__main__":
    print("=" * 50)
    print("XLSB Test")
    print("=" * 50)
    test_xlsb()

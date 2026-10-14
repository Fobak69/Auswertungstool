# 📊 Excel-Prüfer & Korrektur-Assistent

Ein leistungsstarkes, übersichtliches und leicht zu bedienendes Desktop- & Web-Programm zum automatischen Abgleich von Excel-Dateien gegen eine Referenz- bzw. Stammdaten-Datei.

---

## 🌟 Hauptfunktionen

1. **Intelligente Fehlererkennung**:
   - **🔄 Echte Zahlendreher**: Erkennt vertauschte benachbarte Ziffern (z. B. `10042` statt `10024` oder `50676` statt `50667`) mit höchster Konfidenz.
   - **🔀 Zifferntausch**: Erkennt auch nicht benachbarte Vertauschungen (z. B. `15342` statt `12345`).
   - **✏️ Tippfehler / Einzelfehler**: Erkennt abweichende Einzelziffern (z. B. `KND-80499` statt `KND-80490`).
   - **➕➖ Ziffer zu viel oder zu wenig**: Erkennt fehlende oder überflüssige Stellen.
   - **❌ Nicht existent**: Markiert Werte, die in den Stammdaten überhaupt nicht existieren.
   - **⚪ Optionale Spalten (z. B. Leistungen & Ressourcen)**: Wenn ein Feld leer ist, wird es **nicht als Fehler** gewertet. Nur wenn eine Nummer eingetragen ist, wird diese auf Richtigkeit und Zahlendreher geprüft. Jede Spalte verfügt hierzu über eine Checkbox *„Leere Zellen erlauben (optional)“*.
   - **🔢 Ressourcen (Benutzerdefiniertes Format `0000`, ersetzt `Standard`)**: In der Spalte für die Ressourcennummer (z. B. Spalte „C“ *„Resourcen Nummer“*) wird beim Speichern der Typ *„Standard“* automatisch durch das benutzerdefinierte Excel-Zahlenformat **`0000`** ersetzt. Dadurch werden 1-, 2- oder 3-stellige Zahlen (z. B. `45`, `120`) in Microsoft Excel immer sauber 4-stellig mit führenden Nullen (`0045`, `0120`) dargestellt und sind optimal lesbar.
   - **⚡ Leistungen (dürfen 2-stellig sein)**: Leistungsnummern können problemlos 2-stellig sein (z. B. `12`, `45`, `78`) und werden nicht künstlich aufgefüllt. Zahlendreher (z. B. `54` statt `45`) werden dennoch zuverlässig erkannt.
   - **⚠️ Pflichtfelder**: Nur wenn eine Spalte ausdrücklich als Pflichtfeld definiert wird (Checkbox *„Leere Zellen erlauben“* deaktiviert), werden leere Zellen beanstandet.

2. **Fokussierte Fehleranzeige (keine langen Listen)**:
   - Nach der Prüfung werden **nur die beanstandeten Fehler** angezeigt (die fehlerfreien Zeilen werden ausgeblendet, können bei Bedarf aber über den Filter eingeblendet werden).
   - Zu jedem Fehler sehen Sie sofort: **Excel-Zelle** (z. B. `A5`, `B6`), **Spalte**, **Zeile**, **Originalwert**, **Diagnose** und den passenden **Korrekturvorschlag / Alternativen**.

3. **General-Button & Automatische Speicherung im Download-Ordner**:
   - **🚀 GENERAL-BUTTON: Datei erstellen & Änderungen übernehmen**:
     - Erstellt die fertige, korrigierte Excel-Datei, die **im Aufbau, Layout und Design exakt 1:1 dem geprüften Original entspricht**.
     - **Sicherheit für Fehler & unbekannte Nummern**: Wenn ein Fehler erkannt wird oder eine Nummer nicht in der Referenzdatei vorkommt, wird sie **niemals automatisch überschrieben**, sondern bleibt exakt so wie im Original erhalten!
     - Es werden nur Ihre manuellen Eingaben/Auswahlen aus der Maske sowie die 4-stellige Formatierung von Ressourcen übernommen.
     - Die fertige Datei wird **automatisch direkt in Ihren PC-Download-Ordner** (`Downloads`) gespeichert!
     - In der Maske erscheint ein grüner Erfolgs-Banner mit dem genauen Speicherpfad und einem Button **`📂 Download-Ordner öffnen`**, um die Datei direkt im Windows Explorer zu sehen.
   - **✍️ Manuelle Eingabe**: In der Maske können Sie vorgeschlagene Alternativen per Klick auf `✓ Übernehmen` bestätigen oder beliebige Nummern im Eingabefeld manuell eintragen.
   - **⚡ Nur Zahlendreher korrigieren**: Über diesen separaten Button können Sie bei Bedarf ausschließlich eindeutige Zahlendreher automatisch anpassen lassen.
   - **↩️ Rückgängig**: Alle Korrekturen können bei Bedarf zurückgesetzt werden.

4. **Modernes Drag & Drop**:
   - Einfaches Hineinziehen der **1. Referenzdatei** und der **2. Zu prüfenden Datei** in die markierten Bereiche.
   - Automatische Erkennung und Zuordnung gleicher Spaltennamen.
   - Auswahl beliebiger Arbeitsblätter bei mehrseitigen Excel-Mappen.

5. **Export**:
   - **💾 Korrigierte Excel herunterladen (.xlsx)**: Speichert Ihre Datei im Original-Layout mit den vorgenommenen Korrekturen.
   - **📊 Detaillierter Prüfbericht (.xlsx)**: Exportiert eine übersichtliche Fehlerdokumentation.

---

## 🚀 Schnellstart

### Methode 1: Per Doppelklick (Empfohlen)
Klicken Sie einfach doppelt auf:
```text
Start_Excel_Prüfer.bat
```
Das Programm öffnet sich automatisch als eigenständiges Desktop-Fenster (oder in Ihrem bevorzugten Browser).

### Methode 2: Über die Kommandozeile
```powershell
python app.py
```
Oder für den direkten Browser-Modus:
```powershell
python app.py --browser
```

---

## 🧪 Sofort-Test mit Beispieldaten
Klicken Sie in der oberen rechten Ecke des Programms auf **`✨ Demo-Dateien laden`**.
- Die enthaltenen Beispieldateien (`Beispiel_Referenz_Stammdaten.xlsx` und `Beispiel_Zu_Pruefen.xlsx`) werden sofort geladen.
- Klicken Sie auf **`🚀 Prüfung starten`**.
- Sie sehen sofort alle erkannten Zahlendreher, Tippfehler und fehlenden Werte in der Übersicht!
- Testen Sie den **General-Button**, um die korrigierte Excel-Datei sofort in Ihren Download-Ordner zu exportieren.

---

## ⚙️ Technische Details
- **Backend**: Python 3.13, `openpyxl`, `pandas`, `bottle`
- **Algorithmus**: Damerau-Levenshtein-Distanz mit O(1) Anagramm-Hashing für blitzschnelle Suche auch bei zehntausenden Zeilen
- **Frontend**: HTML5, CSS3, JavaScript (ohne externe Internet-Abhängigkeiten, funktioniert 100% offline)

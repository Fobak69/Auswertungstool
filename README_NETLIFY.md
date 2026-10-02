# 🚀 Excel-Prüfer & Korrektur-Assistent (Netlify Web-Version)

Diese Version läuft **zu 100 % im Browser** („Over the Air“) – ohne Python, ohne Installation und auf jedem Gerät (PC, Mac, Laptop, iPad, Tablet, Smartphone).

---

## ⚡ In 30 Sekunden auf Netlify veröffentlichen (Kostenlos)

1. Öffnen Sie die Website: **[https://app.netlify.com/drop](https://app.netlify.com/drop)**
2. Ziehen Sie diesen gesamten Ordner **`netlify`** einfach mit der Maus in das gestrichelte Feld auf der Netlify-Webseite hinein.
3. **Fertig!** Netlify gibt Ihnen sofort eine weltweite Internet-Adresse (z. B. `https://excel-pruefer-xyz.netlify.app`), die Sie anklicken und an Kollegen weitergeben können!

---

## 🔒 Datenschutz & Sicherheit
- **100 % lokal im Browser (Client-Side)**: Ihre Excel-Dateien werden **niemals** auf einen Server hochgeladen oder gespeichert.
- Die gesamte Prüfung, Fehlerdiagnose und Formatierung (Spalte C auf `0000`, Zahlendreher-Erkennung etc.) findet ausschließlich im Arbeitsspeicher Ihres eigenen Browsers statt.
- Vollständig DSGVO-konform.

---

## 🌟 Enthaltene Funktionen
- **Zahlendreher & Tippfehler**: Blitzschnelle Erkennung per Anagramm-Hashing und Damerau-Levenshtein.
- **Benutzerdefiniertes Format `0000`**: Spalte C (*„Resourcen Nummer“*) wird automatisch mit `0000` formatiert (Typ „Standard“ wird ersetzt, Zahlen wie `45` und `120` werden als `0045` und `0120` lesbar).
- **Leistungen**: Dürfen 2-stellig sein und leere Zeilen enthalten (kein Fehler bei leeren Zellen).
- **General-Button**: 
  - Belässt nicht-existente Nummern und unbestätigte Fehler im Originalzustand.
  - Übernimmt alle manuellen Eingaben und formatierte Ressourcen.
  - Baut die Datei im 1:1 Original-Layout auf und lädt sie direkt über den Browser herunter.
- **✨ Demo-Dateien laden**: 1-Klick-Button zum sofortigen Ausprobieren mit Beispieldaten direkt im Browser.

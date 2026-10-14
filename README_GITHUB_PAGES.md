# Anleitung: Excel-Prüfer kostenlos über GitHub Pages hosten

Da die Anwendung zu **100% im Browser (clientseitig)** läuft und keinen Server benötigt, kann sie **völlig kostenlos, dauerhaft und sicher über GitHub Pages** gehostet werden.

Ihre Firmendaten und Excel-Dateien bleiben dabei zu 100% geschützt: Es werden keinerlei Daten an fremde Server gesendet – die gesamte Berechnung und Speicherung (Mitarbeiterliste) findet lokal im Browser des Nutzers statt.

---

## Variante 1: Über dieses bestehende Repository (Ordner `/docs`) – Empfohlen!

Der Ordner `docs/` enthält bereits alle fertigen Web-Dateien (`index.html`, `app.js`, `style.css`, `engine.js`, `libs/`).

1. **Repository zu GitHub hochladen** (z. B. mit Git oder GitHub Desktop):
   ```bash
   git init
   git add .
   git commit -m "Excel-Prüfer mit Arbeitszeiten & Gesamttabellen-Editor"
   git branch -M main
   git remote add origin https://github.com/IHR-BENUTZERNAME/IHR-REPO-NAME.git
   git push -u origin main
   ```

2. **GitHub Pages aktivieren:**
   - Öffnen Sie Ihr Repository auf **github.com**.
   - Klicken Sie oben rechts auf den Reiter **⚙️ Settings** (Einstellungen).
   - Klicken Sie links im Menü auf **Pages**.
   - Unter **Build and deployment** / **Source**:
     - Wählen Sie: **Deploy from a branch**
     - Branch: **main**
     - Ordner: **/docs** (wichtig!)
   - Klicken Sie auf **Save**.

3. **Fertig!**
   - Nach ca. 30–60 Sekunden ist Ihre Web-App weltweit erreichbar unter:
   - `https://IHR-BENUTZERNAME.github.io/IHR-REPO-NAME/`

---

## Variante 2: Als eigenes, schlankes Web-Repository

Wenn Sie nur die Web-App (ohne Python-Dateien) in einem separaten GitHub-Projekt veröffentlichen möchten:

1. Erstellen Sie auf GitHub ein neues Repository (z. B. `excel-pruefer-web`).
2. Laden Sie einfach den Inhalt des Ordners `netlify/` bzw. `docs/` in die oberste Ebene (Root `/`) dieses neuen Repositories hoch.
3. Gehen Sie in diesem Repository auf **Settings** -> **Pages**.
4. Wählen Sie Branch **main** und Ordner **/** (Root) -> **Save**.
5. Die Seite ist sofort unter `https://IHR-BENUTZERNAME.github.io/excel-pruefer-web/` erreichbar.

---

## Vorteile des Hostings über GitHub Pages:
- ✅ **100% kostenlos** – dauerhaft ohne Hosting-Kosten.
- ✅ **Automatische HTTPS-Verschlüsselung** (grünes Schloss).
- ✅ **Keine Serverwartung nötig** – keine Updates, keine Abstürze.
- ✅ **Volle Privatsphäre** – alle Excel-Daten und Mitarbeiter-Stammdaten bleiben auf dem jeweiligen PC/Browser des Mitarbeiters.
- ✅ **Funktioniert auf jedem Gerät** – PC, Laptop, Mac, iPad oder Tablet.

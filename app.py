# -*- coding: utf-8 -*-
"""
Excel-Prüfer & Korrektur-Assistent - Web-Server & Desktop-App Launcher
Startet die Anwendung als modernes Desktop-Fenster oder im Web-Browser.
"""

import argparse
import io
import json
import os
import socket
import sys
import threading
import time
import webbrowser

from bottle import Bottle, request, response, static_file, HTTPResponse
from comparator import ExcelInspectionEngine

# Flask/Bottle Instanz
app = Bottle()
engine = ExcelInspectionEngine()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "netlify") if os.path.exists(os.path.join(BASE_DIR, "netlify")) else os.path.join(BASE_DIR, "static")


@app.route("/")
def index():
    return static_file("index.html", root=STATIC_DIR)


@app.route("/static/<filepath:path>")
def server_static(filepath):
    return static_file(filepath, root=STATIC_DIR)


@app.route("/<filepath:path>")
def server_root_files(filepath):
    if filepath.startswith("api/"):
        return HTTPResponse(status=404)
    return static_file(filepath, root=STATIC_DIR)


@app.route("/api/load-demo", method="POST")
def load_demo():
    """Lädt die vorbereiteten Demo-Dateien direkt für einen Schnelltest."""
    try:
        ref_path = os.path.join(BASE_DIR, "Beispiel_Referenz_Stammdaten.xlsx")
        tgt_path = os.path.join(BASE_DIR, "Beispiel_Zu_Pruefen.xlsx")

        if not os.path.exists(ref_path) or not os.path.exists(tgt_path):
            from demo_data import generate_demo_files
            generate_demo_files()

        with open(ref_path, "rb") as f:
            engine.load_reference(f.read(), "Beispiel_Referenz_Stammdaten.xlsx")

        with open(tgt_path, "rb") as f:
            engine.load_target(f.read(), "Beispiel_Zu_Pruefen.xlsx")

        ref_sheet = engine.ref_data["sheet_names"][0]
        tgt_sheet = engine.target_data["sheet_names"][0]
        mapping = engine.auto_map_columns(ref_sheet, tgt_sheet)

        return {
            "status": "success",
            "message": "Beispieldaten erfolgreich geladen!",
            "ref_data": {
                "filename": engine.ref_filename,
                "sheets": engine.ref_data["sheet_names"]
            },
            "target_data": {
                "filename": engine.target_filename,
                "sheets": engine.target_data["sheet_names"]
            },
            "mapping": mapping,
            "ref_sheet": ref_sheet,
            "target_sheet": tgt_sheet
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": f"Fehler beim Laden der Demo: {str(e)}"}


@app.route("/api/upload/reference", method="POST")
def upload_reference():
    """Nimmt die hochgeladene Referenzdatei entgegen."""
    try:
        upload = request.files.get("file")
        if not upload:
            response.status = 400
            return {"status": "error", "message": "Keine Datei ausgewählt."}

        filename = upload.raw_filename or "Referenz.xlsx"
        file_bytes = upload.file.read()
        
        ref_info = engine.load_reference(file_bytes, filename)
        
        # Falls bereits Ziel-Datei existiert, Mapping berechnen
        mapping = []
        if engine.target_data and engine.target_data.get("sheet_names"):
            tgt_sheet = engine.target_data["sheet_names"][0]
            ref_sheet = ref_info["sheet_names"][0]
            mapping = engine.auto_map_columns(ref_sheet, tgt_sheet)

        return {
            "status": "success",
            "filename": filename,
            "sheets": ref_info["sheet_names"],
            "mapping": mapping
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": f"Fehler beim Lesen der Referenzdatei: {str(e)}"}


@app.route("/api/upload/target", method="POST")
def upload_target():
    """Nimmt die zu prüfende Datei entgegen."""
    try:
        upload = request.files.get("file")
        if not upload:
            response.status = 400
            return {"status": "error", "message": "Keine Datei ausgewählt."}

        filename = upload.raw_filename or "Pruefdatei.xlsx"
        file_bytes = upload.file.read()

        tgt_info = engine.load_target(file_bytes, filename)

        # Mapping berechnen, falls Referenzdatei bereits da ist
        mapping = []
        if engine.ref_data and engine.ref_data.get("sheet_names"):
            ref_sheet = engine.ref_data["sheet_names"][0]
            tgt_sheet = tgt_info["sheet_names"][0]
            mapping = engine.auto_map_columns(ref_sheet, tgt_sheet)

        return {
            "status": "success",
            "filename": filename,
            "sheets": tgt_info["sheet_names"],
            "mapping": mapping
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": f"Fehler beim Lesen der Prüfdatei: {str(e)}"}


@app.route("/api/map-columns", method="POST")
def map_columns():
    """Berechnet Spalten-Mapping für spezifische Blätter."""
    try:
        data = request.json or {}
        ref_sheet = data.get("ref_sheet")
        tgt_sheet = data.get("target_sheet")

        if not ref_sheet or not tgt_sheet:
            response.status = 400
            return {"status": "error", "message": "Blattname fehlt."}

        mapping = engine.auto_map_columns(ref_sheet, tgt_sheet)
        ref_headers = [h["name"] for h in engine.ref_data["sheets"].get(ref_sheet, {}).get("headers", [])]

        return {
            "status": "success",
            "mapping": mapping,
            "ref_headers": ref_headers
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": str(e)}


@app.route("/api/run-check", method="POST")
def run_check():
    """Startet die Prüfung."""
    try:
        data = request.json or {}
        ref_sheet = data.get("ref_sheet")
        tgt_sheet = data.get("target_sheet")
        mapping = data.get("mapping", [])
        mode = data.get("mode", "pool")
        date_options = data.get("date_options", {})

        if not ref_sheet or not tgt_sheet:
            response.status = 400
            return {"status": "error", "message": "Bitte Blätter auswählen."}

        result = engine.run_check(ref_sheet, tgt_sheet, mapping, mode=mode, date_options=date_options)
        return {
            "status": "success",
            "stats": result["stats"],
            "results": result["results"]
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": f"Prüfung fehlgeschlagen: {str(e)}"}


@app.route("/api/correct/single", method="POST")
def correct_single():
    """Wendet eine einzelne manuelle oder ausgewählte Korrektur an (inkl. automatischer Übernahme gleicher Fehlerwerte)."""
    try:
        data = request.json or {}
        cell_id = data.get("cell_id")
        new_value = data.get("new_value")

        if not cell_id:
            response.status = 400
            return {"status": "error", "message": "Zell-ID fehlt."}

        res = engine.apply_single_correction(cell_id, new_value)
        return {
            "status": "success",
            "updated": res.get("item", {}),
            "updated_count": res.get("updated_count", 1),
            "updated_ids": res.get("updated_ids", [cell_id]),
            "total_corrected": len(engine.applied_corrections)
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": str(e)}


@app.route("/api/correct/batch", method="POST")
def correct_batch():
    """Wendet automatische Korrekturen auf Knopfdruck an."""
    try:
        data = request.json or {}
        mode = data.get("mode", "zahlendreher_only")

        batch_res = engine.apply_auto_corrections(mode=mode)
        return {
            "status": "success",
            "corrected_count": batch_res["corrected_count"],
            "total_applied": batch_res["total_applied"],
            "results": engine.results
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": str(e)}


@app.route("/api/correct/general-and-save", method="POST")
def general_and_save():
    """
    Wendet den General-Button an:
    1. Übernimmt vom Nutzer manuell eingegebene oder bestätigte Werte.
    2. Fehler oder nicht in der Referenzdatei existierende Nummern bleiben unverändert so wie sie sind!
    3. Erzeugt die korrigierte Excel-Datei im exakten Original-Layout (inkl. 4-stelliger Ressourcen).
    4. Speichert die Datei direkt im Download-Ordner des PCs (C:\\Users\\...\\Downloads).
    """
    try:
        data = request.json or {}
        # Manuelle Eingaben aus der Maske übernehmen (falls vorhanden)
        manual_list = data.get("manual_corrections", [])
        for item in manual_list:
            cell_id = item.get("cell_id")
            new_val = item.get("new_value")
            if cell_id and new_val is not None:
                engine.apply_single_correction(cell_id, str(new_val).strip())

        batch_res = engine.apply_auto_corrections(mode="general_button")
        excel_bytes = engine.export_corrected_excel(highlight_corrections=False)

        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(downloads_dir, exist_ok=True)

        base_name = os.path.splitext(engine.target_filename)[0]
        out_name = f"{base_name}_KORRIGIERT.xlsx"
        save_path = os.path.join(downloads_dir, out_name)

        # Falls die Datei bereits in Excel geöffnet ist (Windows PermissionError / Lock),
        # automatisch eine freie Variante wie "_KORRIGIERT_(1).xlsx" wählen
        final_save_path = save_path
        final_out_name = out_name
        counter = 1
        written = False
        while not written and counter <= 100:
            try:
                with open(final_save_path, "wb") as f:
                    f.write(excel_bytes)
                written = True
            except (PermissionError, OSError):
                final_out_name = f"{base_name}_KORRIGIERT_({counter}).xlsx"
                final_save_path = os.path.join(downloads_dir, final_out_name)
                counter += 1

        if not written:
            raise PermissionError(f"Konnte Datei '{out_name}' nicht im Download-Ordner speichern, da sie in einem anderen Programm geöffnet ist.")

        return {
            "status": "success",
            "corrected_count": batch_res["corrected_count"],
            "total_applied": batch_res["total_applied"],
            "results": engine.results,
            "filename": final_out_name,
            "saved_path": final_save_path,
            "download_url": "/api/export/corrected"
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": f"Fehler beim Speichern der Korrektur: {str(e)}"}


@app.route("/api/open-downloads", method="POST")
def open_downloads():
    """Öffnet den Windows Download-Ordner im Explorer."""
    try:
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        if hasattr(os, "startfile"):
            os.startfile(downloads_dir)
        else:
            import subprocess
            subprocess.Popen(["explorer", downloads_dir])
        return {"status": "success"}
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": str(e)}


@app.route("/api/correct/reset", method="POST")
def correct_reset():
    """Setzt alle Korrekturen zurück."""
    try:
        res = engine.reset_corrections()
        return {
            "status": "success",
            "results": engine.results
        }
    except Exception as e:
        response.status = 500
        return {"status": "error", "message": str(e)}


@app.route("/api/export/corrected")
def export_corrected():
    """Exportiert die korrigierte Excel-Datei."""
    try:
        highlight = request.query.get("highlight", "true").lower() == "true"
        excel_bytes = engine.export_corrected_excel(highlight_corrections=highlight)

        base_name = os.path.splitext(engine.target_filename)[0]
        out_name = f"{base_name}_KORRIGIERT.xlsx"

        res = HTTPResponse(
            body=excel_bytes,
            status=200,
            headers={
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Disposition": f'attachment; filename="{out_name}"'
            }
        )
        return res
    except Exception as e:
        response.status = 500
        return f"Fehler beim Exportieren: {str(e)}"


@app.route("/api/export/report")
def export_report():
    """Exportiert den Prüfbericht als Excel-Datei."""
    try:
        report_bytes = engine.export_report_excel()
        base_name = os.path.splitext(engine.target_filename)[0]
        out_name = f"Pruefbericht_{base_name}.xlsx"

        res = HTTPResponse(
            body=report_bytes,
            status=200,
            headers={
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Disposition": f'attachment; filename="{out_name}"'
            }
        )
        return res
    except Exception as e:
        response.status = 500
        return f"Fehler beim Erstellen des Berichts: {str(e)}"


def get_local_ip():
    """Ermittelt die IP-Adresse dieses Computers im lokalen Netzwerk (LAN/WLAN)."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"


def find_free_port(start_port=5500, max_attempts=50):
    for p in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start_port


def main():
    parser = argparse.ArgumentParser(description="Excel-Prüfer & Korrektur-Assistent")
    parser.add_argument("--browser", action="store_true", help="Im Standard-Webbrowser statt im Desktop-Fenster öffnen")
    parser.add_argument("--port", type=int, default=0, help="Port für den Webserver")
    args = parser.parse_args()

    port = args.port if args.port > 0 else find_free_port(5500)
    local_ip = get_local_ip()
    local_url = f"http://127.0.0.1:{port}"
    url = local_url
    network_url = f"http://{local_ip}:{port}" if local_ip != "127.0.0.1" else None

    # Server für das lokale Netzwerk (0.0.0.0) im Hintergrund starten
    server_thread = threading.Thread(
        target=lambda: app.run(host="0.0.0.0", port=port, quiet=True),
        daemon=True
    )
    server_thread.start()
    time.sleep(0.5)

    print("=" * 68)
    print("   📊 EXCEL-PRÜFER & KORREKTUR-ASSISTENT")
    print(f"   Lokal auf diesem PC:         {local_url}")
    if network_url:
        print(f"   Von anderen PCs im Netzwerk: {network_url}")
    print("=" * 68)

    if args.browser:
        print(f"Öffne Browser unter: {url}")
        webbrowser.open(url)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("Beende Server...")
            sys.exit(0)
    else:
        # Versuche pywebview für ein echtes Desktop-App-Fenster
        try:
            import webview
            print("Starte Desktop-App-Fenster...")
            window = webview.create_window(
                title="Excel-Prüfer & Korrektur-Assistent",
                url=url,
                width=1280,
                height=850,
                min_size=(960, 640),
                resizable=True
            )
            webview.start()
        except Exception as e:
            print(f"Hinweis: Desktop-Fenster konnte nicht geöffnet werden ({e}).")
            print(f"Öffne stattdessen Ihren Standard-Browser: {url}")
            webbrowser.open(url)
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                sys.exit(0)


if __name__ == "__main__":
    main()

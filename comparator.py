# -*- coding: utf-8 -*-
"""
Excel-Prüfer & Korrektur-Assistent - Vergleichs- & Analysemodul
Intelligente Erkennung von Zahlendrehern, Tippfehlern und nicht existenten Werten.
"""

import collections
import datetime
import io
import re
from typing import Dict, List, Any, Optional, Tuple
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import openpyxl.utils.datetime as openpyxl_datetime


def parse_date_value(val: Any) -> Optional[datetime.date]:
    """Konvertiert verschiedene Datumsrepräsentationen (datetime, date, serial number, string) in ein datetime.date."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, (int, float)):
        # Excel date serial number (z.B. 20000 bis 90000 entspricht den Jahren ~1954 bis ~2146)
        if 20000 <= val <= 90000:
            try:
                dt = openpyxl_datetime.from_excel(val)
                if isinstance(dt, datetime.datetime):
                    return dt.date()
                if isinstance(dt, datetime.date):
                    return dt
            except Exception:
                pass
    s = str(val).strip()
    if not s or s.lower() in ("[object object]", "object objekt", "none"):
        return None
    # Falls s ein numerischer String wie '45561' ist
    try:
        f = float(s)
        if 20000 <= f <= 90000:
            dt = openpyxl_datetime.from_excel(f)
            if isinstance(dt, datetime.datetime):
                return dt.date()
            if isinstance(dt, datetime.date):
                return dt
    except (ValueError, OverflowError):
        pass

    # Gängige Datumsformate prüfen
    date_formats = [
        "%d.%m.%Y", "%d.%m.%y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d/%m/%Y", "%d/%m/%y",
        "%d-%m-%Y", "%d-%m-%y",
        "%Y.%m.%d",
        "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"
    ]
    for fmt in date_formats:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def normalize_cell_value(val: Any) -> str:
    """Normalisiert einen Zellwert für den String- und Zahlenabgleich."""
    if val is None:
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%d.%m.%Y")
    if isinstance(val, bool):
        return "WAHR" if val else "FALSCH"
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        # Runden auf übliche Nachkommastellen bei Float-Ungenauigkeiten
        formatted = f"{val:.8f}".rstrip("0").rstrip(".")
        return formatted
    
    s = str(val).strip()
    if s in ("[object Object]", "object Objekt", "(object Objekt)"):
        return ""
    # Falls der String wie eine Ganzzahl-Float aussieht (z.B. "12345.0")
    try:
        f = float(s)
        if f.is_integer() and "." in s and s.endswith(".0"):
            return str(int(f))
    except (ValueError, OverflowError):
        pass
    return s


def pad_number(val: Any, min_digits: int = 4) -> str:
    """
    Füllt eine Zahl auf mindestens min_digits Stellen mit führenden Nullen auf.
    Z.B.:
      '45'   -> '0045'
      '123'  -> '0123'
      '1234' -> '1234'
      'L-45' -> 'L-0045'
    """
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""

    # Falls Float wie '45.0' vorliegt
    try:
        f = float(s)
        if f.is_integer() and "." in s and s.endswith(".0"):
            s = str(int(f))
    except (ValueError, OverflowError):
        pass

    # Reine Ziffernfolge
    if s.isdigit():
        return s.zfill(min_digits)

    # Präfix + Ziffern (z. B. 'L-45' -> 'L-0045')
    m = re.match(r"^([^\d]*)(\d+)$", s)
    if m:
        prefix, num = m.groups()
        return prefix + num.zfill(min_digits)

    return s


def classify_difference(target: str, ref: str) -> Tuple[str, float, str]:
    """
    Analysiert den Unterschied zwischen Prüfwert und Referenzwert.
    Gibt (Fehlertyp, Konfidenz-Score 0..1, Detail-Beschreibung) zurück.
    """
    if target == ref:
        return ("OK", 1.0, "Exakte Übereinstimmung")
    
    l_target = len(target)
    l_ref = len(ref)

    # 1. Gleiche Länge: Zahlendreher, Zifferntausch oder Tippfehler
    if l_target == l_ref and l_target > 0:
        diffs = [(i, target[i], ref[i]) for i in range(l_target) if target[i] != ref[i]]
        
        # Genau 2 Zeichen abweichend
        if len(diffs) == 2:
            i, c_t1, c_r1 = diffs[0]
            j, c_t2, c_r2 = diffs[1]
            if c_t1 == c_r2 and c_t2 == c_r1:
                if j == i + 1:
                    return (
                        "ZAHLENDREHER",
                        0.98,
                        f"Zahlendreher an Pos. {i+1}-{j+1} ('{c_t1}{c_t2}' statt '{c_r1}{c_r2}')"
                    )
                else:
                    return (
                        "ZIFFERENTAUSCH",
                        0.93,
                        f"Zifferntausch an Pos. {i+1} & {j+1} ('{c_t1}' und '{c_t2}' vertauscht)"
                    )
            else:
                return ("ABWEICHUNG_2", 0.72, f"2 Zeichen unterschiedlich (Pos. {i+1}, {j+1})")
        
        # Genau 1 Zeichen abweichend (Tippfehler)
        elif len(diffs) == 1:
            i, c_t, c_r = diffs[0]
            return (
                "TIPPFEHLER",
                0.88,
                f"Tippfehler an Pos. {i+1} ('{c_t}' statt '{c_r}')"
            )
        
        # Mehr als 2 Zeichen abweichend, aber selbe Zeichenmenge (Permutation)
        elif sorted(target) == sorted(ref):
            return ("PERMUTATION", 0.80, "Ziffern vollständig durcheinander geraten")

    # 2. Längenunterschied um genau 1 Zeichen (Ziffer fehlt oder zu viel)
    elif abs(l_target - l_ref) == 1:
        if l_target > l_ref:
            # Prüfwert hat 1 Zeichen zu viel
            for i in range(l_target):
                if target[:i] + target[i+1:] == ref:
                    return (
                        "ZIFFER_ZUVIEL",
                        0.85,
                        f"Überflüssiges Zeichen an Pos. {i+1} ('{target[i]}')"
                    )
        else:
            # Prüfwert fehlt 1 Zeichen
            for i in range(l_ref):
                if ref[:i] + ref[i+1:] == target:
                    return (
                        "ZIFFER_FEHLT",
                        0.85,
                        f"Fehlendes Zeichen an Pos. {i+1} (sollte '{ref[i]}' sein)"
                    )

    # 3. Damerau-Levenshtein für sonstige Nahe-Treffer
    dist = damerau_levenshtein(target, ref)
    max_len = max(l_target, l_ref)
    if max_len > 0:
        sim = 1.0 - (dist / max_len)
        if dist <= 2 and sim >= 0.65:
            return ("AEHNLICH", round(sim, 2), f"Ähnlicher Wert (Distanz {dist})")

    return ("UNBEKANNT", 0.0, "Keine Übereinstimmung")


def damerau_levenshtein(s1: str, s2: str) -> int:
    """Berechnet die Damerau-Levenshtein-Distanz (inklusive Buchstabendreher)."""
    len1, len2 = len(s1), len(s2)
    d = [[0] * (len2 + 1) for _ in range(len1 + 1)]
    for i in range(len1 + 1):
        d[i][0] = i
    for j in range(len2 + 1):
        d[0][j] = j
    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            cost = 0 if s1[i - 1] == s2[j - 1] else 1
            d[i][j] = min(
                d[i - 1][j] + 1,       # Löschen
                d[i][j - 1] + 1,       # Einfügen
                d[i - 1][j - 1] + cost # Ersetzen
            )
            if i > 1 and j > 1 and s1[i - 1] == s2[j - 2] and s1[i - 2] == s2[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + 1)  # Vertauschung
    return d[len1][len2]


class ColumnIndex:
    """Effizienter Index für Referenzspalten mit O(1) Lookups für Zahlendreher."""

    def __init__(self, values: List[str]):
        # Alle eindeutigen nicht-leeren Werte
        self.exact_set = set(v for v in values if v)
        self.by_len = collections.defaultdict(list)
        self.by_chars = collections.defaultdict(list)

        for v in self.exact_set:
            self.by_len[len(v)].append(v)
            chars_key = "".join(sorted(v))
            self.by_chars[chars_key].append(v)

    def contains(self, val: str) -> bool:
        return val in self.exact_set

    def find_candidates(self, target: str, max_candidates: int = 5) -> List[Dict[str, Any]]:
        """Findet die besten Korrektur-Kandidaten aus der Referenz."""
        if not target or target in self.exact_set:
            return []

        candidates = []
        target_len = len(target)
        sorted_chars = "".join(sorted(target))

        # 1. Sofort-Check: Anagramme / Zahlendreher (O(1) Hash-Lookup!)
        if sorted_chars in self.by_chars:
            for ref_val in self.by_chars[sorted_chars]:
                err_type, score, detail = classify_difference(target, ref_val)
                candidates.append({
                    "value": ref_val,
                    "score": score,
                    "type": err_type,
                    "detail": detail
                })

        # 2. Check Kandidaten mit Längen target_len - 1, target_len, target_len + 1
        checked_lengths = [target_len]
        if target_len > 1:
            checked_lengths.append(target_len - 1)
        checked_lengths.append(target_len + 1)

        already_seen = set(c["value"] for c in candidates)

        for l in checked_lengths:
            for ref_val in self.by_len.get(l, []):
                if ref_val in already_seen:
                    continue
                err_type, score, detail = classify_difference(target, ref_val)
                if score >= 0.65:
                    candidates.append({
                        "value": ref_val,
                        "score": score,
                        "type": err_type,
                        "detail": detail
                    })
                    already_seen.add(ref_val)

        # 3. Fallback: Immer die am nächsten liegenden Alternativen vorschlagen (selbst bei komplett unbekannten Nummern)
        if len(candidates) < max_candidates and self.exact_set:
            remaining_vals = [v for v in self.exact_set if v not in already_seen]
            if remaining_vals:
                t_is_num = target.isdigit()
                t_num = int(target) if t_is_num else 0

                scored_fallback = []
                for r_val in remaining_vals:
                    dist = damerau_levenshtein(target, r_val)
                    num_diff = abs(t_num - int(r_val)) if (t_is_num and r_val.isdigit()) else 99999999
                    max_len = max(len(target), len(r_val))
                    sim = max(0.1, round(1.0 - (dist / max(1, max_len)), 2))
                    scored_fallback.append((dist, num_diff, r_val, sim))

                scored_fallback.sort(key=lambda x: (x[0], x[1]))
                for dist, _, r_val, sim in scored_fallback:
                    if len(candidates) >= max_candidates:
                        break
                    candidates.append({
                        "value": r_val,
                        "score": sim,
                        "type": "NAECHSTE_ALTERNATIVE",
                        "detail": f"Nächstgelegene Alternative in Stammdaten (Distanz: {dist})"
                    })
                    already_seen.add(r_val)

        # Nach Score absteigend sortieren
        candidates.sort(key=lambda x: x["score"], reverse=True)
        return candidates[:max_candidates]


class ExcelInspectionEngine:
    """Verwaltet das Laden, Prüfen, Korrigieren und Exportieren von Excel-Dateien."""

    def __init__(self):
        self.ref_data: Dict[str, Any] = {}
        self.target_data: Dict[str, Any] = {}
        self.results: List[Dict[str, Any]] = []
        self.target_workbook_bytes: Optional[bytes] = None
        self.ref_workbook_bytes: Optional[bytes] = None
        self.target_filename: str = "Pruefdatei.xlsx"
        self.ref_filename: str = "Referenzdatei.xlsx"
        self.applied_corrections: Dict[str, str] = {}  # cell_key ("row_col") -> new_value

    def load_reference(self, file_bytes: bytes, filename: str = "Referenz.xlsx") -> Dict[str, Any]:
        """Lädt die Referenzdatei und extrahiert Blätter und Spalten."""
        self.ref_workbook_bytes = file_bytes
        self.ref_filename = filename
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

        sheets = {}
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            # Headers in Zeile 1
            headers = []
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=col_idx).value
                header_title = str(val).strip() if val is not None else f"Spalte {get_column_letter(col_idx)}"
                headers.append({
                    "col_idx": col_idx,
                    "letter": get_column_letter(col_idx),
                    "name": header_title
                })

            # Spaltenwerte erfassen
            columns_data = {}
            for h in headers:
                col_idx = h["col_idx"]
                vals = []
                for row_idx in range(2, ws.max_row + 1):
                    c_val = ws.cell(row=row_idx, column=col_idx).value
                    norm_val = normalize_cell_value(c_val)
                    if norm_val:
                        vals.append(norm_val)
                columns_data[h["name"]] = {
                    "col_idx": col_idx,
                    "letter": h["letter"],
                    "unique_count": len(set(vals)),
                    "total_count": len(vals),
                    "values": vals,
                    "sample": vals[:5]
                }

            sheets[sheetname] = {
                "headers": headers,
                "columns": columns_data,
                "row_count": ws.max_row
            }

        self.ref_data = {
            "filename": filename,
            "sheets": sheets,
            "sheet_names": wb.sheetnames
        }
        return self.ref_data

    def load_target(self, file_bytes: bytes, filename: str = "Pruefdatei.xlsx") -> Dict[str, Any]:
        """Lädt die zu prüfende Datei."""
        self.target_workbook_bytes = file_bytes
        self.target_filename = filename
        self.applied_corrections = {}
        self.results = []
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

        sheets = {}
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            headers = []
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=col_idx).value
                header_title = str(val).strip() if val is not None else f"Spalte {get_column_letter(col_idx)}"
                headers.append({
                    "col_idx": col_idx,
                    "letter": get_column_letter(col_idx),
                    "name": header_title
                })

            sheets[sheetname] = {
                "headers": headers,
                "row_count": ws.max_row,
                "col_count": ws.max_column
            }

        self.target_data = {
            "filename": filename,
            "sheets": sheets,
            "sheet_names": wb.sheetnames
        }
        return self.target_data

    def auto_map_columns(self, ref_sheet: str, target_sheet: str) -> List[Dict[str, Any]]:
        """Ermittelt automatisch passende Spalten zwischen Referenz- und Prüfblatt."""
        if not self.ref_data or not self.target_data:
            return []

        ref_cols = self.ref_data["sheets"].get(ref_sheet, {}).get("headers", [])
        tgt_cols = self.target_data["sheets"].get(target_sheet, {}).get("headers", [])

        mapping = []
        for tgt in tgt_cols:
            tgt_clean = re.sub(r"[^a-zA-Z0-9äöüÄÖÜß]", "", tgt["name"].lower())
            best_match = None
            for ref in ref_cols:
                ref_clean = re.sub(r"[^a-zA-Z0-9äöüÄÖÜß]", "", ref["name"].lower())
                if tgt_clean == ref_clean and tgt_clean:
                    best_match = ref["name"]
                    break
            
            # Falls kein Match über Header-Namen gefunden wurde, nach gleicher Position/Buchstabe suchen
            if not best_match:
                for ref in ref_cols:
                    if ref.get("col_idx") == tgt["col_idx"] or ref.get("letter", "").upper() == tgt["letter"].upper():
                        best_match = ref["name"]
                        break

            tgt_letter_upper = tgt["letter"].upper()
            is_default_col = (tgt_letter_upper in ("B", "C", "E")) or (tgt["col_idx"] in (2, 3, 5))
            is_ressource = (tgt_letter_upper == "C") or (tgt["col_idx"] == 3) or any(k in tgt["name"].lower() for k in ("ressource", "resource", "resour", "res-nr", "res_nr", "resnr"))
            is_leistung = (tgt_letter_upper == "E") or (tgt["col_idx"] == 5) or any(k in tgt["name"].lower() for k in ("leistung", "leist", "leist-nr", "leist_nr", "leistnr"))
            mapping.append({
                "target_col_idx": tgt["col_idx"],
                "target_col_letter": tgt["letter"],
                "target_col_name": tgt["name"],
                "ref_col_name": best_match or "",
                "selected": is_default_col,
                "allow_empty": is_leistung,  # Nur bei den Leistungen gesetzt und sonst nirgends
                "pad_to_4": is_ressource  # Ressourcen in Spalte C müssen immer 4-stellig sein
            })

        return mapping

    def run_check(self, ref_sheet: str, target_sheet: str, column_mapping: List[Dict[str, Any]], mode: str = "pool", date_options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Führt die vollständige Prüfung durch.
        mode='pool': Gültigkeitsprüfung (Zellwert muss in der Referenzspalte existieren).
        mode='row': Zeilenweiser 1:1 Abgleich (Zeile N Prüfdatei gegen Zeile N Referenzdatei).
        date_options: Konfiguration der Datumsprüfung für Spalte A.
        """
        if not self.target_workbook_bytes or not self.ref_workbook_bytes:
            raise ValueError("Referenz- oder Prüfdatei fehlt.")

        ref_wb = openpyxl.load_workbook(io.BytesIO(self.ref_workbook_bytes), data_only=True)
        tgt_wb = openpyxl.load_workbook(io.BytesIO(self.target_workbook_bytes), data_only=True)

        ws_ref = ref_wb[ref_sheet]
        ws_tgt = tgt_wb[target_sheet]
        self.last_target_sheet = target_sheet
        self.last_ref_sheet = ref_sheet
        self.last_mapping = column_mapping
        self.applied_corrections = {}

        # Referenz-Indizes für die ausgewählten Spalten aufbauen
        col_indices: Dict[str, ColumnIndex] = {}
        # Zeilenweise Referenzdaten falls mode == 'row'
        ref_rows_data: Dict[str, Dict[int, str]] = {}

        for m in column_mapping:
            if not m.get("selected") or not m.get("ref_col_name"):
                continue
            ref_col_name = m["ref_col_name"]
            
            # Finde Spaltenindex in Referenztabelle
            ref_col_idx = None
            for c in range(1, ws_ref.max_column + 1):
                val = ws_ref.cell(row=1, column=c).value
                if val and str(val).strip() == ref_col_name:
                    ref_col_idx = c
                    break

            pad_to_4 = m.get("pad_to_4", False) or (m.get("target_col_letter", "").upper() == "C") or (m.get("target_col_idx") == 3)
            if ref_col_idx:
                vals = []
                row_map = {}
                for r in range(2, ws_ref.max_row + 1):
                    raw_val = ws_ref.cell(row=r, column=ref_col_idx).value
                    norm_val = normalize_cell_value(raw_val)
                    if norm_val:
                        vals.append(norm_val)
                        if pad_to_4 or (norm_val.isdigit() and len(norm_val) <= 4):
                            padded_ref = pad_number(norm_val, 4)
                            vals.append(padded_ref)
                            row_map[r] = padded_ref
                            unpadded = norm_val.lstrip("0")
                            if unpadded:
                                vals.append(unpadded)
                        else:
                            row_map[r] = norm_val
                
                col_indices[m["target_col_name"]] = ColumnIndex(vals)
                ref_rows_data[m["target_col_name"]] = row_map

        # Prüfdatei analysieren
        results = []
        stats = {
            "total_cells": 0,
            "ok": 0,
            "zahlendreher": 0,
            "zifferntausch": 0,
            "tippfehler": 0,
            "ziffer_zuviel_zuwenig": 0,
            "nicht_existent": 0,
            "leer": 0,
            "datum_warnung": 0,
            "total_errors": 0
        }

        # Datumsprüfung für Spalte A vorbereiten (Standard: aktiv für letzte 3 Wochen)
        if date_options is None:
            date_options = {"enabled": True, "mode": "weeks", "weeks": 3, "allow_future": False}
        
        date_check_enabled = date_options.get("enabled", True)
        today = datetime.date.today()
        date_mode = date_options.get("mode", "weeks")

        if date_mode == "custom" and (date_options.get("start_date") or date_options.get("end_date")):
            s_str = date_options.get("start_date")
            e_str = date_options.get("end_date")
            min_date = parse_date_value(s_str) if s_str else (today - datetime.timedelta(days=21))
            max_date = parse_date_value(e_str) if e_str else today
            s_disp = min_date.strftime("%d.%m.%Y") if min_date else "?"
            e_disp = max_date.strftime("%d.%m.%Y") if max_date else "?"
            weeks_info = f"Zeitraum von {s_disp} bis {e_disp}"
        else:
            try:
                weeks = int(date_options.get("weeks", 3))
            except (ValueError, TypeError):
                weeks = 3
            min_date = today - datetime.timedelta(days=weeks * 7)
            allow_future = date_options.get("allow_future", False)
            max_date = today if not allow_future else None
            weeks_info = f"letzte {weeks} Wochen (ab {min_date.strftime('%d.%m.%Y')})"

        col_a_header = ws_tgt.cell(row=1, column=1).value
        col_a_name = str(col_a_header).strip() if col_a_header else "Datum"

        # Wir lesen alle Zeilen der Prüfdatei
        for r_idx in range(2, ws_tgt.max_row + 1):
            # 1. Datumsprüfung für Spalte A
            if date_check_enabled:
                cell_coord_a = f"A{r_idx}"
                cell_key_a = f"{r_idx}_1"
                raw_date_val = ws_tgt.cell(row=r_idx, column=1).value
                parsed_d = parse_date_value(raw_date_val)
                stats["total_cells"] += 1

                val_str_a = parsed_d.strftime("%d.%m.%Y") if parsed_d else normalize_cell_value(raw_date_val)

                if raw_date_val is None or str(raw_date_val).strip() == "":
                    stats["total_errors"] += 1
                    stats["leer"] += 1
                    stats["datum_warnung"] += 1
                    results.append({
                        "id": cell_key_a,
                        "row": r_idx,
                        "col_idx": 1,
                        "col_letter": "A",
                        "col_name": col_a_name,
                        "ref_col_name": "",
                        "cell": cell_coord_a,
                        "original_value": "",
                        "current_value": "",
                        "status": "DATUM_WARNUNG",
                        "status_label": "Datum eventuell falsch",
                        "badge_class": "badge-tippfehler",
                        "score": 0.0,
                        "suggestion": today.strftime("%d.%m.%Y"),
                        "detail": "Datum in Spalte A fehlt komplett",
                        "candidates": [{"value": today.strftime("%d.%m.%Y"), "detail": "Heutiges Datum"}],
                        "is_corrected": False
                    })
                elif parsed_d is None:
                    stats["total_errors"] += 1
                    stats["tippfehler"] += 1
                    stats["datum_warnung"] += 1
                    results.append({
                        "id": cell_key_a,
                        "row": r_idx,
                        "col_idx": 1,
                        "col_letter": "A",
                        "col_name": col_a_name,
                        "ref_col_name": "",
                        "cell": cell_coord_a,
                        "original_value": val_str_a,
                        "current_value": val_str_a,
                        "status": "DATUM_WARNUNG",
                        "status_label": "Datum eventuell falsch",
                        "badge_class": "badge-tippfehler",
                        "score": 0.0,
                        "suggestion": today.strftime("%d.%m.%Y"),
                        "detail": f"Ungültiges Datumsformat oder Zahl ('{val_str_a}')",
                        "candidates": [{"value": today.strftime("%d.%m.%Y"), "detail": "Heutiges Datum"}],
                        "is_corrected": False
                    })
                elif min_date and parsed_d < min_date:
                    stats["total_errors"] += 1
                    stats["tippfehler"] += 1
                    stats["datum_warnung"] += 1
                    days_diff = (today - parsed_d).days
                    results.append({
                        "id": cell_key_a,
                        "row": r_idx,
                        "col_idx": 1,
                        "col_letter": "A",
                        "col_name": col_a_name,
                        "ref_col_name": "",
                        "cell": cell_coord_a,
                        "original_value": val_str_a,
                        "current_value": val_str_a,
                        "status": "DATUM_WARNUNG",
                        "status_label": "Datum eventuell falsch",
                        "badge_class": "badge-tippfehler",
                        "score": 0.5,
                        "suggestion": today.strftime("%d.%m.%Y"),
                        "detail": f"Datum {val_str_a} liegt {days_diff} Tage zurück (erwartet: {weeks_info})",
                        "candidates": [{"value": today.strftime("%d.%m.%Y"), "detail": "Heutiges Datum"}],
                        "is_corrected": False
                    })
                elif max_date and parsed_d > max_date:
                    stats["total_errors"] += 1
                    stats["tippfehler"] += 1
                    stats["datum_warnung"] += 1
                    days_diff = (parsed_d - today).days
                    results.append({
                        "id": cell_key_a,
                        "row": r_idx,
                        "col_idx": 1,
                        "col_letter": "A",
                        "col_name": col_a_name,
                        "ref_col_name": "",
                        "cell": cell_coord_a,
                        "original_value": val_str_a,
                        "current_value": val_str_a,
                        "status": "DATUM_WARNUNG",
                        "status_label": "Datum eventuell falsch",
                        "badge_class": "badge-tippfehler",
                        "score": 0.5,
                        "suggestion": today.strftime("%d.%m.%Y"),
                        "detail": f"Datum {val_str_a} liegt in der Zukunft (+{days_diff} Tage)",
                        "candidates": [{"value": today.strftime("%d.%m.%Y"), "detail": "Heutiges Datum"}],
                        "is_corrected": False
                    })
                else:
                    stats["ok"] += 1
                    results.append({
                        "id": cell_key_a,
                        "row": r_idx,
                        "col_idx": 1,
                        "col_letter": "A",
                        "col_name": col_a_name,
                        "ref_col_name": "",
                        "cell": cell_coord_a,
                        "original_value": val_str_a,
                        "current_value": val_str_a,
                        "status": "OK",
                        "status_label": "Datum gültig",
                        "badge_class": "badge-ok",
                        "score": 1.0,
                        "suggestion": "",
                        "detail": f"Datum {val_str_a} liegt im gültigen Zeitraum ({weeks_info})",
                        "candidates": [],
                        "is_corrected": False
                    })

            # 2. Reguläre Spaltenprüfung für gemappte Spalten (B, C, E, etc.)
            for m in column_mapping:
                if not m.get("selected") or not m.get("ref_col_name"):
                    continue

                tgt_col_idx = m["target_col_idx"]
                if date_check_enabled and (tgt_col_idx == 1 or m.get("target_col_letter", "").upper() == "A"):
                    # Spalte A wurde bereits über die Datumsprüfung validiert
                    continue
                tgt_col_name = m["target_col_name"]
                tgt_letter = m["target_col_letter"]
                cell_coord = f"{tgt_letter}{r_idx}"
                cell_key = f"{r_idx}_{tgt_col_idx}"

                raw_val = ws_tgt.cell(row=r_idx, column=tgt_col_idx).value
                val_str = normalize_cell_value(raw_val)
                stats["total_cells"] += 1

                index = col_indices.get(tgt_col_name)
                if not index:
                    continue

                # Status ermitteln
                allow_empty = bool(m.get("allow_empty", False))
                if not val_str:
                    if allow_empty:
                        # Feld ist leer, aber als optional/erlaubt markiert (z.B. Leistungen, die nicht in jeder Zeile vorkommen)
                        stats["ok"] += 1
                        results.append({
                            "id": cell_key,
                            "row": r_idx,
                            "col_idx": tgt_col_idx,
                            "col_letter": tgt_letter,
                            "col_name": tgt_col_name,
                            "ref_col_name": m["ref_col_name"],
                            "cell": cell_coord,
                            "original_value": "",
                            "current_value": "",
                            "status": "OK_LEER",
                            "status_label": "Leer (erlaubt)",
                            "badge_class": "badge-leer-ok",
                            "score": 1.0,
                            "suggestion": "",
                            "detail": "Feld ist leer (als optional deklariert, kein Fehler)",
                            "candidates": [],
                            "is_corrected": False
                        })
                    else:
                        stats["leer"] += 1
                        stats["total_errors"] += 1
                        results.append({
                            "id": cell_key,
                            "row": r_idx,
                            "col_idx": tgt_col_idx,
                            "col_letter": tgt_letter,
                            "col_name": tgt_col_name,
                            "ref_col_name": m["ref_col_name"],
                            "cell": cell_coord,
                            "original_value": "",
                            "current_value": "",
                            "status": "LEER",
                            "status_label": "Pflichtfeld leer",
                            "badge_class": "badge-leer",
                            "score": 0.0,
                            "suggestion": "",
                            "detail": "Zelle ist leer, obwohl Pflichtfeld",
                            "candidates": [],
                            "is_corrected": False
                        })
                    continue

                pad_to_4 = m.get("pad_to_4", False) or (m.get("target_col_letter", "").upper() == "C") or (m.get("target_col_idx") == 3)
                if pad_to_4 and val_str:
                    target_to_check = pad_number(val_str, 4)
                    unpadded_target = val_str.lstrip("0")
                    needed_padding = (target_to_check != val_str)
                else:
                    target_to_check = val_str
                    unpadded_target = val_str
                    needed_padding = False

                # 1. Exakte Übereinstimmung
                if mode == "pool":
                    is_ok = index.contains(target_to_check) or index.contains(val_str) or (bool(unpadded_target) and index.contains(unpadded_target))
                else:
                    # Zeilenweiser Modus
                    expected_row_val = ref_rows_data.get(tgt_col_name, {}).get(r_idx, "")
                    expected_unpadded = expected_row_val.lstrip("0") if expected_row_val else ""
                    is_ok = (target_to_check == expected_row_val or val_str == expected_row_val or (bool(unpadded_target) and unpadded_target == expected_unpadded))

                if is_ok:
                    stats["ok"] += 1
                    if pad_to_4 and needed_padding:
                        # Automatisch für 4-stelligen Export mit Formatierung 0000 vormerken, aber NICHT als Fehler anzeigen:
                        self.applied_corrections[cell_key] = target_to_check
                        results.append({
                            "id": cell_key,
                            "row": r_idx,
                            "col_idx": tgt_col_idx,
                            "col_letter": tgt_letter,
                            "col_name": tgt_col_name,
                            "ref_col_name": m["ref_col_name"],
                            "cell": cell_coord,
                            "original_value": val_str,
                            "current_value": target_to_check,
                            "status": "OK",
                            "status_label": "Gültig",
                            "badge_class": "badge-ok",
                            "score": 1.0,
                            "suggestion": target_to_check,
                            "detail": f"Gültige Ressource (wird beim Speichern automatisch 4-stellig als '{target_to_check}' formatiert)",
                            "candidates": [],
                            "is_corrected": False
                        })
                    else:
                        results.append({
                            "id": cell_key,
                            "row": r_idx,
                            "col_idx": tgt_col_idx,
                            "col_letter": tgt_letter,
                            "col_name": tgt_col_name,
                            "ref_col_name": m["ref_col_name"],
                            "cell": cell_coord,
                            "original_value": val_str,
                            "current_value": val_str,
                            "status": "OK",
                            "status_label": "Gültig",
                            "badge_class": "badge-ok",
                            "score": 1.0,
                            "suggestion": val_str,
                            "detail": "Exakte Übereinstimmung mit Referenz",
                        })
                else:
                    stats["total_errors"] += 1
                    # Kandidaten suchen basierend auf target_to_check (damit Zahlendreher korrekt ermittelt werden)
                    candidates = index.find_candidates(target_to_check)
                    
                    if candidates:
                        best = candidates[0]
                        err_type = best["type"]
                        sugg_val = best["value"]
                        detail_desc = best["detail"]
                        conf_score = best["score"]

                        if err_type == "ZAHLENDREHER":
                            stats["zahlendreher"] += 1
                            lbl = "Zahlendreher"
                            bclass = "badge-zahlendreher"
                        elif err_type == "ZIFFERENTAUSCH":
                            stats["zifferntausch"] += 1
                            lbl = "Zifferntausch"
                            bclass = "badge-zifferntausch"
                        elif err_type == "TIPPFEHLER":
                            stats["tippfehler"] += 1
                            lbl = "Tippfehler (1 Ziffer)"
                            bclass = "badge-tippfehler"
                        elif err_type in ("ZIFFER_ZUVIEL", "ZIFFER_FEHLT"):
                            stats["ziffer_zuviel_zuwenig"] += 1
                            lbl = "Ziffer zuviel/fehlt"
                            bclass = "badge-tippfehler"
                        elif err_type == "NAECHSTE_ALTERNATIVE":
                            stats["nicht_existent"] += 1
                            lbl = "Falscher Wert (Alternative ermittelt)"
                            bclass = "badge-fehler"
                        else:
                            lbl = "Ähnlicher Wert"
                            bclass = "badge-aehnlich"

                        results.append({
                            "id": cell_key,
                            "row": r_idx,
                            "col_idx": tgt_col_idx,
                            "col_letter": tgt_letter,
                            "col_name": tgt_col_name,
                            "ref_col_name": m["ref_col_name"],
                            "cell": cell_coord,
                            "original_value": val_str,
                            "current_value": val_str,
                            "status": err_type,
                            "status_label": lbl,
                            "badge_class": bclass,
                            "score": conf_score,
                            "suggestion": sugg_val,
                            "detail": detail_desc,
                            "candidates": candidates,
                            "is_corrected": False
                        })
                    else:
                        stats["nicht_existent"] += 1
                        results.append({
                            "id": cell_key,
                            "row": r_idx,
                            "col_idx": tgt_col_idx,
                            "col_letter": tgt_letter,
                            "col_name": tgt_col_name,
                            "ref_col_name": m["ref_col_name"],
                            "cell": cell_coord,
                            "original_value": val_str,
                            "current_value": val_str,
                            "status": "NICHT_EXISTENT",
                            "status_label": "Nicht existent",
                            "badge_class": "badge-fehler",
                            "score": 0.0,
                            "suggestion": "",
                            "detail": "Wert kommt in der Referenzspalte nicht vor (keine Ähnlichkeit)",
                            "candidates": [],
                            "is_corrected": False
                        })

        self.results = results
        return {
            "stats": stats,
            "results": results,
            "target_sheet": target_sheet,
            "ref_sheet": ref_sheet
        }

    def apply_single_correction(self, cell_id: str, new_value: str, propagate_same_value: bool = True) -> Dict[str, Any]:
        """
        Wendet eine Korrektur auf eine Zelle an.
        Wenn propagate_same_value=True ist, werden alle anderen Zellen in derselben Spalte,
        die denselben ursprünglichen (falschen) Wert hatten, automatisch miterfasst.
        """
        new_val_clean = str(new_value).strip()

        # Finde Ziel-Zelle
        target_item = None
        for r in self.results:
            if r["id"] == cell_id:
                target_item = r
                break

        updated_items = []
        if target_item and propagate_same_value:
            orig_val = target_item.get("original_value")
            col_idx = target_item.get("col_idx")

            for r in self.results:
                # Dieselbe Spalte und derselbe Originalwert
                if r.get("col_idx") == col_idx and r.get("original_value") == orig_val:
                    r["current_value"] = new_val_clean
                    r["is_corrected"] = (new_val_clean != r.get("original_value"))
                    if r["is_corrected"]:
                        self.applied_corrections[r["id"]] = new_val_clean
                    else:
                        self.applied_corrections.pop(r["id"], None)
                    updated_items.append(r)
        elif target_item:
            target_item["current_value"] = new_val_clean
            target_item["is_corrected"] = (new_val_clean != target_item.get("original_value"))
            if target_item["is_corrected"]:
                self.applied_corrections[cell_id] = new_val_clean
            else:
                self.applied_corrections.pop(cell_id, None)
            updated_items.append(target_item)
        else:
            self.applied_corrections[cell_id] = new_val_clean

        return {
            "updated_count": len(updated_items),
            "updated_ids": [it["id"] for it in updated_items],
            "item": target_item or {}
        }

    def apply_auto_corrections(self, mode: str = "zahlendreher_only") -> Dict[str, Any]:
        """
        Wendet automatische Korrekturen an.
        mode='zahlendreher_only': Nur eindeutige Zahlendreher / Zifferntausch
        mode='all_high_confidence': Alle Vorschläge mit Score >= 0.85
        """
        corrected_count = 0
        for r in self.results:
            if r["status"] in ("OK", "OK_LEER"):
                continue
            if not r["suggestion"]:
                continue

            should_apply = False
            if mode == "zahlendreher_only":
                if r["status"] in ("ZAHLENDREHER", "ZIFFERENTAUSCH") and r["score"] >= 0.90:
                    should_apply = True
            elif mode == "all_high_confidence":
                if r["score"] >= 0.85:
                    should_apply = True
            elif mode == "all_suggestions":
                should_apply = True
            elif mode == "general_button":
                # Alle Korrekturvorschläge für Zahlendreher, Tippfehler etc. beim General-Button automatisch anwenden
                if r["status"] not in ("DATUM_WARNUNG", "NICHT_EXISTENT", "LEER") and r["suggestion"]:
                    should_apply = True

            if should_apply:
                r["current_value"] = r["suggestion"]
                r["is_corrected"] = True
                self.applied_corrections[r["id"]] = r["suggestion"]
                corrected_count += 1

        return {
            "corrected_count": corrected_count,
            "total_applied": len(self.applied_corrections)
        }

    def reset_corrections(self) -> Dict[str, Any]:
        """Setzt alle angewendeten Korrekturen zurück."""
        for r in self.results:
            r["current_value"] = r["original_value"]
            r["is_corrected"] = False
        self.applied_corrections = {}
        return {"status": "ok", "message": "Alle Korrekturen zurückgesetzt."}

    def export_corrected_excel(self, highlight_corrections: bool = False) -> bytes:
        """
        Exportiert die korrigierte Excel-Datei.
        Behält das Original-Layout und Design exakt bei.
        """
        if not self.target_workbook_bytes:
            raise ValueError("Keine Prüfdatei geladen.")

        wb = openpyxl.load_workbook(io.BytesIO(self.target_workbook_bytes))
        
        # Formatierung für korrigierte Zellen (falls explizit gewünscht)
        corr_fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
        corr_font = Font(color="0F5132", bold=True)

        # Ziel-Arbeitsblatt ermitteln
        if hasattr(self, "last_target_sheet") and self.last_target_sheet in wb.sheetnames:
            ws = wb[self.last_target_sheet]
        else:
            ws = wb.active

        for cell_key, new_val in self.applied_corrections.items():
            try:
                row_str, col_str = cell_key.split("_")
                r_idx = int(row_str)
                c_idx = int(col_str)
                
                cell = ws.cell(row=r_idx, column=c_idx)
                
                # Falls führende Nullen vorhanden sind (z. B. "0045" oder "0123"),
                # MUSS der Wert als Text mit number_format = '@' gespeichert werden,
                # da Excel führende Nullen bei reinen Zahlen sonst abschneidet!
                try:
                    s_val = str(new_val).strip() if new_val is not None else ""
                    if not s_val:
                        cell.value = None
                    elif s_val.isdigit() and s_val.startswith("0") and len(s_val) > 1:
                        cell.number_format = "@"
                        cell.value = s_val
                    elif s_val.isdigit():
                        cell.value = int(s_val)
                    elif re.match(r"^-?\d+\.\d+$", s_val):
                        cell.value = float(s_val)
                    else:
                        cell.value = s_val
                except Exception:
                    cell.value = new_val

                if highlight_corrections:
                    cell.fill = corr_fill
                    cell.font = corr_font
            except Exception as e:
                print(f"Fehler beim Aktualisieren von Zelle {cell_key}: {e}")

        # Formatierung für alle Spalten vorgeben:
        # 1. Spalte A (Datum): Format "DD.MM.YYYY" (TT.MM.JJJJ), damit kein 5-stelliger Zahlenwert angezeigt wird
        # 2. Spalte C (Ressourcen): Format "0000", damit führende Nullen bei 2- oder 3-stelligen Zahlen angezeigt werden
        # 3. Alle anderen Spalten: Format "General" (Standard)
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            is_col_a = (col_idx == 1 or col_letter == "A")
            is_col_c = (col_idx == 3 or col_letter == "C")

            if is_col_a:
                ws.column_dimensions[col_letter].number_format = "DD.MM.YYYY"
                for r in range(2, ws.max_row + 1):
                    c = ws.cell(row=r, column=col_idx)
                    # Datumsangaben im Original NICHT verändern oder in fremde Typen konvertieren!
                    # Nur für Zahlen-/Datumswerte das Format auf DD.MM.YYYY setzen,
                    # Text-Datumsangaben (z.B. '14.09.2026') 1:1 unberührt als Text lassen.
                    if isinstance(c.value, (int, float, datetime.date, datetime.datetime)):
                        c.number_format = "DD.MM.YYYY"
            elif is_col_c:
                ws.column_dimensions[col_letter].number_format = "0000"
                for r in range(2, ws.max_row + 1):
                    c = ws.cell(row=r, column=col_idx)
                    c.number_format = "0000"
                    if c.value is not None:
                        val_s = str(c.value).strip()
                        if val_s.isdigit():
                            c.value = int(val_s)
            else:
                ws.column_dimensions[col_letter].number_format = "General"
                for r in range(2, ws.max_row + 1):
                    c = ws.cell(row=r, column=col_idx)
                    c.number_format = "General"

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def export_report_excel(self) -> bytes:
        """Erzeugt einen detaillierten Excel-Prüfbericht aller Fehler und Korrekturen."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Prüfbericht"

        # Tabellenkopf
        headers = [
            "Zelle", "Zeile", "Spalte", "Originaler Wert", "Aktueller/Korrigierter Wert",
            "Status", "Konfidenz", "Korrekturvorschlag", "Fehlerdetails", "Wurde Korrigiert?"
        ]
        
        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Zeilen schreiben (nur Fehler & korrigierte Zellen)
        row_counter = 2
        for r in self.results:
            if r["status"] in ("OK", "OK_LEER") and not r["is_corrected"]:
                continue

            ws.cell(row=row_counter, column=1, value=r["cell"])
            ws.cell(row=row_counter, column=2, value=r["row"])
            ws.cell(row=row_counter, column=3, value=r["col_name"])
            ws.cell(row=row_counter, column=4, value=r["original_value"])
            ws.cell(row=row_counter, column=5, value=r["current_value"])
            ws.cell(row=row_counter, column=6, value=r["status_label"])
            ws.cell(row=row_counter, column=7, value=f"{int(r['score'] * 100)}%" if r['score'] > 0 else "-")
            ws.cell(row=row_counter, column=8, value=r["suggestion"])
            ws.cell(row=row_counter, column=9, value=r["detail"])
            ws.cell(row=row_counter, column=10, value="JA" if r["is_corrected"] else "NEIN")
            
            # Zeilen einfärben nach Fehlertyp
            if r["status"] == "ZAHLENDREHER":
                fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            elif r["status"] == "NICHT_EXISTENT":
                fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            else:
                fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

            for c in range(1, 11):
                ws.cell(row=row_counter, column=c).fill = fill

            row_counter += 1

        # Spaltenbreiten anpassen
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

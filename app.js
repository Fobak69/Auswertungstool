/**
 * Excel-Auswertung Single-Page Web-App
 * Lokale Verarbeitung mit SheetJS (xlsx.js)
 * Tabs: 'Detail-Auswertung' & 'Wochenübersicht (Personal)'
 * Inklusive:
 *  - Auftrags-Orte Verknüpfung
 *  - Tägliche Stundenerfassung (Mo-So Matrix) mit strikt EINER Zeile pro Mitarbeiter
 *  - Manuelle Zusätze (Krank, Urlaub, Überstunden, Bemerkung) stehen direkt in der Zelle des Tages
 */

(() => {
  'use strict';

  // ==========================================================================
  // 1. ANWENDUNGSSTATUS & KONFIGURATION
  // ==========================================================================
  const STORAGE_KEY_STAMMDATEN = 'excel_app_stammdaten_v2';
  const STORAGE_KEY_MANUAL_ENTRIES = 'excel_app_manual_entries_v2';

  const state = {
    // Stammdaten: Ressourcen
    stammdaten: [],
    stammdatenMap: new Map(), // Key: normalized ressourcennummer -> Object

    // Stammdaten: Aufträge & Orte
    auftraege: [],
    auftraegeMap: new Map(), // Key: normalized auftragsnummer -> Ort
    auftragCompositeMap: new Map(), // Key: composite "463040##12", "463040-12" -> fullOrt
    auftragVariantsMap: new Map(), // Key: "463040" -> Array<{ ort, zusatzort, fullOrt }>

    stammdatenFilename: '',
    stammdatenUpdatedAt: null,

    // Manuelle Sonderzeiten (Krank, Urlaub, Überstundenabbau)
    // Array<{ id: string, ressourcennummer: string, dateIso: string, type: string, stunden: number, note: string, createdAt: string }>
    manualEntries: [],

    // Bewegungsdaten (Tagesberichte)
    bewegungsdaten: [],
    bewegungsdatenFilename: '',

    // Filter für Tab 1 (Detail-Auswertung)
    filters: {
      dateFrom: '',
      dateTo: '',
      category: '',
      search: ''
    },

    // Sortierung für Tab 1
    sort: {
      field: 'date',
      asc: false
    },

    // Ausgewählte Kalenderwoche für Tab 2
    selectedKwKey: '',

    // Filterung & Suche für Tab 2 (Wochenübersicht)
    showAllEmployeesInWeekly: true,
    weeklySearch: '',

    // Aktiver Tab ('tabDetail' | 'tabWoche')
    activeTab: 'tabDetail'
  };

  // Demo-Datensätze:
  // Spalte A = Mitarbeiternummer, Spalte B = Name Mitarbeiter
  // Spalte C = Auftragsnummer, Spalte D = Baustellen Ort, Spalte E = zusatzort (z. B. 463040 Spielplatz allgemein)
  // Spalte F = Fahrzeug und Geräte-Nummer, Spalte G = Fahrzeug/Geräte Name
  const DEMO_STAMMDATEN = [
    // Spalte A & B: Mitarbeiter
    { ressourcennummer: 'M001', name: 'Max Mustermann', kategorie: 'Mitarbeiter' },
    { ressourcennummer: 'M002', name: 'Anna Schmidt', kategorie: 'Mitarbeiter' },
    { ressourcennummer: 'M003', name: 'Lukas Weber', kategorie: 'Mitarbeiter' },
    { ressourcennummer: 'M004', name: 'Sarah Meyer', kategorie: 'Mitarbeiter' },
    { ressourcennummer: 'M005', name: 'Kevin Fischer', kategorie: 'Mitarbeiter' },
    { ressourcennummer: 'M006', name: 'Tariq Al-Mansoor', kategorie: 'Mitarbeiter' },

    // Spalte F & G: Fahrzeuge & Geräte
    { ressourcennummer: 'F101', name: 'Mercedes Sprinter (B-EX 101)', kategorie: 'Fahrzeug' },
    { ressourcennummer: 'F102', name: 'MAN Kipper 3-Achser (B-EX 202)', kategorie: 'Fahrzeug' },
    { ressourcennummer: 'F103', name: 'VW Caddy Service (B-EX 303)', kategorie: 'Fahrzeug' },
    { ressourcennummer: 'G201', name: 'Mobilbagger Liebherr A914', kategorie: 'Maschine' },
    { ressourcennummer: 'G202', name: 'Minibagger Kubota KX057', kategorie: 'Maschine' },
    { ressourcennummer: 'G203', name: 'Rüttelplatte Wacker DPU 6555', kategorie: 'Maschine' }
  ];

  const DEMO_AUFTRAEGE = [
    // Spalte C = Auftragsnummer, Spalte D = Baustellen Ort, Spalte E = zusatzort
    { auftragsnummer: '463040', ort: 'Spielplatz allgemein', zusatzort: '12 - Schillerpark' },
    { auftragsnummer: '463040', ort: 'Spielplatz allgemein', zusatzort: '05 - Goetheplatz' },
    { auftragsnummer: '463040', ort: 'Spielplatz allgemein', zusatzort: '08 - Stadtpark Süd' },
    { auftragsnummer: 'AUF-2026-101', ort: 'Berlin-Mitte', zusatzort: 'Alexanderplatz' },
    { auftragsnummer: 'AUF-2026-102', ort: 'Potsdam', zusatzort: 'Gewerbepark Babelsberg' },
    { auftragsnummer: 'AUF-2026-103', ort: 'Berlin-Charlottenburg', zusatzort: 'Kurfürstendamm' },
    { auftragsnummer: 'AUF-2026-095', ort: 'Königs Wusterhausen', zusatzort: 'Trassenbau Süd' }
  ];

  const DEMO_MANUAL_ENTRIES = [
    {
      id: 'demo_man_1',
      ressourcennummer: 'M004', // Sarah Meyer
      dateIso: '2026-09-24', // Donnerstag KW 39
      type: 'Urlaub',
      stunden: 8.0,
      note: 'Erholungsurlaub',
      createdAt: '2026-09-21T08:00:00Z'
    },
    {
      id: 'demo_man_2',
      ressourcennummer: 'M005', // Kevin Fischer
      dateIso: '2026-09-25', // Freitag KW 39 (hat 4h gearbeitet + 2h Überstundenabbau)
      type: 'Überstundenabbau',
      stunden: 2.0,
      note: '2h Gleitzeit früher',
      createdAt: '2026-09-22T08:00:00Z'
    },
    {
      id: 'demo_man_3',
      ressourcennummer: 'M006', // Tariq Al-Mansoor
      dateIso: '2026-09-22', // Dienstag KW 39
      type: 'Krank',
      stunden: 8.0,
      note: 'AU attestiert',
      createdAt: '2026-09-22T07:30:00Z'
    }
  ];

  const DEMO_BEWEGUNGSDATEN_RAW = [
    // Montag 21.09.2026 (KW 39)
    ['2026-09-21', '463040-12', 'M001', 8.5, 'Spielplatzbau', 'Klettergerüst aufbauen Spielplatz 12 (Schillerpark)', 552.50],
    ['2026-09-21', 'AUF-2026-101', 'G201', 7.0, 'Maschineneinsatz', 'Baggerarbeiten Baugrube', 665.00],
    ['2026-09-21', '463040', 'M002', 8.0, 'Spielplatzbau', 'Fallschutzmatten verlegen Spielplatz 12', 480.00],
    ['2026-09-21', '463040-05', 'M003', 8.5, 'Spielplatzbau', 'Schaukelanlage montieren Spielplatz 05', 510.00],
    ['2026-09-21', 'AUF-2026-102', 'F101', 4.0, 'Transport', 'Werkzeug- und Materialanfuhr', 180.00],
    ['2026-09-21', 'AUF-2026-103', 'M004', 8.0, 'Pflasterbau', 'Unterbau verdichten und abziehen', 480.00],
    ['2026-09-21', 'AUF-2026-103', 'G203', 5.0, 'Maschineneinsatz', 'Flächenverdichtung Parkplatz', 225.00],

    // Dienstag 22.09.2026
    ['2026-09-22', 'AUF-2026-101', 'M001', 8.0, 'Erdaushub', 'Fundamentaushub Einzelfundamente', 520.00],
    ['2026-09-22', 'AUF-2026-101', 'G201', 8.0, 'Maschineneinsatz', 'Fundamente schachten', 760.00],
    ['2026-09-22', 'AUF-2026-101', 'M002', 8.0, 'Betonarbeiten', 'Sauberkeitsschicht einbringen', 480.00],
    ['2026-09-22', 'AUF-2026-102', 'M003', 8.0, 'Kabeltiefbau', 'Schutzrohre DN110 verlegen', 480.00],
    ['2026-09-22', 'AUF-2026-102', 'M005', 8.5, 'Kabeltiefbau', 'Rohrverlegung und Trassenband', 510.00],
    ['2026-09-22', 'AUF-2026-102', 'F102', 6.0, 'Transport', 'Sandlieferung 18t verfüllen', 450.00],
    ['2026-09-22', 'AUF-2026-103', 'M004', 8.5, 'Pflasterbau', 'Verbundsteinpflaster verlegen', 510.00],

    // Mittwoch 23.09.2026
    ['2026-09-23', 'AUF-2026-101', 'M001', 9.0, 'Bewehrung', 'Stahlbewehrung verlegen', 585.00],
    ['2026-09-23', 'AUF-2026-101', 'M002', 8.5, 'Schalung', 'Fundamentschalung aufbauen', 510.00],
    ['2026-09-23', 'AUF-2026-102', 'M003', 8.0, 'Kabeltiefbau', 'Kabelzug 4x240mm2', 480.00],
    ['2026-09-23', 'AUF-2026-102', 'M005', 8.0, 'Kabeltiefbau', 'Kabelzughilfe und Abdichtung', 480.00],
    ['2026-09-23', 'AUF-2026-103', 'M004', 8.0, 'Pflasterbau', 'Pflasterflächen einsanden', 480.00],
    ['2026-09-23', 'AUF-2026-103', 'M006', 8.5, 'Pflasterbau', 'Randsteine setzen in Beton', 510.00],
    ['2026-09-23', 'AUF-2026-103', 'G202', 6.5, 'Maschineneinsatz', 'Bodenbewegung Randbereiche', 455.00],

    // Donnerstag 24.09.2026
    ['2026-09-24', 'AUF-2026-101', 'M001', 8.0, 'Betonarbeiten', 'Betonage Bodenplatte C25/30', 520.00],
    ['2026-09-24', 'AUF-2026-101', 'M002', 8.5, 'Betonarbeiten', 'Betonverteilung und Rütteln', 552.50],
    ['2026-09-24', 'AUF-2026-102', 'M003', 8.5, 'Kabeltiefbau', 'Verfüllung und Verdichtung', 510.00],
    ['2026-09-24', 'AUF-2026-102', 'M005', 8.0, 'Kabeltiefbau', 'Verdichtungskontrolle Dynamisch', 480.00],
    ['2026-09-24', 'AUF-2026-102', 'G203', 6.0, 'Maschineneinsatz', 'Lagenweise Verdichtung Graben', 270.00],
    ['2026-09-24', 'AUF-2026-103', 'M006', 8.0, 'Pflasterbau', 'Anschneidearbeiten Ecken', 480.00],

    // Freitag 25.09.2026
    ['2026-09-25', 'AUF-2026-101', 'M001', 6.5, 'Nachbehandlung', 'Betonnachbehandlung Folie/Wasser', 422.50],
    ['2026-09-25', 'AUF-2026-101', 'M002', 6.5, 'Ausschalen', 'Randschalungen entfernen/reinigen', 422.50],
    ['2026-09-25', 'AUF-2026-102', 'M003', 6.0, 'Dokumentation', 'Trassenaufmaß und Fotos', 360.00],
    ['2026-09-25', 'AUF-2026-102', 'M005', 4.0, 'Baustellenräumung', 'Werkzeug und Absperrung rückbauen', 240.00],
    ['2026-09-25', 'AUF-2026-102', 'F101', 4.0, 'Transport', 'Rücktransport Werkzeugmagazin', 180.00],
    ['2026-09-25', 'AUF-2026-103', 'M004', 6.5, 'Übergabe', 'Endreinigung und Übergabe Bauherr', 390.00],
    ['2026-09-25', 'AUF-2026-103', 'M006', 6.5, 'Baustellenräumung', 'Abfuhr Reste und Besenreinigung', 390.00],

    // Vorwoche KW 38
    ['2026-09-16', 'AUF-2026-095', 'M001', 8.0, 'Baustelleneinrichtung', 'Bauzaun und Container aufstellen', 520.00],
    ['2026-09-16', 'AUF-2026-095', 'M002', 8.0, 'Baustelleneinrichtung', 'Strom- und Wasseranschluss', 480.00],
    ['2026-09-17', 'AUF-2026-095', 'M001', 8.5, 'Erdaushub', 'Mutterboden abtragen', 552.50],
    ['2026-09-17', 'AUF-2026-095', 'G201', 8.0, 'Maschineneinsatz', 'Baggerarbeiten Abschieben', 760.00],
    ['2026-09-18', 'AUF-2026-095', 'M003', 7.5, 'Vermessung', 'Schnurgerüst einmessen', 450.00]
  ];

  // ==========================================================================
  // 2. INITIALISIERUNG
  // ==========================================================================
  document.addEventListener('DOMContentLoaded', () => {
    initApp();
  });

  function initApp() {
    setupDomEvents();
    loadStammdatenFromStorage();
    loadManualEntriesFromStorage();
    renderApp();
  }

  // ==========================================================================
  // 3. STORAGE & STAMMDATEN MANAGEMENT
  // ==========================================================================
  function normalizeKey(str) {
    if (str === null || str === undefined) return '';
    let s = String(str).trim().toLowerCase();
    // Excel Float-Artefakt wie "10.0" bereinigen
    if (/^\d+\.0+$/.test(s)) {
      s = s.replace(/\.0+$/, '');
    }
    return s;
  }

  // Liefert für rein numerische IDs die Zahl ohne führende Nullen (z. B. "0376" -> "376", "0010" -> "10")
  function normalizeResourceKey(str) {
    const norm = normalizeKey(str);
    if (!norm) return '';
    if (/^\d+$/.test(norm)) {
      const stripped = norm.replace(/^0+/, '');
      return stripped || '0';
    }
    return norm;
  }

  function loadStammdatenFromStorage() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_STAMMDATEN) || localStorage.getItem('excel_app_stammdaten_v1');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed.records) && parsed.records.length > 0) {
          state.stammdaten = parsed.records;
          state.auftraege = Array.isArray(parsed.auftraege) ? parsed.auftraege : [];
          state.stammdatenFilename = parsed.filename || 'Stammdaten.xlsx';
          state.stammdatenUpdatedAt = parsed.savedAt ? new Date(parsed.savedAt) : new Date();

          buildLookupMaps();
          updateStammdatenStatusUI();
          return true;
        }
      }
    } catch (e) {
      console.error('Fehler beim Laden der Stammdaten aus LocalStorage:', e);
    }
    updateStammdatenStatusUI();
    return false;
  }

  function saveStammdatenToStorage(records, auftraege = [], filename = 'Stammdaten.xlsx') {
    state.stammdaten = records;
    state.auftraege = auftraege;
    state.stammdatenFilename = filename;
    state.stammdatenUpdatedAt = new Date();

    const isRealFile = !filename.toLowerCase().includes('demo') && !filename.toLowerCase().includes('muster');
    if (isRealFile) {
      // Demo-Abwesenheiten (z. B. M004, M005, M006) automatisch aus dem Speicher bereinigen
      const prevCount = state.manualEntries.length;
      state.manualEntries = state.manualEntries.filter(e => {
        if (e.id && e.id.startsWith('demo_')) return false;
        if (e.ressourcennummer && e.ressourcennummer.startsWith('M00')) return false;
        return true;
      });
      if (state.manualEntries.length !== prevCount) {
        saveManualEntriesToStorage();
      }
      // Falls noch Demo-Tagesberichte im Speicher waren: leeren
      if (state.bewegungsdatenFilename && state.bewegungsdatenFilename.toLowerCase().includes('demo')) {
        state.bewegungsdaten = [];
        state.bewegungsdatenFilename = '';
        updateBewegungsdatenStatusUI();
      }
    }

    buildLookupMaps();

    try {
      const payload = {
        savedAt: state.stammdatenUpdatedAt.toISOString(),
        filename: state.stammdatenFilename,
        count: records.length,
        records: records,
        auftraegeCount: auftraege.length,
        auftraege: auftraege
      };
      localStorage.setItem(STORAGE_KEY_STAMMDATEN, JSON.stringify(payload));
      showToast(`Stammdaten gespeichert: ${records.length} Ressourcen & ${auftraege.length} Aufträge/Orte`, 'success');
    } catch (e) {
      console.error('Fehler beim Speichern in LocalStorage:', e);
      showToast('Konnte Stammdaten nicht im LocalStorage sichern.', 'error');
    }

    if (state.bewegungsdaten.length > 0) {
      relinkBewegungsdaten();
    }

    updateStammdatenStatusUI();
    renderApp();
  }

  function clearStammdatenStorage() {
    try {
      localStorage.removeItem(STORAGE_KEY_STAMMDATEN);
      localStorage.removeItem('excel_app_stammdaten_v1');
      localStorage.removeItem(STORAGE_KEY_MANUAL_ENTRIES);

      state.stammdaten = [];
      state.stammdatenMap.clear();
      state.auftraege = [];
      state.auftraegeMap.clear();
      if (state.auftragCompositeMap) state.auftragCompositeMap.clear();
      if (state.auftragVariantsMap) state.auftragVariantsMap.clear();
      state.manualEntries = [];
      state.stammdatenFilename = '';
      state.stammdatenUpdatedAt = null;

      if (state.bewegungsdaten.length > 0) {
        relinkBewegungsdaten();
      }

      updateStammdatenStatusUI();
      renderApp();
      showToast('Gesamter Speicher (Stammdaten, Aufträge & Abwesenheiten) geleert.', 'info');
    } catch (e) {
      console.error('Fehler beim Löschen des LocalStorage:', e);
    }
  }

  function formatFullOrt(ort, zusatzort) {
    const o = String(ort || '').trim();
    const z = String(zusatzort || '').trim();
    if (o && z) {
      if (o.toLowerCase().includes(z.toLowerCase())) return o;
      return `${o} (${z})`;
    }
    return o || z || '';
  }

  function buildLookupMaps() {
    state.stammdatenMap.clear();
    for (const item of state.stammdaten) {
      if (!item) continue;
      if (item.ressourcennummer) {
        const rawKey = normalizeKey(item.ressourcennummer);
        const resKey = normalizeResourceKey(item.ressourcennummer);
        state.stammdatenMap.set(rawKey, item);
        state.stammdatenMap.set(resKey, item);
        // Falls rein numerisch: auch 4-stellig mit führenden Nullen mappen (z. B. "0376")
        if (/^\d+$/.test(resKey) && resKey.length < 4) {
          state.stammdatenMap.set(resKey.padStart(4, '0'), item);
        }
      }
      if (item.name) {
        state.stammdatenMap.set(normalizeKey(item.name), item);
      }
    }

    state.auftraegeMap.clear();
    state.auftragCompositeMap = new Map();
    state.auftragVariantsMap = new Map();

    for (const a of state.auftraege) {
      if (!a) continue;
      const aNr = String(a.auftragsnummer || '').trim();
      if (!aNr) continue;

      const normNr = normalizeKey(aNr);
      const pureOrt = String(a.ort || '').trim();
      const zusatz = String(a.zusatzort || '').trim();
      const full = formatFullOrt(pureOrt, zusatz);

      // Variantenliste pro Auftragsnummer (z. B. 463040 hat mehrere Spielplätze)
      if (!state.auftragVariantsMap.has(normNr)) {
        state.auftragVariantsMap.set(normNr, []);
      }
      state.auftragVariantsMap.get(normNr).push({
        ort: pureOrt,
        zusatzort: zusatz,
        fullOrt: full
      });

      // Composite Keys für direkte Verknüpfungen (z. B. "463040##507", "463040-507", "463040 507")
      if (zusatz) {
        const normZ = normalizeKey(zusatz);
        const strippedZ = normZ.replace(/^0+/, '');
        state.auftragCompositeMap.set(`${normNr}##${normZ}`, full);
        state.auftragCompositeMap.set(`${normNr}-${normZ}`, full);
        state.auftragCompositeMap.set(`${normNr}/${normZ}`, full);
        state.auftragCompositeMap.set(`${normNr} ${normZ}`, full);
        state.auftragCompositeMap.set(`${normNr}_${normZ}`, full);
        state.auftragCompositeMap.set(`${normNr}.${normZ}`, full);

        if (strippedZ && strippedZ !== normZ) {
          state.auftragCompositeMap.set(`${normNr}##${strippedZ}`, full);
          state.auftragCompositeMap.set(`${normNr}-${strippedZ}`, full);
          state.auftragCompositeMap.set(`${normNr} ${strippedZ}`, full);
        }

        // Falls im Zusatzort eine reine Zahl steht (z. B. "12" aus "12 - Schillerpark")
        const numMatch = zusatz.match(/^(\d+)/);
        if (numMatch) {
          const num = numMatch[1];
          const strippedNum = num.replace(/^0+/, '');
          state.auftragCompositeMap.set(`${normNr}##${num}`, full);
          state.auftragCompositeMap.set(`${normNr}-${num}`, full);
          state.auftragCompositeMap.set(`${normNr}/${num}`, full);
          state.auftragCompositeMap.set(`${normNr} ${num}`, full);
          state.auftragCompositeMap.set(`${normNr}_${num}`, full);
          state.auftragCompositeMap.set(`${normNr}.${num}`, full);

          if (strippedNum && strippedNum !== num) {
            state.auftragCompositeMap.set(`${normNr}##${strippedNum}`, full);
            state.auftragCompositeMap.set(`${normNr}-${strippedNum}`, full);
          }
        }
      }

      // Basis-Zuordnung (Fallback): Falls kein Zusatzort angegeben wird
      if (!state.auftraegeMap.has(normNr) || !zusatz) {
        state.auftraegeMap.set(normNr, pureOrt || full);
      }
    }

    // Spezieller Sammelauftrag 463040: Standardort ist "Spielplatz allgemein"
    if (state.auftragVariantsMap.has('463040')) {
      state.auftraegeMap.set('463040', 'Spielplatz allgemein');
    }
  }

  // ==========================================================================
  // 3b. MANUELLE ABWESENHEITEN (Krank, Urlaub, Überstundenabbau)
  // ==========================================================================
  function loadManualEntriesFromStorage() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_MANUAL_ENTRIES);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed)) {
          const isRealStamm = state.stammdaten.length > 0 && !state.stammdatenFilename.toLowerCase().includes('demo') && !state.stammdatenFilename.toLowerCase().includes('muster');
          if (isRealStamm) {
            state.manualEntries = parsed.filter(e => {
              if (e.id && e.id.startsWith('demo_')) return false;
              if (e.ressourcennummer && e.ressourcennummer.startsWith('M00')) return false;
              return true;
            });
          } else {
            state.manualEntries = parsed;
          }
        }
      }
    } catch (e) {
      console.error('Fehler beim Laden der manuellen Einträge:', e);
    }
  }

  function saveManualEntriesToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY_MANUAL_ENTRIES, JSON.stringify(state.manualEntries));
    } catch (e) {
      console.error('Fehler beim Speichern der manuellen Einträge:', e);
    }
    updateStammdatenStatusUI();
    renderApp();
  }

  function addOrUpdateManualEntry(entryData) {
    const { id, ressourcennummer, dateIso, type, stunden, note } = entryData;
    const rKey = normalizeKey(ressourcennummer);

    // Bestehenden Eintrag suchen
    let existingIndex = -1;
    if (id) {
      existingIndex = state.manualEntries.findIndex(e => e.id === id);
    } else {
      existingIndex = state.manualEntries.findIndex(e => 
        normalizeKey(e.ressourcennummer) === rKey && e.dateIso === dateIso
      );
    }

    const payload = {
      id: id || ('man_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6)),
      ressourcennummer: ressourcennummer,
      dateIso: dateIso,
      type: type || 'Krank',
      stunden: Number(stunden) || 0,
      note: note || '',
      createdAt: new Date().toISOString()
    };

    if (existingIndex !== -1) {
      state.manualEntries[existingIndex] = payload;
      showToast(`Eintrag (${payload.type} ${formatNumber(payload.stunden, 1)} Std.) aktualisiert.`, 'success');
    } else {
      state.manualEntries.push(payload);
      showToast(`Eintrag (${payload.type} ${formatNumber(payload.stunden, 1)} Std.) hinzugefügt.`, 'success');
    }

    saveManualEntriesToStorage();
  }

  function deleteManualEntryById(id) {
    const idx = state.manualEntries.findIndex(e => e.id === id);
    if (idx !== -1) {
      const removed = state.manualEntries.splice(idx, 1)[0];
      saveManualEntriesToStorage();
      showToast(`Eintrag (${removed.type}) gelöscht.`, 'info');
      return true;
    }
    return false;
  }

  function findManualEntry(ressourcennummer, dateIso) {
    if (!ressourcennummer || !dateIso) return null;
    const rKey = normalizeKey(ressourcennummer);
    return state.manualEntries.find(e => 
      normalizeKey(e.ressourcennummer) === rKey && e.dateIso === dateIso
    ) || null;
  }

  function updateStammdatenStatusUI() {
    const resCount = state.stammdaten.length;
    const aufCount = state.auftraege.length;
    const manCount = state.manualEntries.length;

    const card = document.getElementById('cardStammdaten');
    const badge = document.getElementById('badgeStammdatenStatus');
    const indicator = document.getElementById('indicatorStamm');
    const txtStatus = document.getElementById('txtStammdatenStatus');

    const modalStatus = document.getElementById('modalStammStatusText');
    const modalAuftragStatus = document.getElementById('modalAuftragStatusText');
    const modalManualCount = document.getElementById('modalManualCountText');
    const modalUpdated = document.getElementById('modalStammUpdatedText');
    const modalFilename = document.getElementById('modalStammFilenameText');
    const modalBody = document.getElementById('modalStammTableBody');
    const modalAuftragBody = document.getElementById('modalAuftragTableBody');
    const modalManualListBody = document.getElementById('modalManualListBody');

    if (resCount > 0 || aufCount > 0) {
      card.classList.add('loaded');
      badge.textContent = `${resCount} Ressourcen aktiv`;
      indicator.classList.add('active');
      const timeStr = state.stammdatenUpdatedAt ? formatGermanDateTime(state.stammdatenUpdatedAt) : '';
      txtStatus.textContent = `${resCount} Ressourcen & ${aufCount} Orte im Speicher (${state.stammdatenFilename || 'Referenz'})`;

      if (modalStatus) modalStatus.textContent = `${resCount} Ressourcen hinterlegt`;
      if (modalAuftragStatus) modalAuftragStatus.textContent = `${aufCount} Aufträge mit Ort hinterlegt`;
      if (modalManualCount) modalManualCount.textContent = `${manCount} Sonderzeiten erfasst`;
      if (modalUpdated) modalUpdated.textContent = timeStr || 'Unbekannt';
      if (modalFilename) modalFilename.textContent = state.stammdatenFilename || '--';

      if (modalBody) {
        modalBody.innerHTML = state.stammdaten.map(item => `
          <tr>
            <td><strong>${escapeHtml(item.ressourcennummer)}</strong></td>
            <td>${escapeHtml(item.name)}</td>
            <td><span class="badge ${getCategoryBadgeClass(item.kategorie)}">${escapeHtml(item.kategorie)}</span></td>
          </tr>
        `).join('');
      }

      if (modalAuftragBody) {
        if (state.auftraege.length > 0) {
          modalAuftragBody.innerHTML = state.auftraege.map(a => `
            <tr>
              <td><span class="order-code">${escapeHtml(a.auftragsnummer)}</span></td>
              <td><span class="ort-tag">${escapeHtml(a.ort || '–')}</span></td>
              <td>${a.zusatzort ? `<span class="ort-zusatz-badge">${escapeHtml(a.zusatzort)}</span>` : '<span class="text-muted">–</span>'}</td>
            </tr>
          `).join('');
        } else {
          modalAuftragBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted">Keine gesonderten Aufträge/Orte vorhanden.</td></tr>`;
        }
      }
    } else {
      card.classList.remove('loaded');
      badge.textContent = 'Warten auf Datei';
      indicator.classList.remove('active');
      txtStatus.textContent = 'Keine Stammdaten im Browser gespeichert.';

      if (modalStatus) modalStatus.textContent = '0 Ressourcen';
      if (modalAuftragStatus) modalAuftragStatus.textContent = '0 Aufträge';
      if (modalManualCount) modalManualCount.textContent = '0 Einträge';
      if (modalUpdated) modalUpdated.textContent = 'Nie';
      if (modalFilename) modalFilename.textContent = '--';
      if (modalBody) {
        modalBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted">Keine Stammdaten vorhanden.</td></tr>`;
      }
      if (modalAuftragBody) {
        modalAuftragBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted">Keine Aufträge vorhanden.</td></tr>`;
      }
    }

    if (modalManualListBody) {
      if (state.manualEntries.length > 0) {
        const sortedMan = [...state.manualEntries].sort((a, b) => b.dateIso.localeCompare(a.dateIso));
        modalManualListBody.innerHTML = sortedMan.map(m => {
          const emp = lookupStammdaten(m.ressourcennummer);
          const empName = emp ? emp.name : m.ressourcennummer;
          return `
            <tr>
              <td><strong>${escapeHtml(formatGermanDate(parseAnyDate(m.dateIso)))}</strong></td>
              <td>${escapeHtml(empName)} (${escapeHtml(m.ressourcennummer)})</td>
              <td><span class="badge-absence ${getAbsenceBadgeClass(m.type)}">${escapeHtml(m.type)}</span></td>
              <td class="text-right font-bold">${formatNumber(m.stunden, 1)} Std.</td>
              <td style="font-size: 0.8rem; color: #475569;">${escapeHtml(m.note || '–')}</td>
              <td class="text-center">
                <button type="button" class="btn btn-chip" style="color: #dc2626;" onclick="window.deleteManualEntryDirect('${m.id}')">Löschen</button>
              </td>
            </tr>
          `;
        }).join('');
      } else {
        modalManualListBody.innerHTML = `<tr><td colspan="6" class="text-center text-muted">Keine Abwesenheiten eingetragen.</td></tr>`;
      }
    }

    updateCategoryDropdown();
    populateManualModalEmployees();
  }

  function updateCategoryDropdown() {
    const sel = document.getElementById('filterCategory');
    if (!sel) return;

    const currentVal = sel.value;
    const categories = new Set();

    for (const item of state.stammdaten) {
      if (item.kategorie && item.kategorie.trim()) {
        categories.add(item.kategorie.trim());
      }
    }
    for (const item of state.bewegungsdaten) {
      if (item.kategorie && item.kategorie !== 'Nicht zugeordnet') {
        categories.add(item.kategorie);
      }
    }

    if (categories.size === 0) {
      categories.add('Mitarbeiter');
      categories.add('Fahrzeug');
      categories.add('Maschine');
    }

    const sortedCats = Array.from(categories).sort();
    let html = `<option value="">Alle Kategorien</option>`;
    for (const cat of sortedCats) {
      const selected = cat === currentVal ? 'selected' : '';
      html += `<option value="${escapeHtml(cat)}" ${selected}>${escapeHtml(cat)}</option>`;
    }
    sel.innerHTML = html;
  }

  // ==========================================================================
  // 4. DATEI-PARSING (SheetJS)
  // ==========================================================================
  function parseStammdatenFile(file) {
    if (!file) return;
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const parsedRecords = [];
        const parsedAuftraege = [];
        const seenAuftraege = new Set();
        const seenRessourcen = new Set();

        for (const sheetName of workbook.SheetNames) {
          const worksheet = workbook.Sheets[sheetName];
          const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          if (!rawRows || rawRows.length === 0) continue;

          // Header-Prüfung
          const firstRow = rawRows[0].map(c => String(c).toLowerCase().trim());
          const hasHeader = firstRow.some(c => 
            c.includes('mitarbeiter') || c.includes('pers') || c.includes('name') || 
            c.includes('auftrag') || c.includes('ort') || c.includes('baustelle') || 
            c.includes('zusatz') || c.includes('fahrzeug') || c.includes('gerät') || 
            c.includes('maschine') || c.includes('ressource')
          );
          const startIdx = hasHeader ? 1 : 0;

          // Standard-Spaltenzuordnung gemäß Benutzeranforderung:
          // Spalte A (0) = Mitarbeiternummer, Spalte B (1) = Name Mitarbeiter
          // Spalte C (2) = Auftragsnummer, Spalte D (3) = Baustellen Ort, Spalte E (4) = zusatzort
          // Spalte F (5) = Fahrzeug und Geräte-Nummer, Spalte G (6) = Fahrzeug/Geräte Name
          let colMA_Nr = 0;
          let colMA_Name = 1;
          let colAuf_Nr = 2;
          let colAuf_Ort = 3;
          let colAuf_Zusatz = 4;
          let colFG_Nr = 5;
          let colFG_Name = 6;

          if (hasHeader) {
            firstRow.forEach((h, idx) => {
              if (
                h.includes('personalnummer') || h.includes('personal-nr') || h.includes('pers.-nr') ||
                h.includes('personalnr') || h.includes('mitarbeiternummer') || h.includes('mitarbeiter-nr') ||
                (h.includes('mitarbeiter') && (h.includes('nr') || h.includes('id') || h.includes('nummer'))) ||
                (h.includes('personal') && (h.includes('nr') || h.includes('id') || h.includes('nummer')))
              ) {
                colMA_Nr = idx;
              } else if (
                h === 'name' || h.includes('name mitarbeiter') || h.includes('mitarbeitername') ||
                (h.includes('name') && !h.includes('fahrzeug') && !h.includes('gerät') && !h.includes('maschine')) ||
                (h.includes('mitarbeiter') && !h.includes('nr') && !h.includes('id') && !h.includes('nummer'))
              ) {
                colMA_Name = idx;
              } else if (h.includes('auftragsnummer') || h.includes('auftrag') || h.includes('projekt') || h === 'auftragnummer') {
                colAuf_Nr = idx;
              } else if (h.includes('baustellen ort') || h.includes('baustelle') || (h.includes('ort') && !h.includes('zusatz'))) {
                colAuf_Ort = idx;
              } else if (h.includes('spielplatz') || h.includes('zusatzort') || h.includes('zusatz') || h.includes('zusatznummer')) {
                colAuf_Zusatz = idx;
              } else if (
                (h.includes('fahrzeug') && (h.includes('nr') || h.includes('nummer'))) ||
                (h.includes('gerät') && (h.includes('nr') || h.includes('nummer'))) ||
                h.includes('kfz-nr') || h.includes('maschinen-nr') || h.includes('fahrzeugnummer')
              ) {
                colFG_Nr = idx;
              } else if (
                h === 'fahrzeug' || h === 'gerät' || h === 'maschine' ||
                h.includes('fahrzeug/geräte name') || h.includes('gerätename') ||
                (h.includes('fahrzeug') && h.includes('name')) ||
                (h.includes('gerät') && h.includes('name')) ||
                h.includes('maschinenname')
              ) {
                colFG_Name = idx;
              }
            });
          }

          // Dediziertes Auftragsblatt erkennen (falls Legacy-Datei mit 2 Blättern vorliegt)
          const sheetLower = sheetName.toLowerCase();
          const isDedicatedAuftragSheet = (sheetLower.includes('auftrag') || sheetLower.includes('baustelle')) && rawRows[0].length <= 3;

          if (isDedicatedAuftragSheet) {
            for (let r = startIdx; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row || row.length === 0) continue;
              const aNr = String(row[0] || '').trim();
              const aOrt = String(row[1] || '').trim();
              const aZusatz = String(row[2] || '').trim();
              if (aNr || aOrt || aZusatz) {
                const aKey = normalizeKey(`${aNr}##${aZusatz}##${aOrt}`);
                if (!seenAuftraege.has(aKey)) {
                  seenAuftraege.add(aKey);
                  parsedAuftraege.push({ auftragsnummer: aNr, ort: aOrt, zusatzort: aZusatz });
                }
              }
            }
          } else {
            // Standard-Verarbeitung: Spalten A..G parallel erfassen
            for (let r = startIdx; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row || row.length === 0) continue;

              // 1. Spalte A & B: Mitarbeiter
              const mNr = String(row[colMA_Nr] || '').trim();
              const mName = String(row[colMA_Name] || '').trim();
              if (mNr || mName) {
                const rKey = normalizeResourceKey(mNr) || normalizeKey(mName);
                if (!seenRessourcen.has(rKey)) {
                  seenRessourcen.add(rKey);
                  parsedRecords.push({
                    ressourcennummer: mNr,
                    name: mName || `Mitarbeiter ${mNr}`,
                    kategorie: 'Mitarbeiter'
                  });
                }
              }

              // 2. Spalte C, D & E: Aufträge, Baustellen-Orte & Zusatzorte
              const aNr = String(row[colAuf_Nr] || '').trim();
              const aOrt = String(row[colAuf_Ort] || '').trim();
              const aZusatz = String(row[colAuf_Zusatz] || '').trim();
              if (aNr || aOrt || aZusatz) {
                const aKey = normalizeKey(`${aNr}##${aZusatz}##${aOrt}`);
                if (!seenAuftraege.has(aKey)) {
                  seenAuftraege.add(aKey);
                  parsedAuftraege.push({
                    auftragsnummer: aNr,
                    ort: aOrt,
                    zusatzort: aZusatz
                  });
                }
              }

              // 3. Spalte F & G: Fahrzeuge & Geräte
              const fNr = String(row[colFG_Nr] || '').trim();
              const fName = String(row[colFG_Name] || '').trim();
              if (fNr || fName) {
                const fKey = normalizeKey(fNr) || normalizeKey(fName);
                if (!seenRessourcen.has(fKey)) {
                  seenRessourcen.add(fKey);

                  const lowerName = fName.toLowerCase();
                  let kat = 'Fahrzeug';
                  if (lowerName.includes('bagger') || lowerName.includes('rüttel') || 
                      lowerName.includes('walze') || lowerName.includes('stampfer') || 
                      lowerName.includes('gerät') || lowerName.includes('maschine') ||
                      lowerName.includes('radlader') || lowerName.includes('dumper') ||
                      lowerName.includes('häcksler') || lowerName.includes('schneider') ||
                      lowerName.includes('aggregat') || lowerName.includes('kompressor')) {
                    kat = 'Maschine';
                  } else if (lowerName.includes('lkw') || lowerName.includes('pkw') ||
                             lowerName.includes('sprinter') || lowerName.includes('transporter') ||
                             lowerName.includes('caddy') || lowerName.includes('bulli') ||
                             lowerName.includes('bus') || lowerName.includes('pritsch') ||
                             lowerName.includes('anhänger') || lowerName.includes('fahrzeug')) {
                    kat = 'Fahrzeug';
                  } else {
                    kat = 'Fahrzeug / Gerät';
                  }

                  parsedRecords.push({
                    ressourcennummer: fNr,
                    name: fName || `Gerät ${fNr}`,
                    kategorie: kat
                  });
                }
              }
            }
          }
        }

        if (parsedRecords.length === 0 && parsedAuftraege.length === 0) {
          showToast('Keine gültigen Datensätze in den Stammdaten gefunden.', 'error');
          return;
        }

        saveStammdatenToStorage(parsedRecords, parsedAuftraege, file.name);
      } catch (err) {
        console.error('Fehler beim Lesen der Stammdaten:', err);
        showToast('Fehler beim Einlesen der Excel-Datei: ' + err.message, 'error');
      }
    };

    reader.onerror = () => {
      showToast('Konnte Datei nicht lesen.', 'error');
    };

    reader.readAsArrayBuffer(file);
  }

  function parseBewegungsdatenFile(file) {
    if (!file) return;
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });

        if (!rawRows || rawRows.length === 0) {
          showToast('Die ausgewählte Bewegungsdaten-Datei ist leer.', 'error');
          return;
        }

        let startIndex = 0;
        const firstRow = rawRows[0].map(c => String(c).toLowerCase().trim());
        const hasHeader = firstRow.some(cell => 
          cell.includes('datum') || cell.includes('auftrag') || cell.includes('ressource') || cell.includes('stunde')
        );
        if (hasHeader) {
          startIndex = 1;
        }

        const rawItems = [];
        for (let i = startIndex; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0 || !row.some(c => String(c).trim() !== '')) continue;
          rawItems.push(row);
        }

        state.bewegungsdatenFilename = file.name;
        processBewegungsdatenRows(rawItems);
        showToast(`${state.bewegungsdaten.length} Tagesberichte eingelesen.`, 'success');
      } catch (err) {
        console.error('Fehler beim Lesen der Bewegungsdaten:', err);
        showToast('Fehler beim Einlesen der Excel-Datei: ' + err.message, 'error');
      }
    };

    reader.onerror = () => {
      showToast('Konnte Datei nicht lesen.', 'error');
    };

    reader.readAsArrayBuffer(file);
  }

  function processBewegungsdatenRows(rawRows) {
    const list = [];

    for (const row of rawRows) {
      const rawDate = row[0];
      const parsedDate = parseAnyDate(rawDate);
      const auftragsnummer = String(row[1] || '').trim();
      const ressourcennummer = String(row[2] || '').trim();
      const stundenRaw = row[3];
      const stunden = parseGermanNumber(stundenRaw);
      const leistung = String(row[4] || '').trim();
      const beschreibung = String(row[5] || '').trim();
      const preisRaw = row[6];
      const preis = parseGermanNumber(preisRaw);

      const stamm = lookupStammdaten(ressourcennummer);
      let ort = lookupAuftragOrt(auftragsnummer, beschreibung, leistung);
      if (!ort && row.length > 7 && row[7]) {
        ort = String(row[7]).trim();
      }

      const kwInfo = parsedDate ? getISOWeekDetails(parsedDate) : {
        week: 0,
        year: 0,
        key: 'unbekannt',
        label: 'Ohne Datum',
        rangeText: '--'
      };

      list.push({
        id: Math.random().toString(36).substring(2, 9),
        rawDate: rawDate,
        dateObj: parsedDate,
        dateIso: parsedDate ? formatDateIso(parsedDate) : '',
        dateDisplay: parsedDate ? formatGermanDate(parsedDate) : String(rawDate || '--'),
        weekday: parsedDate ? getWeekdayShort(parsedDate) : '',
        kwInfo: kwInfo,

        auftragsnummer: auftragsnummer || '–',
        ort: ort || '',

        ressourcennummer: ressourcennummer,

        name: stamm ? stamm.name : (ressourcennummer ? `Ressource (${ressourcennummer})` : 'Unbekannt'),
        kategorie: stamm ? stamm.kategorie : 'Nicht zugeordnet',
        isKnown: !!stamm,
        isMitarbeiter: stamm ? (String(stamm.kategorie).trim().toLowerCase() === 'mitarbeiter') : false,

        stunden: stunden,
        leistung: leistung || '–',
        beschreibung: beschreibung || '–',
        preis: preis,
        isManualAbsence: false
      });
    }

    state.bewegungsdaten = list;

    updateBewegungsdatenStatusUI();
    populateKalenderwochenSelect();
    renderApp();
  }

  function relinkBewegungsdaten() {
    for (const item of state.bewegungsdaten) {
      if (item.isManualAbsence) continue;
      const stamm = lookupStammdaten(item.ressourcennummer);
      if (stamm) {
        item.name = stamm.name;
        item.kategorie = stamm.kategorie;
        item.isKnown = true;
        item.isMitarbeiter = String(stamm.kategorie).trim().toLowerCase() === 'mitarbeiter';
      } else {
        item.name = item.ressourcennummer ? `Ressource (${item.ressourcennummer})` : 'Unbekannt';
        item.kategorie = 'Nicht zugeordnet';
        item.isKnown = false;
        item.isMitarbeiter = false;
      }
      item.ort = lookupAuftragOrt(item.auftragsnummer, item.beschreibung, item.leistung);
    }
    updateCategoryDropdown();
  }

  function lookupStammdaten(ressourcennummer) {
    if (!ressourcennummer) return null;
    const rawKey = normalizeKey(ressourcennummer);
    if (state.stammdatenMap.has(rawKey)) return state.stammdatenMap.get(rawKey);

    const resKey = normalizeResourceKey(ressourcennummer);
    if (state.stammdatenMap.has(resKey)) return state.stammdatenMap.get(resKey);

    if (/^\d+$/.test(resKey) && resKey.length < 4) {
      const padded = resKey.padStart(4, '0');
      if (state.stammdatenMap.has(padded)) return state.stammdatenMap.get(padded);
    }
    return null;
  }

  function lookupAuftragOrt(auftragsnummer, beschreibung = '', leistung = '') {
    if (!auftragsnummer) return '';
    const rawNr = String(auftragsnummer).trim();
    const cleanNr = normalizeKey(rawNr);
    const cleanLeistung = normalizeKey(leistung);
    const cleanDesc = normalizeKey(beschreibung);

    // 1. Wenn Leistung angegeben ist: Speziell für Sammelaufträge wie 463040 prüfen
    if (cleanLeistung) {
      const compKey = `${cleanNr}##${cleanLeistung}`;
      if (state.auftragCompositeMap && state.auftragCompositeMap.has(compKey)) {
        return state.auftragCompositeMap.get(compKey);
      }
      const strippedLeistung = cleanLeistung.replace(/^0+/, '');
      if (strippedLeistung && state.auftragCompositeMap.has(`${cleanNr}##${strippedLeistung}`)) {
        return state.auftragCompositeMap.get(`${cleanNr}##${strippedLeistung}`);
      }
    }

    // 2. Direkte Composite-Übereinstimmung (z. B. "463040##12", "463040-12")
    if (state.auftragCompositeMap && state.auftragCompositeMap.has(cleanNr)) {
      return state.auftragCompositeMap.get(cleanNr);
    }

    // Zerlegung prüfen: z. B. "463040-12" oder "463040/12" oder "463040 12"
    const splitMatch = rawNr.match(/^([A-Za-z0-9]+)[\s\-_/.:]+([A-Za-z0-9]+.*)$/);
    if (splitMatch) {
      const base = normalizeKey(splitMatch[1]);
      const ext = normalizeKey(splitMatch[2]);
      const compositeKey = `${base}##${ext}`;
      if (state.auftragCompositeMap && state.auftragCompositeMap.has(compositeKey)) {
        return state.auftragCompositeMap.get(compositeKey);
      }
      const numMatch = ext.match(/^(\d+)/);
      if (numMatch) {
        const numKey = `${base}##${numMatch[1]}`;
        if (state.auftragCompositeMap && state.auftragCompositeMap.has(numKey)) {
          return state.auftragCompositeMap.get(numKey);
        }
      }
    }

    // 3. Variantenabgleich bei Sammelaufträgen (z. B. 463040 Spielplatz allgemein mit Zusatznummern)
    if (state.auftragVariantsMap && state.auftragVariantsMap.has(cleanNr)) {
      const variants = state.auftragVariantsMap.get(cleanNr);
      if (variants && variants.length > 0) {
        const contextText = `${rawNr} ${cleanDesc} ${cleanLeistung}`.toLowerCase();
        
        for (const v of variants) {
          if (!v.zusatzort) continue;
          const zClean = normalizeKey(v.zusatzort);

          // Exakte Übereinstimmung mit Leistung
          if (cleanLeistung && (zClean === cleanLeistung || zClean.replace(/^0+/, '') === cleanLeistung.replace(/^0+/, ''))) {
            return v.fullOrt;
          }

          // Textsuche im Kontext (z. B. "Schillerpark")
          if (contextText.includes(zClean)) {
            return v.fullOrt;
          }

          // Zahlensuche als ganzes Wort (z. B. "12" oder "#12" oder "Spielplatz 12")
          const numMatch = zClean.match(/^(\d+)/);
          if (numMatch) {
            const num = numMatch[1];
            const numRegex = new RegExp(`(?:\\b|#|nr\\.?\\s*|spielplatz\\s*)${num}\\b`, 'i');
            if (numRegex.test(contextText)) {
              return v.fullOrt;
            }
          }
        }

        // Falls Auftrag 463040 (Spielplatz allgemein) ist:
        if (cleanNr === '463040') {
          return 'Spielplatz allgemein (463040)';
        }
      }
    }

    // 4. Basis-Auftragsort als Fallback
    if (state.auftraegeMap.has(cleanNr)) {
      return state.auftraegeMap.get(cleanNr);
    }

    if (cleanNr === '463040') {
      return 'Spielplatz allgemein (463040)';
    }

    return '';
  }

  function updateBewegungsdatenStatusUI() {
    const card = document.getElementById('cardBewegungsdaten');
    const badge = document.getElementById('badgeBewegungStatus');
    const indicator = document.getElementById('indicatorBewegung');
    const txtStatus = document.getElementById('txtBewegungStatus');
    const count = state.bewegungsdaten.length;

    if (count > 0) {
      card.classList.add('loaded');
      badge.textContent = `${count} Berichte aktiv`;
      indicator.classList.add('active');
      txtStatus.textContent = `${count} Zeilen geladen (${state.bewegungsdatenFilename || 'Tagesberichte'})`;
    } else {
      card.classList.remove('loaded');
      badge.textContent = 'Warten auf Datei';
      indicator.classList.remove('active');
      txtStatus.textContent = 'Noch keine Tagesberichte geladen.';
    }

    document.getElementById('navCountDetail').textContent = count + state.manualEntries.length;
  }

  // ==========================================================================
  // 5. KALENDERWOCHEN (ISO 8601) LOGIK & DROPDOWN
  // ==========================================================================
  function getISOWeekDetails(dateObj) {
    if (!dateObj || isNaN(dateObj.getTime())) {
      return { week: 0, year: 0, key: 'unbekannt', label: 'Ohne Datum', rangeText: '--' };
    }

    const target = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
    const dayNr = target.getUTCDay() || 7;
    target.setUTCDate(target.getUTCDate() + 4 - dayNr);

    const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((target - yearStart) / 86400000) + 1) / 7);
    const year = target.getUTCFullYear();

    const monday = new Date(target);
    monday.setUTCDate(monday.getUTCDate() - 3);
    const sunday = new Date(monday);
    sunday.setUTCDate(sunday.getUTCDate() + 6);

    const kwKey = `${year}-W${String(weekNo).padStart(2, '0')}`;
    const kwLabel = `KW ${String(weekNo).padStart(2, '0')} / ${year}`;
    const rangeText = `${formatGermanDate(monday)} bis ${formatGermanDate(sunday)}`;

    return {
      week: weekNo,
      year: year,
      key: kwKey,
      label: kwLabel,
      rangeText: rangeText,
      monday: monday,
      sunday: sunday
    };
  }

  function populateKalenderwochenSelect() {
    const sel = document.getElementById('selectKalenderwoche');
    if (!sel) return;

    const kwMap = new Map();

    for (const item of state.bewegungsdaten) {
      if (item.kwInfo && item.kwInfo.key && item.kwInfo.key !== 'unbekannt') {
        if (!kwMap.has(item.kwInfo.key)) {
          kwMap.set(item.kwInfo.key, {
            ...item.kwInfo,
            totalEntries: 0,
            mitarbeiterEntries: 0
          });
        }
        const kwData = kwMap.get(item.kwInfo.key);
        kwData.totalEntries++;
        if (item.isMitarbeiter) {
          kwData.mitarbeiterEntries++;
        }
      }
    }

    for (const m of state.manualEntries) {
      const d = parseAnyDate(m.dateIso);
      if (d) {
        const kw = getISOWeekDetails(d);
        if (kw.key && kw.key !== 'unbekannt') {
          if (!kwMap.has(kw.key)) {
            kwMap.set(kw.key, {
              ...kw,
              totalEntries: 0,
              mitarbeiterEntries: 0
            });
          }
          const kwData = kwMap.get(kw.key);
          kwData.totalEntries++;
          kwData.mitarbeiterEntries++;
        }
      }
    }

    if (kwMap.size === 0) {
      sel.innerHTML = `<option value="">-- Keine Daten vorhanden --</option>`;
      state.selectedKwKey = '';
      return;
    }

    const sortedKws = Array.from(kwMap.values()).sort((a, b) => b.key.localeCompare(a.key));

    let html = '';
    for (const kw of sortedKws) {
      const isSel = (!state.selectedKwKey && kw === sortedKws[0]) || (state.selectedKwKey === kw.key);
      if (isSel) state.selectedKwKey = kw.key;

      html += `<option value="${kw.key}" ${isSel ? 'selected' : ''}>
        ${kw.label} (${kw.rangeText})
      </option>`;
    }
    sel.innerHTML = html;
  }

  // ==========================================================================
  // 6. TAB 1: DETAIL-AUSWERTUNG RENDERN & FILTERN
  // ==========================================================================
  function getAllDetailItemsCombined() {
    const combined = [...state.bewegungsdaten];

    for (const m of state.manualEntries) {
      const stamm = lookupStammdaten(m.ressourcennummer);
      const pDate = parseAnyDate(m.dateIso);
      const kw = pDate ? getISOWeekDetails(pDate) : null;
      const isArbeitszeit = (m.type === 'Arbeitszeit' || String(m.type).toLowerCase().includes('arbeit'));

      combined.push({
        id: m.id,
        rawDate: m.dateIso,
        dateObj: pDate,
        dateIso: m.dateIso,
        dateDisplay: pDate ? formatGermanDate(pDate) : m.dateIso,
        weekday: pDate ? getWeekdayShort(pDate) : '',
        kwInfo: kw,

        auftragsnummer: isArbeitszeit ? 'Nachbuchung' : '–',
        ort: '–',
        ressourcennummer: m.ressourcennummer,

        name: stamm ? stamm.name : (m.ressourcennummer || 'Mitarbeiter'),
        kategorie: 'Mitarbeiter',
        isKnown: true,
        isMitarbeiter: true,

        stunden: m.stunden,
        leistung: isArbeitszeit ? 'Arbeitszeit (Nachbuchung)' : `Abwesenheit: ${m.type}`,
        beschreibung: m.note ? `${m.note} (Manuell erfasst)` : (isArbeitszeit ? 'Nachgetragene Arbeitszeit' : 'Manuelle Erfassung'),
        preis: 0,
        isManualAbsence: !isArbeitszeit,
        isManualWork: isArbeitszeit,
        absenceType: isArbeitszeit ? '' : m.type
      });
    }

    return combined;
  }

  function getFilteredDetailData() {
    const { dateFrom, dateTo, category, search } = state.filters;
    const searchLower = (search || '').trim().toLowerCase();
    const allItems = getAllDetailItemsCombined();

    return allItems.filter(item => {
      if (dateFrom && item.dateIso && item.dateIso < dateFrom) return false;
      if (dateTo && item.dateIso && item.dateIso > dateTo) return false;

      if (category && item.kategorie) {
        if (item.kategorie.toLowerCase() !== category.toLowerCase()) return false;
      }

      if (searchLower) {
        const matchName = item.name && item.name.toLowerCase().includes(searchLower);
        const matchOrder = item.auftragsnummer && item.auftragsnummer.toLowerCase().includes(searchLower);
        const matchOrt = item.ort && item.ort.toLowerCase().includes(searchLower);
        const matchRessource = item.ressourcennummer && item.ressourcennummer.toLowerCase().includes(searchLower);
        const matchLeistung = item.leistung && item.leistung.toLowerCase().includes(searchLower);
        const matchDesc = item.beschreibung && item.beschreibung.toLowerCase().includes(searchLower);

        if (!matchName && !matchOrder && !matchOrt && !matchRessource && !matchLeistung && !matchDesc) {
          return false;
        }
      }

      return true;
    });
  }

  function sortData(list) {
    const { field, asc } = state.sort;
    const modifier = asc ? 1 : -1;

    return [...list].sort((a, b) => {
      let valA, valB;
      switch (field) {
        case 'date':
          valA = a.dateIso || '';
          valB = b.dateIso || '';
          return valA.localeCompare(valB) * modifier;
        case 'order':
          valA = a.auftragsnummer || '';
          valB = b.auftragsnummer || '';
          return valA.localeCompare(valB) * modifier;
        case 'ort':
          valA = a.ort || '';
          valB = b.ort || '';
          return valA.localeCompare(valB, 'de') * modifier;
        case 'name':
          valA = a.name || '';
          valB = b.name || '';
          return valA.localeCompare(valB, 'de', { sensitivity: 'base' }) * modifier;
        case 'category':
          valA = a.kategorie || '';
          valB = b.kategorie || '';
          return valA.localeCompare(valB, 'de') * modifier;
        case 'hours':
          return (a.stunden - b.stunden) * modifier;
        case 'service':
          valA = a.leistung || '';
          valB = b.leistung || '';
          return valA.localeCompare(valB, 'de') * modifier;
        case 'price':
          return (a.preis - b.preis) * modifier;
        default:
          return 0;
      }
    });
  }

  function renderDetailTab() {
    const filtered = getFilteredDetailData();
    const sorted = sortData(filtered);
    const totalCombinedCount = state.bewegungsdaten.length + state.manualEntries.length;

    let sumStunden = 0;
    let sumPreis = 0;
    for (const item of sorted) {
      sumStunden += item.stunden || 0;
      sumPreis += item.preis || 0;
    }

    document.getElementById('kpiSumStunden').textContent = formatNumber(sumStunden, 2);
    document.getElementById('kpiFilteredCount').textContent = sorted.length;
    document.getElementById('kpiTotalRatio').textContent = `von ${totalCombinedCount} Gesamteinträgen`;
    document.getElementById('kpiSumPreis').textContent = formatCurrency(sumPreis);

    document.getElementById('tableResultCount').textContent = `(${sorted.length} Einträge)`;
    document.getElementById('navCountDetail').textContent = sorted.length;

    const foot = document.getElementById('detailTableFoot');
    if (sorted.length > 0) {
      foot.style.display = 'table-footer-group';
      document.getElementById('footSumStunden').textContent = `${formatNumber(sumStunden, 2)} Std./Menge`;
      document.getElementById('footSumPreis').textContent = formatCurrency(sumPreis);
    } else {
      foot.style.display = 'none';
    }

    const tbody = document.getElementById('detailTableBody');
    if (sorted.length === 0) {
      tbody.innerHTML = `
        <tr class="empty-row">
          <td colspan="9">
            <div class="empty-state">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="8" x2="12" y2="12"></line>
                <line x1="12" y1="16" x2="12.01" y2="16"></line>
              </svg>
              <p>${totalCombinedCount === 0 ? 'Keine Bewegungsdaten geladen. Bitte lade eine Excel-Datei hoch oder klicke oben auf "Demo-Daten laden".' : 'Keine Datensätze entsprechen den aktuellen Filtern.'}</p>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    tbody.innerHTML = sorted.map(item => {
      let leistungBadge = escapeHtml(item.leistung);
      if (item.isManualAbsence) {
        leistungBadge = `<span class="badge-absence ${getAbsenceBadgeClass(item.absenceType)}">${escapeHtml(item.absenceType)}</span> ${escapeHtml(item.leistung)}`;
      } else if (item.isManualWork) {
        leistungBadge = `<span class="badge-absence badge-arbeit">Nachbuchung</span> ${escapeHtml(item.leistung)}`;
      }

      return `
        <tr>
          <td>
            <span style="font-weight: 600;">${item.weekday ? item.weekday + ', ' : ''}${escapeHtml(item.dateDisplay)}</span>
          </td>
          <td>
            <span class="order-code">${escapeHtml(item.auftragsnummer)}</span>
          </td>
          <td>
            ${item.ort && item.ort !== '–' ? `<span class="ort-tag">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path>
                <circle cx="12" cy="10" r="3"></circle>
              </svg>
              ${escapeHtml(item.ort)}
            </span>` : '<span class="text-muted">–</span>'}
          </td>
          <td>
            <div class="res-name-box">
              <span class="res-name">${escapeHtml(item.name)}</span>
              <span class="res-id">${escapeHtml(item.ressourcennummer)}</span>
            </div>
          </td>
          <td>
            <span class="badge ${getCategoryBadgeClass(item.kategorie)}">
              ${escapeHtml(item.kategorie)}
            </span>
          </td>
          <td class="text-right font-bold" style="color: var(--primary);">
            ${formatNumber(item.stunden, 2)}
          </td>
          <td>${leistungBadge}</td>
          <td style="max-width: 250px; font-size: 0.82rem; color: #475569;">${escapeHtml(item.beschreibung)}</td>
          <td class="text-right">${formatCurrency(item.preis)}</td>
        </tr>
      `;
    }).join('');
  }

  // ==========================================================================
  // 7. TAB 2: WOCHENÜBERSICHT (PERSONAL) & TAGES-MATRIX (Mo - So)
  //    STRENG EINE ZEILE PRO MITARBEITER – ALLES IN DER TAGESZELLE
  // ==========================================================================
  function renderWochenTab() {
    const selKw = document.getElementById('selectKalenderwoche').value || state.selectedKwKey;
    state.selectedKwKey = selKw;

    // 1. Bewegungsdaten für diese Woche (Kategorie 'Mitarbeiter')
    const weekItems = state.bewegungsdaten.filter(item => {
      const matchWeek = item.kwInfo && item.kwInfo.key === selKw;
      return matchWeek && item.isMitarbeiter;
    });

    // 2. Montag dieser KW bestimmen
    let kwMonday = null;
    if (weekItems.length > 0 && weekItems[0].kwInfo && weekItems[0].kwInfo.monday) {
      kwMonday = new Date(weekItems[0].kwInfo.monday);
    } else if (selKw) {
      const match = selKw.match(/^(\d{4})-W(\d{2})$/);
      if (match) {
        const year = parseInt(match[1], 10);
        const week = parseInt(match[2], 10);
        const simple = new Date(Date.UTC(year, 0, 1 + (week - 1) * 7));
        const dayNr = simple.getUTCDay() || 7;
        simple.setUTCDate(simple.getUTCDate() + 4 - dayNr);
        simple.setUTCDate(simple.getUTCDate() - 3);
        kwMonday = simple;
      }
    }

    // Die 7 Tage der Woche aufbauen (Montag = 0, ..., Sonntag = 6)
    const weekDays = [];
    const dayNames = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
    if (kwMonday) {
      for (let i = 0; i < 7; i++) {
        const d = new Date(kwMonday);
        d.setUTCDate(d.getUTCDate() + i);
        const iso = formatDateIso(d);
        const dayStr = String(d.getUTCDate()).padStart(2, '0') + '.' + String(d.getUTCMonth() + 1).padStart(2, '0') + '.';
        weekDays.push({
          index: i,
          name: dayNames[i],
          dateObj: d,
          dateIso: iso,
          labelShort: dayStr
        });

        const th = document.getElementById(`thDay${i}`);
        if (th) {
          th.innerHTML = `${dayNames[i]}<br><span class="day-sub">${dayStr}</span>`;
        }
      }
    }

    // Einheitliche Mitarbeiter-Zuordnung: JEDER MITARBEITER ERHÄLT GENAU EINE ZEILE!
    const empMap = new Map();

    function getOrCreateEmp(ressourcennummer, fallbackName = '') {
      const rNum = String(ressourcennummer || '').trim();
      const stamm = lookupStammdaten(rNum);
      const canonicalNum = stamm ? stamm.ressourcennummer : (normalizeResourceKey(rNum) || rNum);
      const realName = stamm ? stamm.name : (fallbackName || rNum || 'Unbekannt');
      const key = normalizeResourceKey(canonicalNum) || normalizeKey(realName);

      if (!empMap.has(key)) {
        empMap.set(key, {
          ressourcennummer: canonicalNum,
          name: realName,
          kategorie: stamm ? stamm.kategorie : 'Mitarbeiter',
          workHours: 0,
          absenceHours: 0,
          totalHours: 0,
          workHoursByDay: [0, 0, 0, 0, 0, 0, 0],
          absenceByDay: [null, null, null, null, null, null, null],
          manualWorkByDay: [null, null, null, null, null, null, null],
          absenceList: [],
          detailsByDay: [[], [], [], [], [], [], []],
          datesSet: new Set(),
          orteSet: new Set(),
          absenceTypesSet: new Set()
        });
      }
      return empMap.get(key);
    }

    // A. Alle Mitarbeiter aus den Stammdaten vorinitialisieren
    for (const s of state.stammdaten) {
      if (s.kategorie && s.kategorie.trim().toLowerCase() === 'mitarbeiter') {
        getOrCreateEmp(s.ressourcennummer, s.name);
      }
    }

    // B. Reguläre Arbeitsstunden aus Bewegungsdaten zuordnen
    for (const item of weekItems) {
      const emp = getOrCreateEmp(item.ressourcennummer, item.name);
      const hours = item.stunden || 0;
      emp.workHours += hours;
      emp.totalHours += hours;

      if (item.dateIso) {
        emp.datesSet.add(item.dateIso);
      }
      if (item.ort) {
        emp.orteSet.add(item.ort);
      }

      if (item.dateObj) {
        const foundIdx = weekDays.findIndex(wd => wd.dateIso === item.dateIso);
        if (foundIdx !== -1) {
          emp.workHoursByDay[foundIdx] += hours;
          emp.detailsByDay[foundIdx].push(item);
        }
      }
    }

    // C. Manuelle Einträge (Arbeitszeit-Nachbuchung oder Abwesenheiten wie Krank/Urlaub/Überstunden)
    const isRealStamm = state.stammdaten.length > 0 && !state.stammdatenFilename.toLowerCase().includes('demo') && !state.stammdatenFilename.toLowerCase().includes('muster');
    for (const m of state.manualEntries) {
      if (isRealStamm && ((m.id && m.id.startsWith('demo_')) || (m.ressourcennummer && m.ressourcennummer.startsWith('M00')))) {
        continue;
      }
      const dayIdx = weekDays.findIndex(wd => wd.dateIso === m.dateIso);
      if (dayIdx !== -1) {
        // Mitarbeiter ermitteln – fließt strikt in dieselbe Zeile!
        const emp = getOrCreateEmp(m.ressourcennummer);
        const isArbeitszeit = (m.type === 'Arbeitszeit' || String(m.type).toLowerCase().includes('arbeit'));

        if (isArbeitszeit) {
          // Gilt als REGULÄRE ARBEITSZEIT (Nachbuchung fehlender Stunden)
          emp.workHours += m.stunden;
          emp.totalHours += m.stunden;
          emp.workHoursByDay[dayIdx] += m.stunden;
          emp.manualWorkByDay[dayIdx] = m;
          emp.datesSet.add(m.dateIso);

          emp.detailsByDay[dayIdx].push({
            auftragsnummer: 'Nachbuchung',
            ort: '',
            stunden: m.stunden,
            leistung: 'Arbeitszeit (Nachbuchung)' + (m.note ? `: ${m.note}` : '')
          });
        } else {
          // Gilt als SONDERZEIT / ABWESENHEIT (Krank, Urlaub, Überstundenabbau, etc.)
          emp.absenceByDay[dayIdx] = m;
          emp.absenceHours += m.stunden;
          emp.totalHours += m.stunden;
          emp.datesSet.add(m.dateIso);
          emp.absenceTypesSet.add(m.type);

          emp.absenceList.push({
            dayIdx: dayIdx,
            dayName: weekDays[dayIdx].name,
            dateDisplay: weekDays[dayIdx].labelShort,
            dateIso: m.dateIso,
            type: m.type,
            stunden: m.stunden,
            note: m.note || ''
          });
        }
      }
    }

    const allEmployees = Array.from(empMap.values());
    allEmployees.sort((a, b) => a.name.localeCompare(b.name, 'de', { sensitivity: 'base' }));

    for (const emp of allEmployees) {
      emp.absenceList.sort((a, b) => a.dayIdx - b.dayIdx);
    }

    // Gesamtsummen & Überstunden / Fehlzeiten berechnen (über alle Mitarbeiter der Woche)
    let kwTotalHours = 0;
    let kwTotalWorkHours = 0;
    let kwTotalOvertimeHours = 0;
    let kwTotalAbsenceHours = 0;
    const dayTotals = [0, 0, 0, 0, 0, 0, 0];

    for (const emp of allEmployees) {
      emp.overtimeHours = 0;
      emp.overtimeByDay = [0, 0, 0, 0, 0, 0, 0];
      emp.missingByDay = [0, 0, 0, 0, 0, 0, 0];
      emp.unclarifiedDays = [];

      kwTotalHours += emp.totalHours;
      kwTotalWorkHours += emp.workHours;
      kwTotalAbsenceHours += emp.absenceHours;

      for (let d = 0; d < 7; d++) {
        // Tagessumme Personal summiert STRIKT NUR die tatsächliche Arbeitszeit (keine Abwesenheiten/Sonderzeiten)
        const dWork = emp.workHoursByDay[d];
        dayTotals[d] += dWork;

        const targetH = getDailyTargetHours(d);
        const absH = emp.absenceByDay[d] ? (emp.absenceByDay[d].stunden || 0) : 0;
        const totalRec = dWork + absH;

        // 1. Überstunden: Alles über 8h (Mo-Do), über 7h (Fr) und Samstags/Sonntags (>0h)
        if (dWork > targetH) {
          const ot = dWork - targetH;
          emp.overtimeByDay[d] = ot;
          emp.overtimeHours += ot;
        }

        // 2. Unterstunden / Fehlzeiten: Mo-Fr wenn weniger gearbeitet wurde (totalRec < targetH)
        // Nur prüfen für Mitarbeiter, die in dieser Woche aktiv sind
        if (d < 5 && (emp.totalHours > 0 || emp.datesSet.size > 0)) {
          if (totalRec < targetH) {
            const missing = targetH - totalRec;
            emp.missingByDay[d] = missing;
            if (weekDays[d]) {
              emp.unclarifiedDays.push({
                dayIdx: d,
                dateIso: weekDays[d].dateIso,
                dayName: weekDays[d].name,
                dateDisplay: weekDays[d].labelShort,
                workH: dWork,
                absH: absH,
                targetH: targetH,
                missingH: missing
              });
            }
          }
        }
      }

      kwTotalOvertimeHours += emp.overtimeHours;
    }

    // Filterung der Anzeige: Alle Stammdaten-Mitarbeiter ODER nur aktive (mit Buchungen/Sonderzeiten)
    const showAll = state.showAllEmployeesInWeekly !== false; // Standardmäßig alle anzeigen
    const searchFilter = (state.weeklySearch || '').trim().toLowerCase();

    const activeWeekEmployees = allEmployees.filter(emp => {
      // 1. Aktivitäts-Filter
      if (!showAll) {
        if (!(emp.totalHours > 0 || emp.datesSet.size > 0)) {
          return false;
        }
      }
      // 2. Wochen-Suchfilter
      if (searchFilter) {
        const matchName = emp.name && emp.name.toLowerCase().includes(searchFilter);
        const matchNr = emp.ressourcennummer && emp.ressourcennummer.toLowerCase().includes(searchFilter);
        if (!matchName && !matchNr) return false;
      }
      return true;
    });

    const workingStaffCount = allEmployees.filter(emp => emp.workHours > 0).length;
    const avgWorkHours = workingStaffCount > 0 ? (kwTotalWorkHours / workingStaffCount) : 0;

    // Mini KPI Bar (Screen)
    document.getElementById('kpiWocheMitarbeiterCount').textContent = `${activeWeekEmployees.length} (${workingStaffCount} aktiv)`;
    document.getElementById('kpiWocheGesamtStunden').textContent = `${formatNumber(kwTotalWorkHours, 2)} Std.`;
    const kpiOt = document.getElementById('kpiWocheOvertimeStunden');
    if (kpiOt) kpiOt.textContent = `${formatNumber(kwTotalOvertimeHours, 2)} Std.`;
    document.getElementById('kpiWocheSchnittStunden').textContent = `${formatNumber(avgWorkHours, 1)} Std. / MA`;
    document.getElementById('kpiWocheSonderstunden').textContent = `${formatNumber(kwTotalAbsenceHours, 1)} Std.`;

    const totalMissingCount = activeWeekEmployees.reduce((sum, emp) => sum + emp.unclarifiedDays.length, 0);
    const kpiMissing = document.getElementById('kpiWocheMissingCount');
    if (kpiMissing) {
      if (totalMissingCount === 0) {
        kpiMissing.textContent = '0 Tage (alles geklärt)';
        kpiMissing.style.color = '#059669';
      } else {
        kpiMissing.textContent = `${totalMissingCount} Tag${totalMissingCount === 1 ? '' : 'e'} offen`;
        kpiMissing.style.color = '#dc2626';
      }
    }
    document.getElementById('navCountWoche').textContent = activeWeekEmployees.length;

    renderKlaerungsBox(activeWeekEmployees, selKw);

    updatePrintHeader(selKw, weekItems);

    // ==========================================
    // TABELLE 1: TÄGLICHE MATRIX RENDERN
    // Alles kompakt und lesbar in der Zelle des Tages!
    // ==========================================
    const matrixBody = document.getElementById('wochenTagesMatrixBody');
    if (activeWeekEmployees.length === 0) {
      matrixBody.innerHTML = `
        <tr class="empty-row">
          <td colspan="10">
            <div class="empty-state">
              <p>Keine Einträge für Mitarbeiter in der Kalenderwoche <strong>${escapeHtml(selKw || '--')}</strong> vorhanden.</p>
            </div>
          </td>
        </tr>
      `;
      for (let d = 0; d < 7; d++) {
        document.getElementById(`matrixSumDay${d}`).textContent = '0,00';
      }
      document.getElementById('matrixSumTotal').textContent = '0,00 Std.';
    } else {
      matrixBody.innerHTML = activeWeekEmployees.map(emp => {
        let cellsHtml = '';
        for (let d = 0; d < 7; d++) {
          const workH = emp.workHoursByDay[d];
          const abs = emp.absenceByDay[d];
          const manWork = emp.manualWorkByDay ? emp.manualWorkByDay[d] : null;
          const dayOvertime = emp.overtimeByDay ? emp.overtimeByDay[d] : 0;
          const missingH = emp.missingByDay ? emp.missingByDay[d] : 0;
          const targetH = getDailyTargetHours(d);
          const isWeekend = (d === 5 || d === 6);
          const tdClass = isWeekend ? 'weekend-td' : '';
          const targetDayIso = weekDays[d] ? weekDays[d].dateIso : '';

          let cellInner = '';
          let tooltipParts = [];

          if (workH > 0) {
            tooltipParts.push(emp.detailsByDay[d].map(t => 
              `${t.auftragsnummer}${t.ort ? ` (${t.ort})` : ''}: ${formatNumber(t.stunden, 1)}h [${t.leistung}]`
            ).join(' | '));
          }

          if (dayOvertime > 0) {
            tooltipParts.push(`Überstunden: ${formatNumber(dayOvertime, 1)} Std. (Mehrarbeit über ${targetH}h Soll)`);
          }

          if (abs) {
            tooltipParts.push(`Abwesenheit: ${abs.type} (${formatNumber(abs.stunden, 1)} Std.) ${abs.note ? '– ' + abs.note : ''}`);
          }

          if (manWork && !abs) {
            tooltipParts.push(`Manuelle Arbeitszeit: ${formatNumber(manWork.stunden, 1)} Std. ${manWork.note ? '– ' + manWork.note : ''}`);
          }

          if (missingH > 0) {
            tooltipParts.push(`⚠️ Offene Fehlzeit: ${formatNumber(missingH, 1)} Std. zur Sollzeit (${targetH}h) fehlen! Klicken zum Klären.`);
          }

          // Aufbau der Zelle: Alles kompakt in EINER Zelle des Tages!
          if (workH > 0 && abs) {
            // Fall 1: Arbeitszeit + Abwesenheitszusatz (z. B. 6h Arbeit + 2h Überstundenabbau)
            const noteText = [manWork && manWork.note ? `AZ: ${manWork.note}` : '', abs.note ? `${abs.type}: ${abs.note}` : ''].filter(Boolean).join(' | ');
            const noteHtml = noteText ? `<div class="cell-note-text" title="${escapeHtml(noteText)}">${escapeHtml(noteText)}</div>` : '';

            let overtimeBadge = '';
            if (dayOvertime > 0) {
              overtimeBadge = `<span class="badge-absence badge-overtime" title="Davon ${formatNumber(dayOvertime, 1)}h Überstunden">+${formatNumber(dayOvertime, 1).replace(',0', '')}h ÜSt</span>`;
            }

            let missingBadge = '';
            if (missingH > 0) {
              missingBadge = `<span class="badge-absence badge-missing" title="Noch ${formatNumber(missingH, 1)}h offen zur Sollzeit">⚠️ -${formatNumber(missingH, 1).replace(',0', '')}h</span>`;
            }

            cellInner = `
              <div class="matrix-cell-wrap ${missingH > 0 ? 'cell-missing-highlight' : ''}">
                <div class="cell-hours-row">
                  <span class="work-num">${formatNumber(workH, 2)}</span>
                  ${overtimeBadge}
                  <span class="badge-absence ${getAbsenceBadgeClass(abs.type)}">
                    ${formatAbsenceShort(abs.type, abs.stunden)}
                  </span>
                  ${missingBadge}
                </div>
                ${noteHtml}
              </div>
            `;
          } else if (abs) {
            // Fall 2: Nur Abwesenheit (z. B. ganzer Tag Krank oder Urlaub, Arbeitszeit = 0)
            const noteHtml = abs.note ? `<div class="cell-note-text" title="${escapeHtml(abs.note)}">${escapeHtml(abs.note)}</div>` : '';
            let missingBadge = '';
            if (missingH > 0) {
              missingBadge = `<span class="badge-absence badge-missing" title="Noch ${formatNumber(missingH, 1)}h offen zur Sollzeit">⚠️ -${formatNumber(missingH, 1).replace(',0', '')}h</span>`;
            }

            cellInner = `
              <div class="matrix-cell-wrap ${missingH > 0 ? 'cell-missing-highlight' : ''}">
                <div class="cell-hours-row">
                  <span class="badge-absence ${getAbsenceBadgeClass(abs.type)}">
                    ${escapeHtml(abs.type)} ${formatNumber(abs.stunden, 1)}h
                  </span>
                  ${missingBadge}
                </div>
                ${noteHtml}
              </div>
            `;
          } else if (workH > 0 && dayOvertime > 0) {
            // Fall 3: Arbeitszeit MIT Überstunden (>8h Mo-Do, >7h Fr oder Sa/So)
            const noteHtml = (manWork && manWork.note) ? `<div class="cell-note-text" title="${escapeHtml(manWork.note)}">${escapeHtml(manWork.note)}</div>` : '';

            cellInner = `
              <div class="matrix-cell-wrap">
                <div class="cell-hours-row">
                  <span class="work-num">${formatNumber(workH, 2)}</span>
                  <span class="badge-absence badge-overtime" title="Davon ${formatNumber(dayOvertime, 1)} Std. Überstunden (über Soll ${targetH}h)">
                    +${formatNumber(dayOvertime, 1).replace(',0', '')}h ÜSt
                  </span>
                </div>
                ${noteHtml}
              </div>
            `;
          } else if (workH > 0 && missingH > 0) {
            // Fall 4: Arbeitszeit, aber WENIGER als Sollzeit und noch keine Abwesenheit erfasst!
            const noteHtml = (manWork && manWork.note) ? `<div class="cell-note-text" title="${escapeHtml(manWork.note)}">${escapeHtml(manWork.note)}</div>` : '';

            cellInner = `
              <div class="matrix-cell-wrap cell-missing-highlight">
                <div class="cell-hours-row">
                  <span class="work-num" style="color: #c2410c;">${formatNumber(workH, 2)}</span>
                  <span class="badge-absence badge-missing" title="Sollzeit (${targetH}h) nicht erreicht: ${formatNumber(missingH, 1)}h fehlen! Klicken zum Klären">
                    ⚠️ -${formatNumber(missingH, 1).replace(',0', '')}h
                  </span>
                </div>
                ${noteHtml}
                <div class="cell-missing-text">Grund klären</div>
              </div>
            `;
          } else if (workH > 0 && manWork) {
            // Fall 5: Nachgebuchte Arbeitszeit ohne Überstunden/Minderarbeit
            const noteHtml = manWork.note ? `<div class="cell-note-text" title="${escapeHtml(manWork.note)}">${escapeHtml(manWork.note)}</div>` : '';

            cellInner = `
              <div class="matrix-cell-wrap">
                <div class="cell-hours-row">
                  <span class="work-num">${formatNumber(workH, 2)}</span>
                  <span class="badge-absence badge-arbeit" title="Arbeitszeit manuell nachgebucht">
                    +${formatNumber(manWork.stunden, 1).replace(',0', '')}h AZ
                  </span>
                </div>
                ${noteHtml}
              </div>
            `;
          } else if (workH > 0) {
            // Fall 6: Reine reguläre Arbeitsstunden aus Bewegungsdaten (Soll genau erfüllt)
            cellInner = `
              <div class="matrix-cell-wrap">
                <div class="cell-hours-row">
                  <span class="day-cell has-hours">${formatNumber(workH, 2)}</span>
                </div>
              </div>
            `;
          } else if (missingH > 0) {
            // Fall 7: 0 Stunden erfasst an einem Werktag für aktiven Mitarbeiter!
            cellInner = `
              <div class="matrix-cell-wrap cell-missing-highlight">
                <span class="badge-absence badge-missing" title="0 Std. erfasst (Soll: ${targetH}h). Klicken zum Klären!">
                  ⚠️ 0h (Grund fehlt)
                </span>
              </div>
            `;
          } else {
            // Fall 8: Keine Stunden (Wochenende oder inaktiver Mitarbeiter)
            cellInner = `<span class="zero-dash">–</span>`;
          }

          cellsHtml += `
            <td class="text-right matrix-cell-clickable ${tdClass}" 
                onclick="window.openManualEntryModal('${escapeHtml(emp.ressourcennummer)}', '${targetDayIso}', ${missingH > 0 ? missingH : 'null'})"
                title="${escapeHtml(tooltipParts.join(' // ')) || 'Klicken zum Nachbuchen von Arbeitszeit oder Erfassen von Krank, Urlaub, Überstunden'}">
              ${cellInner}
            </td>
          `;
        }

        let breakdownSubtext = '';
        if (emp.absenceHours > 0) {
          breakdownSubtext = `<div class="sub-hours-info" style="font-size: 0.72rem; color: #64748b; font-weight: normal;">${formatNumber(emp.workHours, 1)}h Arb. + ${formatNumber(emp.absenceHours, 1)}h Sond.</div>`;
        }

        let overtimeSubtext = '';
        if (emp.overtimeHours > 0) {
          overtimeSubtext = `<div class="emp-overtime-sub">davon ${formatNumber(emp.overtimeHours, 2)} Std. Überstd.</div>`;
        }

        return `
          <tr>
            <td><strong class="res-id">${escapeHtml(emp.ressourcennummer)}</strong></td>
            <td><strong>${escapeHtml(emp.name)}</strong></td>
            ${cellsHtml}
            <td class="text-right font-bold" style="color: #1e3a8a; background-color: #f8fafc;">
              ${formatNumber(emp.workHours, 2)} Std.
              ${overtimeSubtext}
              ${breakdownSubtext}
            </td>
          </tr>
        `;
      }).join('');

      for (let d = 0; d < 7; d++) {
        document.getElementById(`matrixSumDay${d}`).textContent = formatNumber(dayTotals[d], 2);
      }
      document.getElementById('matrixSumTotal').innerHTML = `
        <div>${formatNumber(kwTotalWorkHours, 2)} Std.</div>
        ${kwTotalOvertimeHours > 0 ? `<div style="font-size: 0.72rem; color: #b45309; font-weight: 700;">davon ${formatNumber(kwTotalOvertimeHours, 2)}h Überstunden</div>` : ''}
      `;
    }

    // ==========================================
    // TABELLE 2: KOMPAKTE ZUSAMMENFASSUNG RENDERN
    // ==========================================
    const kompaktBody = document.getElementById('wochenTableBody');
    if (activeWeekEmployees.length === 0) {
      kompaktBody.innerHTML = `
        <tr class="empty-row">
          <td colspan="9">
            <div class="empty-state">
              <p>Keine Einträge für diese Kalenderwoche.</p>
            </div>
          </td>
        </tr>
      `;
      document.getElementById('wochenFootSonderSum').textContent = '0,00 Std.';
      document.getElementById('wochenFootStundenSum').textContent = '0,00 Std.';
      const footOt = document.getElementById('wochenFootOvertimeSum');
      if (footOt) footOt.textContent = '0,00 Std.';
      document.getElementById('wochenFootGesamtSum').textContent = '0,00 Std.';
      document.getElementById('wochenFootTageSum').textContent = '0';
    } else {
      kompaktBody.innerHTML = activeWeekEmployees.map((emp, index) => {
        let absenceBreakdownHtml = '';
        if (emp.absenceList.length > 0) {
          const itemsHtml = emp.absenceList.map(item => `
            <div class="absence-item-row">
              <span class="badge-absence ${getAbsenceBadgeClass(item.type)}">${escapeHtml(item.type)}</span>
              <span class="absence-day">${escapeHtml(item.dayName)}, ${escapeHtml(item.dateDisplay)}:</span>
              <span class="absence-hours">${formatNumber(item.stunden, 1)} Std.</span>
              ${item.note ? `<span class="absence-note-tag" title="${escapeHtml(item.note)}">${escapeHtml(item.note)}</span>` : ''}
            </div>
          `).join('');

          absenceBreakdownHtml = `<div class="absence-breakdown-wrapper">${itemsHtml}</div>`;
        } else {
          absenceBreakdownHtml = `<span class="no-absence-text">– keine Sonderzeiten (regulärer Dienst) –</span>`;
        }

        const absenceHoursDisplay = emp.absenceHours > 0 
          ? `<strong style="color: #64748b;">${formatNumber(emp.absenceHours, 2)} Std.</strong>`
          : `<span class="text-muted">–</span>`;

        const overtimeDisplay = emp.overtimeHours > 0
          ? `<strong style="color: #b45309;">${formatNumber(emp.overtimeHours, 2)} Std.</strong>`
          : `<span class="text-muted">–</span>`;

        return `
          <tr>
            <td class="text-center" style="color: var(--text-muted); font-size: 0.8rem;">${index + 1}</td>
            <td>
              <span style="font-weight: 700; color: #0f172a;">${escapeHtml(emp.name)}</span>
            </td>
            <td>
              <span class="res-id" style="font-size: 0.82rem;">${escapeHtml(emp.ressourcennummer)}</span>
            </td>
            <td class="text-center screen-only-col">
              <span class="badge badge-info">${emp.datesSet.size} Tag${emp.datesSet.size === 1 ? '' : 'e'}</span>
            </td>
            <td>
              ${absenceBreakdownHtml}
            </td>
            <td class="text-right font-medium">
              ${absenceHoursDisplay}
            </td>
            <td class="text-right font-bold" style="color: #1e3a8a;">
              ${formatNumber(emp.workHours, 2)} Std.
            </td>
            <td class="text-right font-bold" style="background-color: #fffbeb;">
              ${overtimeDisplay}
            </td>
            <td class="text-right font-bold" style="background-color: #f8fafc; color: #0f172a;">
              ${formatNumber(emp.totalHours, 2)} Std.
            </td>
          </tr>
        `;
      }).join('');

      let totalDays = 0;
      for (const emp of activeWeekEmployees) {
        totalDays += emp.datesSet.size;
      }
      document.getElementById('wochenFootSonderSum').textContent = `${formatNumber(kwTotalAbsenceHours, 2)} Std.`;
      document.getElementById('wochenFootStundenSum').textContent = `${formatNumber(kwTotalWorkHours, 2)} Std.`;
      const footOt = document.getElementById('wochenFootOvertimeSum');
      if (footOt) footOt.textContent = `${formatNumber(kwTotalOvertimeHours, 2)} Std.`;
      document.getElementById('wochenFootGesamtSum').textContent = `${formatNumber(kwTotalHours, 2)} Std.`;
      document.getElementById('wochenFootTageSum').textContent = `${totalDays} Einsatztage`;
    }
  }

  function updatePrintHeader(selKw, weekItems) {
    let kwTitle = selKw ? `Kalenderwoche: ${selKw}` : 'Kalenderwoche: --';
    let dateRange = '--';

    if (weekItems.length > 0 && weekItems[0].kwInfo) {
      kwTitle = weekItems[0].kwInfo.label;
      dateRange = weekItems[0].kwInfo.rangeText;
    } else if (selKw) {
      const match = selKw.match(/^(\d{4})-W(\d{2})$/);
      if (match) {
        kwTitle = `Kalenderwoche ${match[2]} / ${match[1]}`;
      }
    }

    const now = new Date();
    const generatedAt = `${formatGermanDateTime(now)} Uhr`;

    document.getElementById('printHeaderKwTitle').textContent = kwTitle;
    document.getElementById('printHeaderDateRange').textContent = dateRange;
    document.getElementById('printHeaderGeneratedAt').textContent = generatedAt;
  }

  // ==========================================================================
  // 8. EVENT LISTENERS & MODAL MANAGEMENT
  // ==========================================================================
  function setupDomEvents() {
    document.querySelectorAll('.tab-button').forEach(btn => {
      btn.addEventListener('click', () => {
        const tabId = btn.getAttribute('data-tab');
        switchTab(tabId);
      });
    });

    const fileStamm = document.getElementById('fileInputStamm');
    fileStamm.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        parseStammdatenFile(e.target.files[0]);
        e.target.value = '';
      }
    });

    const fileBewegung = document.getElementById('fileInputBewegung');
    fileBewegung.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        parseBewegungsdatenFile(e.target.files[0]);
        e.target.value = '';
      }
    });

    setupDragAndDrop('dropZoneStamm', 'cardStammdaten', parseStammdatenFile);
    setupDragAndDrop('dropZoneBewegung', 'cardBewegungsdaten', parseBewegungsdatenFile);

    const filterFrom = document.getElementById('filterDateFrom');
    const filterTo = document.getElementById('filterDateTo');
    const filterCat = document.getElementById('filterCategory');
    const filterSearch = document.getElementById('filterSearch');
    const btnClearSearch = document.getElementById('btnClearSearch');

    filterFrom.addEventListener('change', () => {
      state.filters.dateFrom = filterFrom.value;
      renderDetailTab();
    });

    filterTo.addEventListener('change', () => {
      state.filters.dateTo = filterTo.value;
      renderDetailTab();
    });

    filterCat.addEventListener('change', () => {
      state.filters.category = filterCat.value;
      renderDetailTab();
    });

    filterSearch.addEventListener('input', () => {
      state.filters.search = filterSearch.value;
      btnClearSearch.style.display = filterSearch.value ? 'block' : 'none';
      renderDetailTab();
    });

    btnClearSearch.addEventListener('click', () => {
      filterSearch.value = '';
      state.filters.search = '';
      btnClearSearch.style.display = 'none';
      renderDetailTab();
      filterSearch.focus();
    });

    document.getElementById('btnResetFilters').addEventListener('click', () => {
      filterFrom.value = '';
      filterTo.value = '';
      filterCat.value = '';
      filterSearch.value = '';
      state.filters = { dateFrom: '', dateTo: '', category: '', search: '' };
      btnClearSearch.style.display = 'none';
      renderDetailTab();
      showToast('Alle Filter zurückgesetzt.', 'info');
    });

    document.querySelectorAll('#detailTable th.sortable').forEach(th => {
      th.addEventListener('click', () => {
        const field = th.getAttribute('data-sort');
        if (state.sort.field === field) {
          state.sort.asc = !state.sort.asc;
        } else {
          state.sort.field = field;
          state.sort.asc = true;
        }

        document.querySelectorAll('#detailTable th.sortable').forEach(h => {
          h.classList.remove('sort-asc', 'sort-desc');
        });
        th.classList.add(state.sort.asc ? 'sort-asc' : 'sort-desc');

        renderDetailTab();
      });
    });

    document.getElementById('selectKalenderwoche').addEventListener('change', (e) => {
      state.selectedKwKey = e.target.value;
      renderWochenTab();
    });

    const chkAllEmp = document.getElementById('chkShowAllEmployees');
    if (chkAllEmp) {
      chkAllEmp.addEventListener('change', (e) => {
        state.showAllEmployeesInWeekly = e.target.checked;
        renderWochenTab();
      });
    }

    const wocheSearch = document.getElementById('wocheSearchInput');
    if (wocheSearch) {
      wocheSearch.addEventListener('input', (e) => {
        state.weeklySearch = e.target.value;
        renderWochenTab();
      });
    }

    document.getElementById('btnPrintWochenbericht').addEventListener('click', () => {
      triggerPrintWochenbericht();
    });

    document.getElementById('btnExportDetailExcel').addEventListener('click', () => {
      exportDetailToExcel();
    });

    document.getElementById('btnDemoData').addEventListener('click', () => {
      loadDemoData();
    });

    const modalSettings = document.getElementById('modalSettings');
    document.getElementById('btnOpenSettings').addEventListener('click', () => {
      updateStammdatenStatusUI();
      modalSettings.style.display = 'flex';
    });
    document.getElementById('btnCloseSettings').addEventListener('click', () => {
      modalSettings.style.display = 'none';
    });
    document.getElementById('btnModalCloseFooter').addEventListener('click', () => {
      modalSettings.style.display = 'none';
    });
    modalSettings.addEventListener('click', (e) => {
      if (e.target === modalSettings) {
        modalSettings.style.display = 'none';
      }
    });

    document.getElementById('btnModalUploadNew').addEventListener('click', () => {
      document.getElementById('fileInputStamm').click();
      modalSettings.style.display = 'none';
    });

    document.getElementById('btnLoadDefaultStamm').addEventListener('click', () => {
      saveStammdatenToStorage(DEMO_STAMMDATEN, DEMO_AUFTRAEGE, 'Muster_Stammdaten.xlsx');
      updateStammdatenStatusUI();
    });

    document.getElementById('btnClearManualOnly').addEventListener('click', () => {
      if (confirm('Möchtest du alle eingetragenen Sonderzeiten und Abwesenheiten löschen?')) {
        state.manualEntries = [];
        saveManualEntriesToStorage();
        showToast('Alle Abwesenheiten wurden gelöscht.', 'info');
      }
    });

    document.getElementById('btnClearStammStorage').addEventListener('click', () => {
      if (confirm('Möchtest du alle gespeicherten Stammdaten, Aufträge und Abwesenheiten wirklich löschen?')) {
        clearStammdatenStorage();
      }
    });

    // Modal für manuelle Sonderzeiten
    const modalManual = document.getElementById('modalManualEntry');
    document.getElementById('btnOpenManualEntry').addEventListener('click', () => {
      openManualEntryModal();
    });
    document.getElementById('btnCloseManualModal').addEventListener('click', () => {
      modalManual.style.display = 'none';
    });
    document.getElementById('btnCancelManualModal').addEventListener('click', () => {
      modalManual.style.display = 'none';
    });
    modalManual.addEventListener('click', (e) => {
      if (e.target === modalManual) {
        modalManual.style.display = 'none';
      }
    });

    document.getElementById('btnSaveManualEntry').addEventListener('click', () => {
      submitManualEntryForm();
    });

    document.getElementById('btnDeleteManualEntry').addEventListener('click', () => {
      const editId = document.getElementById('manualEditId').value;
      if (editId) {
        deleteManualEntryById(editId);
        modalManual.style.display = 'none';
      }
    });

    document.getElementById('manualSelectMitarbeiter').addEventListener('change', checkAndPrepopulateExistingEntry);
    document.getElementById('manualDateInput').addEventListener('change', checkAndPrepopulateExistingEntry);
    document.getElementById('manualDateInput').addEventListener('input', checkAndPrepopulateExistingEntry);

    window.openManualEntryModal = openManualEntryModal;
    window.deleteManualEntryDirect = (id) => {
      deleteManualEntryById(id);
    };
  }

  function getDefaultHoursForDate(dateStrOrObj) {
    if (!dateStrOrObj) return 8.0;
    const d = (dateStrOrObj instanceof Date) ? dateStrOrObj : parseAnyDate(dateStrOrObj);
    if (!d || isNaN(d.getTime())) return 8.0;
    // d.getDay(): 0=So, 1=Mo, 2=Di, 3=Mi, 4=Do, 5=Fr, 6=Sa
    return (d.getDay() === 5) ? 7.0 : 8.0;
  }

  function getDailyTargetHours(dayIdx) {
    // dayIdx: 0=Mo, 1=Di, 2=Mi, 3=Do, 4=Fr, 5=Sa, 6=So
    if (dayIdx >= 0 && dayIdx <= 3) return 8.0; // Mo - Do
    if (dayIdx === 4) return 7.0;               // Fr
    return 0.0;                                // Sa, So (Wochenende)
  }

  function getTargetHoursForDate(dateStrOrObj) {
    const d = (dateStrOrObj instanceof Date) ? dateStrOrObj : parseAnyDate(dateStrOrObj);
    if (!d || isNaN(d.getTime())) return 8.0;
    const day = d.getDay(); // 0=So, 1=Mo, 2=Di, 3=Mi, 4=Do, 5=Fr, 6=Sa
    if (day >= 1 && day <= 4) return 8.0; // Mo - Do
    if (day === 5) return 7.0;            // Fr
    return 0.0;                           // Sa, So
  }

  function getRecordedHoursForEmployeeAndDate(ressourcennummer, dateIso) {
    if (!ressourcennummer || !dateIso) return { workH: 0, absH: 0 };
    const rKey = normalizeResourceKey(ressourcennummer);
    let workH = 0;
    let absH = 0;

    for (const b of state.bewegungsdaten) {
      if (normalizeResourceKey(b.ressourcennummer) === rKey && b.dateIso === dateIso) {
        workH += (b.stunden || 0);
      }
    }

    for (const m of state.manualEntries) {
      if (normalizeResourceKey(m.ressourcennummer) === rKey && m.dateIso === dateIso) {
        if (m.type === 'Arbeitszeit' || String(m.type).toLowerCase().includes('arbeit')) {
          workH += (m.stunden || 0);
        } else {
          absH += (m.stunden || 0);
        }
      }
    }

    return { workH, absH };
  }

  function renderKlaerungsBox(activeEmployees, selKw) {
    const box = document.getElementById('wocheKlaerungBox');
    const badge = document.getElementById('klaerungStatusBadge');
    const content = document.getElementById('klaerungContent');
    if (!box || !badge || !content) return;

    const allUnclarified = [];
    for (const emp of activeEmployees) {
      if (emp.unclarifiedDays && emp.unclarifiedDays.length > 0) {
        for (const item of emp.unclarifiedDays) {
          allUnclarified.push({
            emp: emp,
            ...item
          });
        }
      }
    }

    if (allUnclarified.length === 0) {
      box.className = 'klaerung-box all-clear no-print';
      badge.textContent = '✅ Vollständig geklärt';
      content.innerHTML = `
        <div style="display: flex; align-items: center; gap: 0.5rem; font-size: 0.88rem; color: #166534; font-weight: 500; margin-top: 0.4rem;">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
            <polyline points="22 4 12 14.01 9 11.01"></polyline>
          </svg>
          <span>Alle regulären Arbeitszeiten der Woche (${escapeHtml(selKw || '--')}) sind vollständig erfasst und begründet. Keine offenen Fehlzeiten.</span>
        </div>
      `;
    } else {
      box.className = 'klaerung-box no-print';
      badge.textContent = `⚠️ ${allUnclarified.length} offene${allUnclarified.length === 1 ? 'r Tag' : ' Tage'}`;
      
      const itemsHtml = allUnclarified.map(item => `
        <div class="klaerung-item">
          <div class="klaerung-item-info">
            <div class="klaerung-item-name">${escapeHtml(item.emp.name)} <span style="font-weight: normal; color: #64748b;">(${escapeHtml(item.emp.ressourcennummer)})</span></div>
            <div class="klaerung-item-details">
              <strong>${escapeHtml(item.dayName)}, ${escapeHtml(item.dateDisplay)}:</strong> 
              ${item.workH > 0 ? `${formatNumber(item.workH, 1)} von ${formatNumber(item.targetH, 1)} Std. erfasst` : `0 von ${formatNumber(item.targetH, 1)} Std. erfasst`}
              <span style="font-weight: 800; color: #dc2626; margin-left: 4px;">(${formatNumber(item.missingH, 1)} Std. offen)</span>
            </div>
          </div>
          <button type="button" class="btn-klaeren" 
                  onclick="window.openManualEntryModal('${escapeHtml(item.emp.ressourcennummer)}', '${item.dateIso}', ${item.missingH})"
                  title="Grund für fehlende ${formatNumber(item.missingH, 1)} Std. erfassen">
            ⚡ Grund erfassen
          </button>
        </div>
      `).join('');

      content.innerHTML = `
        <div style="font-size: 0.82rem; color: #7c2d12; margin-top: 0.4rem; margin-bottom: 0.6rem;">
          An folgenden Tagen wurde weniger als die reguläre Sollzeit (Mo–Do 8h, Fr 7h) gearbeitet. Bitte erfasse für die Differenz den jeweiligen Grund (Krankheit, Überstunden abgefeiert, sonstiges oder Arbeitszeit-Nachbuchung):
        </div>
        <div class="klaerung-items-list">
          ${itemsHtml}
        </div>
      `;
    }
  }

  function populateManualModalEmployees() {
    const sel = document.getElementById('manualSelectMitarbeiter');
    if (!sel) return;

    const currentVal = sel.value;
    const empMap = new Map();

    for (const item of state.stammdaten) {
      if (item.kategorie && item.kategorie.toLowerCase() === 'mitarbeiter' && item.ressourcennummer) {
        const key = normalizeResourceKey(item.ressourcennummer) || normalizeKey(item.ressourcennummer);
        if (!empMap.has(key)) {
          empMap.set(key, { id: item.ressourcennummer, label: `${item.name} (${item.ressourcennummer})` });
        }
      }
    }

    // Auch Mitarbeiter aus Bewegungsdaten ergänzen (falls noch nicht in Stammdaten vorhanden)
    for (const item of state.bewegungsdaten) {
      if (item.isMitarbeiter && item.ressourcennummer) {
        const key = normalizeResourceKey(item.ressourcennummer) || normalizeKey(item.ressourcennummer);
        if (!empMap.has(key)) {
          empMap.set(key, { id: item.ressourcennummer, label: `${item.name} (${item.ressourcennummer})` });
        }
      }
    }

    const employees = Array.from(empMap.values());
    employees.sort((a, b) => a.label.localeCompare(b.label, 'de'));

    let html = `<option value="">-- Mitarbeiter wählen --</option>`;
    for (const emp of employees) {
      const isSel = (normalizeResourceKey(emp.id) === normalizeResourceKey(currentVal));
      html += `<option value="${escapeHtml(emp.id)}" ${isSel ? 'selected' : ''}>${escapeHtml(emp.label)}</option>`;
    }
    sel.innerHTML = html;
  }

  function openManualEntryModal(ressourcennummer = '', dateIso = '', forcedMissingHours = null) {
    populateManualModalEmployees();

    const modal = document.getElementById('modalManualEntry');
    const selEmp = document.getElementById('manualSelectMitarbeiter');
    const dateInput = document.getElementById('manualDateInput');
    const typeSelect = document.getElementById('manualTypeSelect');
    const hoursInput = document.getElementById('manualHoursInput');
    const noteInput = document.getElementById('manualNoteInput');
    const editIdInput = document.getElementById('manualEditId');
    const btnDelete = document.getElementById('btnDeleteManualEntry');
    const existingInfo = document.getElementById('manualExistingInfo');
    const hintFriday = document.getElementById('hintFridayHours');
    const contextBanner = document.getElementById('manualContextBanner');

    if (ressourcennummer) {
      const normTarget = normalizeResourceKey(ressourcennummer);
      let found = false;
      for (const opt of selEmp.options) {
        if (normalizeResourceKey(opt.value) === normTarget) {
          selEmp.value = opt.value;
          found = true;
          break;
        }
      }
      if (!found) {
        selEmp.value = ressourcennummer;
      }
    }
    if (dateIso) {
      dateInput.value = dateIso;
    } else if (!dateInput.value) {
      dateInput.value = formatDateIso(new Date());
    }

    const defHours = getDefaultHoursForDate(dateInput.value);
    const targetH = getTargetHoursForDate(dateInput.value);
    if (hintFriday) {
      hintFriday.style.display = (defHours === 7.0) ? 'inline-block' : 'none';
    }

    const existing = findManualEntry(selEmp.value, dateInput.value);
    if (existing) {
      editIdInput.value = existing.id;
      typeSelect.value = existing.type;
      hoursInput.value = Number(existing.stunden).toFixed(1);
      noteInput.value = existing.note || '';
      btnDelete.style.display = 'block';
      existingInfo.style.display = 'block';
      if (contextBanner) contextBanner.style.display = 'none';
    } else {
      editIdInput.value = '';
      btnDelete.style.display = 'none';
      existingInfo.style.display = 'none';

      // Prüfen, ob an diesem Tag Stunden fehlen (Sollzeit nicht erreicht)
      const rec = getRecordedHoursForEmployeeAndDate(selEmp.value, dateInput.value);
      const isWeekday = (targetH > 0);
      const totalRec = rec.workH + rec.absH;
      const missingH = (forcedMissingHours !== null && forcedMissingHours !== undefined) 
        ? forcedMissingHours 
        : (isWeekday && totalRec < targetH ? (targetH - totalRec) : 0);

      if (missingH > 0 && isWeekday) {
        // Fall: Fehlzeit / Unterstunden an einem Werktag!
        hoursInput.value = Number(missingH).toFixed(1);
        typeSelect.value = (rec.workH > 0) ? 'Überstundenabbau' : 'Krank';
        noteInput.value = '';

        if (contextBanner) {
          contextBanner.style.display = 'block';
          contextBanner.innerHTML = `
            <div style="font-weight: 700; margin-bottom: 3px;">⚠️ Fehlzeiten-Nachfrage (Soll-Arbeitszeit: ${formatNumber(targetH, 1)} Std.)</div>
            <div>Bisher erfasst: <strong>${formatNumber(rec.workH, 1)} Std. Arbeit</strong>. Es fehlen <strong>${formatNumber(missingH, 1)} Std.</strong></div>
            <div style="margin-top: 4px; font-size: 0.78rem; opacity: 0.9;">
              Bitte wähle den Grund für die Minderarbeit (z. B. <strong>Überstunden abgefeiert / Abbau</strong>, <strong>Krankheit</strong>, <strong>Urlaub</strong> oder <strong>Arbeitszeit-Nachbuchung</strong>).
            </div>
          `;
        }
      } else {
        typeSelect.value = 'Arbeitszeit';
        hoursInput.value = defHours.toFixed(1);
        noteInput.value = '';
        if (contextBanner) contextBanner.style.display = 'none';
      }
    }

    modal.style.display = 'flex';
  }

  function checkAndPrepopulateExistingEntry() {
    const selEmp = document.getElementById('manualSelectMitarbeiter');
    const dateInput = document.getElementById('manualDateInput');
    const typeSelect = document.getElementById('manualTypeSelect');
    const hoursInput = document.getElementById('manualHoursInput');
    const noteInput = document.getElementById('manualNoteInput');
    const editIdInput = document.getElementById('manualEditId');
    const btnDelete = document.getElementById('btnDeleteManualEntry');
    const existingInfo = document.getElementById('manualExistingInfo');
    const hintFriday = document.getElementById('hintFridayHours');
    const contextBanner = document.getElementById('manualContextBanner');

    const defHours = getDefaultHoursForDate(dateInput.value);
    const targetH = getTargetHoursForDate(dateInput.value);
    if (hintFriday) {
      hintFriday.style.display = (defHours === 7.0) ? 'inline-block' : 'none';
    }

    const existing = findManualEntry(selEmp.value, dateInput.value);
    if (existing) {
      editIdInput.value = existing.id;
      typeSelect.value = existing.type;
      hoursInput.value = Number(existing.stunden).toFixed(1);
      noteInput.value = existing.note || '';
      btnDelete.style.display = 'block';
      existingInfo.style.display = 'block';
      if (contextBanner) contextBanner.style.display = 'none';
    } else {
      editIdInput.value = '';
      btnDelete.style.display = 'none';
      existingInfo.style.display = 'none';

      const rec = getRecordedHoursForEmployeeAndDate(selEmp.value, dateInput.value);
      const isWeekday = (targetH > 0);
      const totalRec = rec.workH + rec.absH;
      const missingH = (isWeekday && totalRec < targetH) ? (targetH - totalRec) : 0;

      if (missingH > 0 && isWeekday) {
        hoursInput.value = Number(missingH).toFixed(1);
        typeSelect.value = (rec.workH > 0) ? 'Überstundenabbau' : 'Krank';
        if (contextBanner) {
          contextBanner.style.display = 'block';
          contextBanner.innerHTML = `
            <div style="font-weight: 700; margin-bottom: 3px;">⚠️ Fehlzeiten-Nachfrage (Soll-Arbeitszeit: ${formatNumber(targetH, 1)} Std.)</div>
            <div>Bisher erfasst: <strong>${formatNumber(rec.workH, 1)} Std. Arbeit</strong>. Es fehlen <strong>${formatNumber(missingH, 1)} Std.</strong></div>
            <div style="margin-top: 4px; font-size: 0.78rem; opacity: 0.9;">
              Bitte wähle den Grund für die Minderarbeit (z. B. <strong>Überstunden abgefeiert / Abbau</strong>, <strong>Krankheit</strong>, <strong>Urlaub</strong> oder <strong>Arbeitszeit-Nachbuchung</strong>).
            </div>
          `;
        }
      } else {
        hoursInput.value = defHours.toFixed(1);
        if (contextBanner) contextBanner.style.display = 'none';
      }
    }
  }

  function submitManualEntryForm() {
    const selEmp = document.getElementById('manualSelectMitarbeiter');
    const dateInput = document.getElementById('manualDateInput');
    const typeSelect = document.getElementById('manualTypeSelect');
    const hoursInput = document.getElementById('manualHoursInput');
    const noteInput = document.getElementById('manualNoteInput');
    const editIdInput = document.getElementById('manualEditId');

    if (!selEmp.value) {
      alert('Bitte wähle einen Mitarbeiter aus.');
      selEmp.focus();
      return;
    }

    if (!dateInput.value) {
      alert('Bitte gib ein Datum an.');
      dateInput.focus();
      return;
    }

    const hours = parseFloat(hoursInput.value);
    if (isNaN(hours) || hours <= 0) {
      alert('Bitte gib eine gültige Stundenzahl größer 0 ein (z. B. 1, 4 oder 8).');
      hoursInput.focus();
      return;
    }

    addOrUpdateManualEntry({
      id: editIdInput.value || null,
      ressourcennummer: selEmp.value,
      dateIso: dateInput.value,
      type: typeSelect.value,
      stunden: hours,
      note: noteInput.value.trim()
    });

    document.getElementById('modalManualEntry').style.display = 'none';
  }

  function setupDragAndDrop(dropZoneId, cardId, fileHandler) {
    const dropZone = document.getElementById(dropZoneId);
    const card = document.getElementById(cardId);
    if (!dropZone || !card) return;

    ['dragenter', 'dragover'].forEach(eventName => {
      card.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        card.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      card.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        card.classList.remove('dragover');
      });
    });

    card.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files && files.length > 0) {
        fileHandler(files[0]);
      }
    });
  }

  function switchTab(tabId) {
    state.activeTab = tabId;

    document.querySelectorAll('.tab-button').forEach(btn => {
      if (btn.getAttribute('data-tab') === tabId) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    document.querySelectorAll('.tab-content').forEach(content => {
      if (content.id === tabId) {
        content.classList.add('active');
      } else {
        content.classList.remove('active');
      }
    });

    renderApp();
  }

  function renderApp() {
    renderDetailTab();
    renderWochenTab();
  }

  // ==========================================================================
  // 9. NATIVE DRUCKFUNKTION & PRINT VORBEREITUNG
  // ==========================================================================
  function triggerPrintWochenbericht() {
    if (state.activeTab !== 'tabWoche') {
      switchTab('tabWoche');
    }

    renderWochenTab();
    document.body.classList.add('printing-woche');

    setTimeout(() => {
      window.print();
      document.body.classList.remove('printing-woche');
    }, 150);
  }

  // ==========================================================================
  // 10. EXCEL EXPORT (GEFILTERTE DETAILANSICHT INKL. ORT & SONDERZEITEN)
  // ==========================================================================
  function exportDetailToExcel() {
    const filtered = getFilteredDetailData();
    const sorted = sortData(filtered);

    if (sorted.length === 0) {
      showToast('Keine Daten für den Excel-Export vorhanden.', 'info');
      return;
    }

    const exportRows = [
      ['Datum', 'Auftragsnummer', 'Ort / Baustelle', 'Ressourcennummer', 'Name', 'Kategorie', 'Stunden bzw. Mengen', 'Leistung', 'Beschreibung', 'Preis']
    ];

    for (const item of sorted) {
      exportRows.push([
        item.dateDisplay,
        item.auftragsnummer,
        item.ort,
        item.ressourcennummer,
        item.name,
        item.kategorie,
        item.stunden,
        item.leistung,
        item.beschreibung,
        item.preis
      ]);
    }

    try {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(exportRows);
      XLSX.utils.book_append_sheet(wb, ws, 'Gefilterte Auswertung');
      const filename = `Tagesberichte_Auswertung_${formatDateIso(new Date())}.xlsx`;
      XLSX.writeFile(wb, filename);
      showToast(`Excel-Export erfolgreich: ${filename}`, 'success');
    } catch (e) {
      console.error('Fehler beim Export:', e);
      showToast('Fehler beim Excel-Export: ' + e.message, 'error');
    }
  }

  // ==========================================================================
  // 11. DEMO-DATEN LADEN (MIT BEISPIEL-ABWESENHEITEN)
  // ==========================================================================
  function loadDemoData() {
    saveStammdatenToStorage(DEMO_STAMMDATEN, DEMO_AUFTRAEGE, 'Demo_Stammdaten.xlsx');
    state.manualEntries = [...DEMO_MANUAL_ENTRIES];
    saveManualEntriesToStorage();
    state.bewegungsdatenFilename = 'Demo_Tagesberichte.xlsx';
    processBewegungsdatenRows(DEMO_BEWEGUNGSDATEN_RAW);
    showToast('Demo-Daten geladen: 12 Ressourcen, 4 Orte, 3 Abwesenheiten & 35 Tagesberichte', 'success');
  }

  // ==========================================================================
  // 12. HILFSFUNKTIONEN (DATUM, ZAHLEN, FORMATIERUNG, BADGES)
  // ==========================================================================
  function parseAnyDate(val) {
    if (!val) return null;
    if (val instanceof Date) {
      return isNaN(val.getTime()) ? null : val;
    }

    if (typeof val === 'number') {
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      return isNaN(date.getTime()) ? null : date;
    }

    const str = String(val).trim();
    if (!str) return null;

    const isoMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      const d = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
      return isNaN(d.getTime()) ? null : d;
    }

    const deMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (deMatch) {
      const d = new Date(parseInt(deMatch[3], 10), parseInt(deMatch[2], 10) - 1, parseInt(deMatch[1], 10));
      return isNaN(d.getTime()) ? null : d;
    }

    const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (slashMatch) {
      const d = new Date(parseInt(slashMatch[3], 10), parseInt(slashMatch[1], 10) - 1, parseInt(slashMatch[2], 10));
      return isNaN(d.getTime()) ? null : d;
    }

    const fallback = new Date(str);
    return isNaN(fallback.getTime()) ? null : fallback;
  }

  function parseGermanNumber(val) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (!val) return 0;
    let clean = String(val).replace(/[^0-9,.-]/g, '').trim();
    if (clean.includes(',') && clean.includes('.')) {
      clean = clean.replace(/\./g, '').replace(',', '.');
    } else if (clean.includes(',')) {
      clean = clean.replace(',', '.');
    }
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  }

  function formatDateIso(d) {
    if (!d) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function formatGermanDate(d) {
    if (!d) return '--';
    const day = String(d.getDate()).padStart(2, '0');
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const y = d.getFullYear();
    return `${day}.${m}.${y}`;
  }

  function formatGermanDateTime(d) {
    if (!d) return '--';
    const dateStr = formatGermanDate(d);
    const h = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return `${dateStr}, ${h}:${min}`;
  }

  function getWeekdayShort(d) {
    const days = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
    return days[d.getDay()] || '';
  }

  function formatNumber(num, decimals = 2) {
    const n = Number(num) || 0;
    return n.toLocaleString('de-DE', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
  }

  function formatCurrency(num) {
    const n = Number(num) || 0;
    return n.toLocaleString('de-DE', {
      style: 'currency',
      currency: 'EUR'
    });
  }

  function getCategoryBadgeClass(kategorie) {
    if (!kategorie) return 'badge-unknown';
    const lower = String(kategorie).trim().toLowerCase();
    if (lower === 'mitarbeiter') return 'badge-mitarbeiter';
    if (lower === 'fahrzeug') return 'badge-fahrzeug';
    if (lower === 'maschine') return 'badge-maschine';
    return 'badge-unknown';
  }

  function getAbsenceBadgeClass(type) {
    if (!type) return 'badge-sonstiges';
    const lower = String(type).trim().toLowerCase();
    if (lower.includes('arbeit')) return 'badge-arbeit';
    if (lower.includes('krank')) return 'badge-krank';
    if (lower.includes('urlaub')) return 'badge-urlaub';
    if (lower.includes('überstunde') || lower.includes('gleitzeit')) return 'badge-ueberstunden';
    if (lower.includes('schulung') || lower.includes('lehrgang')) return 'badge-schulung';
    return 'badge-sonstiges';
  }

  function formatAbsenceShort(type, hours) {
    const hStr = formatNumber(hours, 1).replace(',0', '') + 'h';
    const lower = String(type).trim().toLowerCase();
    if (lower.includes('arbeit')) return `+${hStr} AZ`;
    if (lower.includes('krank')) return `+${hStr} K`;
    if (lower.includes('urlaub')) return `+${hStr} U`;
    if (lower.includes('überstunde') || lower.includes('gleitzeit')) return `+${hStr} ÜA`;
    if (lower.includes('schulung')) return `+${hStr} S`;
    return `+${hStr} ${type}`;
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
      <span>${escapeHtml(message)}</span>
      <span style="cursor: pointer; opacity: 0.7; font-weight: bold;">&times;</span>
    `;

    toast.querySelector('span:last-child').addEventListener('click', () => {
      toast.remove();
    });

    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

})();

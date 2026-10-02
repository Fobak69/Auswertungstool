// ==============================================================================
// Excel-Prüfer Web Controller (100% Client-Side für Netlify & Browser)
// ==============================================================================
console.log("[APP.JS] *** VERSION 3.0 GELADEN - 2026-10-02T15:00 ***");


document.addEventListener("DOMContentLoaded", () => {
  // --- Globaler Drag & Drop Schutz ---
  // Verhindert, dass Chrome/Edge bei Drag & Drop (oder versehentlichem Danebenwerfen)
  // die Datei im Windows-Download-Ordner ablegt oder die Webseite verlässt.
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    window.addEventListener(eventName, (e) => {
      e.preventDefault();
    }, false);
    document.addEventListener(eventName, (e) => {
      e.preventDefault();
    }, false);
  });

  // App-Zustand im Arbeitsspeicher
  const state = {
    refWorkbook: null,
    targetWorkbook: null,
    refFileName: "",
    targetFileName: "",
    currentRefSheet: "",
    currentTargetSheet: "",
    mapping: [],
    allResults: [],
    currentFilter: "errors",
    searchQuery: "",
    appliedCorrections: {}, // cellKey ("row_col") -> newVal
    lastTargetBuffer: null,
    staffList: [],
    staffMap: {},
    timesheetData: null,
    lastInsertedRow: null,
    modalRowMode: "insert", // "insert" oder "edit"
    modalEditRowIdx: null,
    modalPrefill: null,
    modalCaller: "timesheet",
    highlightEmployeeRes: null
  };

  // DOM Elemente
  const dropRef = document.getElementById("drop-ref");
  const inputRef = document.getElementById("input-ref");
  const refFileInfo = document.getElementById("ref-file-info");
  const refFilename = document.getElementById("ref-filename");
  const selectRefSheet = document.getElementById("select-ref-sheet");
  const refStorageBadge = document.getElementById("ref-storage-badge");
  const refInfoDetail = document.getElementById("ref-info-detail");
  const btnChangeRef = document.getElementById("btn-change-ref");
  const btnClearRef = document.getElementById("btn-clear-ref");

  const dropTgt = document.getElementById("drop-tgt");
  const inputTgt = document.getElementById("input-tgt");
  const tgtFileInfo = document.getElementById("tgt-file-info");
  const tgtFilename = document.getElementById("tgt-filename");
  const selectTgtSheet = document.getElementById("select-tgt-sheet");
  const tgtStatusBadge = document.getElementById("tgt-status-badge");
  const btnChangeTgt = document.getElementById("btn-change-tgt");

  // Mitarbeiter & Personal DOM
  const dropStaff = document.getElementById("drop-staff");
  const inputStaff = document.getElementById("input-staff");
  const staffFileInfo = document.getElementById("staff-file-info");
  const staffInfoTitle = document.getElementById("staff-info-title");
  const staffInfoDetail = document.getElementById("staff-info-detail");
  const staffDropZoneContent = document.getElementById("staff-drop-zone-content");
  const staffStorageBadge = document.getElementById("staff-storage-badge");
  const btnChangeStaff = document.getElementById("btn-change-staff");
  const btnClearStaff = document.getElementById("btn-clear-staff");
  const btnDownloadStaffTemplate = document.getElementById("btn-download-staff-template");

  // Arbeitszeiten / Timesheet DOM
  const timesheetSection = document.getElementById("timesheet-section");
  const selectTsDateCol = document.getElementById("select-ts-date-col");
  const selectTsHoursCol = document.getElementById("select-ts-hours-col");
  const selectTsResourceCol = document.getElementById("select-ts-resource-col");
  const chkTsOnlyStaff = document.getElementById("chk-ts-only-staff");
  const tsStaffStatusBadge = document.getElementById("ts-staff-status-badge");
  const timesheetThead = document.getElementById("timesheet-thead");
  const timesheetTbody = document.getElementById("timesheet-tbody");
  const timesheetTfoot = document.getElementById("timesheet-tfoot");
  const btnExportTimesheet = document.getElementById("btn-export-timesheet");

  const btnLoadDemo = document.getElementById("btn-load-demo");
  const configSection = document.getElementById("config-section");
  const mappingTbody = document.getElementById("mapping-tbody");
  const checkAllCols = document.getElementById("check-all-cols");
  const btnStartCheck = document.getElementById("btn-start-check");

  const resultsSection = document.getElementById("results-section");
  const resultsTbody = document.getElementById("results-tbody");
  const tableSearchInput = document.getElementById("table-search-input");
  const filterPills = document.getElementById("filter-pills");

  const btnGeneralApply = document.getElementById("btn-general-apply");
  const btnFixZahlendreher = document.getElementById("btn-fix-zahlendreher");
  const btnResetFixes = document.getElementById("btn-reset-fixes");
  const btnExportExcel = document.getElementById("btn-export-excel");
  const btnExportReport = document.getElementById("btn-export-report");
  const toast = document.getElementById("toast");

  // Ansichten-Umschalter (Tabs)
  const tabBtnDiagnostics = document.getElementById("tab-btn-diagnostics");
  const tabBtnFulltable = document.getElementById("tab-btn-fulltable");
  const viewDiagnostics = document.getElementById("view-diagnostics");
  const viewFulltable = document.getElementById("view-fulltable");
  const tabBadgeErrors = document.getElementById("tab-badge-errors");
  const tabBadgeTotalRows = document.getElementById("tab-badge-total-rows");

  // Gesamte Excel-Datei Filter & Tabelle
  const filterFullStaff = document.getElementById("filter-full-staff");
  const filterFullDate = document.getElementById("filter-full-date");
  const filterFullStatus = document.getElementById("filter-full-status");
  const filterFullSearch = document.getElementById("filter-full-search");
  const fulltableCountsText = document.getElementById("fulltable-counts-text");
  const btnResetFullFilters = document.getElementById("btn-reset-full-filters");
  const fulltableThead = document.getElementById("fulltable-thead");
  const fulltableTbody = document.getElementById("fulltable-tbody");

  // Zeile nachtragen / Bearbeiten Modal DOM
  const modalInsertRow = document.getElementById("modal-insert-row");
  const modalIcon = document.getElementById("modal-icon");
  const modalTitle = document.getElementById("modal-title");
  const modalDesc = document.getElementById("modal-desc");
  const modalEditBanner = document.getElementById("modal-edit-banner");
  const modalEditRowNum = document.getElementById("modal-edit-row-num");
  const btnSwitchToInsert = document.getElementById("btn-switch-to-insert");
  const modalInsertOnlySections = document.getElementById("modal-insert-only-sections");
  const modalFieldsTitle = document.getElementById("modal-fields-title");
  const btnCloseInsertModal = document.getElementById("btn-close-insert-modal");
  const btnCancelInsertRow = document.getElementById("btn-cancel-insert-row");
  const btnConfirmInsertRow = document.getElementById("btn-confirm-insert-row");
  const btnConfirmIcon = document.getElementById("btn-confirm-icon");
  const btnConfirmInsertText = document.getElementById("btn-confirm-insert-text");
  const btnOpenInsertRowActionbar = document.getElementById("btn-open-insert-row-actionbar");
  const btnOpenInsertRowFulltable = document.getElementById("btn-open-insert-row-fulltable");
  const btnOpenInsertRowTimesheet = document.getElementById("btn-open-insert-row-timesheet");
  const insertTargetRowSelect = document.getElementById("insert-target-row-select");
  const insertTemplateRowSelect = document.getElementById("insert-template-row-select");
  const insertRowSearch = document.getElementById("insert-row-search");
  const insertAfterRowContainer = document.getElementById("insert-after-row-container");
  const insertAutoHint = document.getElementById("insert-auto-hint");
  const insertFieldsGrid = document.getElementById("insert-fields-grid");

  // Buchung auswählen Modal (wenn mehrere Buchungen an einem Tag existieren)
  const modalChooseBooking = document.getElementById("modal-choose-booking");
  const btnCloseChooseModal = document.getElementById("btn-close-choose-modal");
  const btnCloseChoose = document.getElementById("btn-close-choose");
  const btnChooseAddNew = document.getElementById("btn-choose-add-new");
  const chooseBookingTitle = document.getElementById("choose-booking-title");
  const chooseBookingSubtitle = document.getElementById("choose-booking-subtitle");
  const chooseBookingList = document.getElementById("choose-booking-list");

  // --- IndexedDB Speicher für die Referenzdatei ---
  const IDB_NAME = "ExcelPrueferStorage";
  const IDB_VERSION = 1;
  const IDB_STORE = "app_data";
  const IDB_KEY_REF = "saved_reference_file";

  function openAppDB() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) {
        reject(new Error("IndexedDB wird von diesem Browser nicht unterstützt."));
        return;
      }
      const req = window.indexedDB.open(IDB_NAME, IDB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(IDB_STORE)) {
          db.createObjectStore(IDB_STORE);
        }
      };
      req.onsuccess = (e) => resolve(e.target.result);
      req.onerror = (e) => reject(e.target.error || new Error("Konnte IndexedDB nicht öffnen"));
    });
  }

  async function saveRefFileToStorage(fileName, buffer, selectedSheet) {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readwrite");
        const store = tx.objectStore(IDB_STORE);
        const data = {
          fileName: fileName,
          buffer: buffer,
          selectedSheet: selectedSheet || "",
          savedAt: new Date().toISOString()
        };
        const req = store.put(data, IDB_KEY_REF);
        req.onsuccess = () => resolve(true);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.warn("Fehler beim Speichern der Referenzdatei in IndexedDB:", err);
      return false;
    }
  }

  async function updateRefSheetInStorage(selectedSheet) {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readwrite");
        const store = tx.objectStore(IDB_STORE);
        const getReq = store.get(IDB_KEY_REF);
        getReq.onsuccess = () => {
          if (getReq.result) {
            const item = getReq.result;
            item.selectedSheet = selectedSheet;
            store.put(item, IDB_KEY_REF);
          }
          resolve(true);
        };
        getReq.onerror = () => resolve(false);
      });
    } catch (err) {
      console.warn("Konnte gewähltes Arbeitsblatt nicht im Speicher aktualisieren:", err);
    }
  }

  async function loadRefFileFromStorage() {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readonly");
        const store = tx.objectStore(IDB_STORE);
        const req = store.get(IDB_KEY_REF);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.warn("Fehler beim Laden der Referenzdatei aus IndexedDB:", err);
      return null;
    }
  }

  async function deleteRefFileFromStorage() {
    try {
      const db = await openAppDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, "readwrite");
        const store = tx.objectStore(IDB_STORE);
        const req = store.delete(IDB_KEY_REF);
        req.onsuccess = () => resolve(true);
        req.onerror = (e) => reject(e.target.error);
      });
    } catch (err) {
      console.warn("Fehler beim Löschen der Referenzdatei aus IndexedDB:", err);
      return false;
    }
  }

  async function restoreSavedReferenceFile() {
    try {
      const saved = await loadRefFileFromStorage();
      if (saved && saved.buffer) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(saved.buffer);
        state.refWorkbook = wb;
        state.refFileName = saved.fileName || "Referenz_Stammdaten.xlsx";

        if (refFilename) refFilename.textContent = state.refFileName;
        populateSheetSelect(selectRefSheet, wb.worksheets);

        if (saved.selectedSheet && wb.worksheets.some(ws => ws.name === saved.selectedSheet)) {
          selectRefSheet.value = saved.selectedSheet;
        }
        state.currentRefSheet = selectRefSheet.value;

        if (refFileInfo) refFileInfo.classList.remove("hidden");
        const dropContent = dropRef ? dropRef.querySelector(".drop-zone-content") : null;
        if (dropContent) dropContent.classList.add("hidden");

        if (refStorageBadge) {
          refStorageBadge.textContent = "Im Browser gespeichert";
          refStorageBadge.className = "badge badge-ok";
        }
        if (refInfoDetail) {
          refInfoDetail.textContent = "💾 Im Browser gespeichert (bleibt dauerhaft erhalten)";
        }

        // Mitarbeiter automatisch aus Referenzdatei laden
        extractStaffFromReferenceWorkbook(wb);

        checkReadyForConfig();
        console.log(`Referenzdatei "${state.refFileName}" aus dem Browser-Speicher wiederhergestellt.`);
      }
    } catch (err) {
      console.warn("Konnte gespeicherte Referenzdatei nicht laden:", err);
    }
  }

  // --- Drag & Drop Einrichten ---
  setupDropZone(dropRef, inputRef, async (file) => {
    try {
      state.refFileName = file.name;
      const arrayBuffer = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      state.refWorkbook = wb;

      refFilename.textContent = file.name;
      populateSheetSelect(selectRefSheet, wb.worksheets);
      state.currentRefSheet = selectRefSheet.value;
      refFileInfo.classList.remove("hidden");
      dropRef.querySelector(".drop-zone-content").classList.add("hidden");

      if (refStorageBadge) {
        refStorageBadge.textContent = "Im Browser gespeichert";
        refStorageBadge.className = "badge badge-ok";
      }

      await saveRefFileToStorage(file.name, arrayBuffer, state.currentRefSheet);

      // Mitarbeiter automatisch aus "Personalnummer" (Spalte A) & "Name" (Spalte B) extrahieren
      const staffCount = extractStaffFromReferenceWorkbook(wb);

      checkReadyForConfig();

      // Falls Prüfdatei bereits geladen ist: Spaltenauswahl, Matrix & Gesamttabelle sofort aktualisieren
      if (state.targetWorkbook && state.currentTargetSheet) {
        if (typeof populateFullTableFilters === "function") populateFullTableFilters();
        if (typeof renderTimesheetMatrix === "function") renderTimesheetMatrix();
        if (typeof renderFullExcelTable === "function") renderFullExcelTable();
      }

      if (staffCount > 0) {
        showToast(`✅ Referenzdatei "${file.name}" geladen (${staffCount} Mitarbeiter erkannt) & im Browser gespeichert.`);
      } else {
        showToast(`Referenzdatei "${file.name}" geladen und im Browser gespeichert.`);
      }
    } catch (err) {
      console.error(err);
      const isOldXls = file.name.toLowerCase().endsWith(".xls") && !file.name.toLowerCase().endsWith(".xlsx");
      if (isOldXls) {
        alert(`Die Datei "${file.name}" liegt im alten Excel 97-2003-Format (.xls) vor.\n\nBitte öffnen Sie die Datei kurz in Excel und speichern Sie sie über "Datei > Speichern unter" als "Excel-Arbeitsmappe (.xlsx)" ab.`);
      } else {
        alert("Fehler beim Laden der Referenzdatei: " + err.message);
      }
    }
  });

  setupDropZone(dropTgt, inputTgt, async (file) => {
    try {
      state.targetFileName = file.name;
      state.appliedCorrections = {};
      const arrayBuffer = await file.arrayBuffer();
      state.lastTargetBuffer = arrayBuffer;
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      state.targetWorkbook = wb;

      tgtFilename.textContent = file.name;
      populateSheetSelect(selectTgtSheet, wb.worksheets);
      state.currentTargetSheet = selectTgtSheet.value;
      tgtFileInfo.classList.remove("hidden");
      dropTgt.querySelector(".drop-zone-content").classList.add("hidden");

      if (tgtStatusBadge) {
        tgtStatusBadge.textContent = "Bereit zur Prüfung";
        tgtStatusBadge.className = "badge badge-ok";
        tgtStatusBadge.style.display = "inline-block";
      }

      checkReadyForConfig();
      showToast(`Prüfdatei "${file.name}" geladen.`);
    } catch (err) {
      console.error(err);
      const isOldXls = file.name.toLowerCase().endsWith(".xls") && !file.name.toLowerCase().endsWith(".xlsx");
      if (isOldXls) {
        alert(`Die Datei "${file.name}" liegt im alten Excel 97-2003-Format (.xls) vor.\n\nBitte öffnen Sie die Datei kurz in Excel und speichern Sie sie über "Datei > Speichern unter" als "Excel-Arbeitsmappe (.xlsx)" ab.`);
      } else {
        alert("Fehler beim Laden der Prüfdatei: " + err.message);
      }
    }
  });

  // --- Mitarbeiter-Stammdaten: Browser-Speicher & Extraktion aus Referenzdatei ---
  const STORAGE_KEY_STAFF = "excel_pruefer_mitarbeiter_v1";

  function loadStaffFromStorage() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_STAFF);
      if (stored) {
        const list = JSON.parse(stored);
        if (Array.isArray(list) && list.length > 0) {
          applyStaffList(list);
          return true;
        }
      }
    } catch (e) {
      console.warn("Fehler beim Laden der Mitarbeiter aus localStorage:", e);
    }
    updateStaffUI();
    return false;
  }

  // --- Mitarbeiter-Stammdaten: Hilfsfunktionen, Registrierung & Normalisierung ---
  function registerStaffEntry(rawRes, name) {
    if (!rawRes || !name) return;
    const clean = String(rawRes).normalize("NFC").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
    const cleanName = String(name).normalize("NFC").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
    if (!clean || !cleanName) return;

    // 1. Exakter String & Kleinschreibung
    state.staffMap[clean] = cleanName;
    state.staffMap[clean.toLowerCase()] = cleanName;

    // 2. Ohne float-Endung .0 oder ,0 (Excel)
    const noFloat = clean.replace(/[,.]0+$/, "");
    state.staffMap[noFloat] = cleanName;
    state.staffMap[noFloat.toLowerCase()] = cleanName;

    // 3. Wenn es Ziffern sind (z. B. "1773" oder "1373" oder "0045"):
    // Unpadded und alle gängigen Padding-Längen (1- bis 10-stellig) mappen
    if (/^\d+$/.test(noFloat)) {
      const numVal = parseInt(noFloat, 10);
      const unpadded = String(numVal);
      state.staffMap[unpadded] = cleanName;
      for (let len = 1; len <= 10; len++) {
        state.staffMap[unpadded.padStart(len, "0")] = cleanName;
      }
    } else {
      const digitsMatch = noFloat.match(/\b\d+\b/);
      if (digitsMatch) {
        const numVal = parseInt(digitsMatch[0], 10);
        const unpadded = String(numVal);
        state.staffMap[unpadded] = cleanName;
        for (let len = 1; len <= 10; len++) {
          state.staffMap[unpadded.padStart(len, "0")] = cleanName;
        }
      }
    }

    // 4. Auch nach dem Namen indexieren
    state.staffMap[cleanName] = cleanName;
    state.staffMap[cleanName.toLowerCase()] = cleanName;

    // 5. Bei Namen mit Althön / Altöhn oder Schreibweisen-Varianten:
    const lowerName = cleanName.toLowerCase();
    const isAlthoen = lowerName.includes("althön") || lowerName.includes("altöhn") || lowerName.includes("althoen");
    if (isAlthoen) {
      state.staffMap["althön"] = cleanName;
      state.staffMap["altöhn"] = cleanName;
      state.staffMap["althoen"] = cleanName;
      state.staffMap["uwe althön"] = cleanName;
      state.staffMap["uwe altöhn"] = cleanName;
      state.staffMap["althön, uwe"] = cleanName;
      state.staffMap["altöhn, uwe"] = cleanName;
      state.staffMap["althön uwe"] = cleanName;
      state.staffMap["altöhn uwe"] = cleanName;

      // Beide Nummern (1773 und 1373) für Althön registrieren
      ["1773", "1373"].forEach(num => {
        state.staffMap[num] = cleanName;
        for (let len = 1; len <= 10; len++) {
          state.staffMap[num.padStart(len, "0")] = cleanName;
        }
        state.staffMap[`${num} ${lowerName}`] = cleanName;
        state.staffMap[`${num} althön`] = cleanName;
        state.staffMap[`${num} altöhn`] = cleanName;
      });
    }

    // Einzelne Namensbestandteile indexieren
    const nameParts = cleanName.split(/[\s,–-]+/).filter(p => p.length > 2);
    nameParts.forEach(part => {
      const pLow = part.toLowerCase();
      if (!state.staffMap[pLow]) state.staffMap[pLow] = cleanName;
    });
  }

  function getStaffName(rawRes) {
    if (rawRes === null || rawRes === undefined) return null;
    const s = String(rawRes).normalize("NFC").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
    if (!s) return null;

    // 1. Direkter Treffer
    if (state.staffMap[s]) return state.staffMap[s];

    // 2. Ohne .0 oder ,0 Float-Endung
    const noFloat = s.replace(/[,.]0+$/, "");
    if (state.staffMap[noFloat]) return state.staffMap[noFloat];

    // 3. Numerischer Abgleich (unpadded und padded 1 bis 10 Stellen)
    if (/^\d+$/.test(noFloat)) {
      const numVal = parseInt(noFloat, 10);
      const unpadded = String(numVal);
      if (state.staffMap[unpadded]) return state.staffMap[unpadded];
      for (let len = 1; len <= 10; len++) {
        const p = unpadded.padStart(len, "0");
        if (state.staffMap[p]) return state.staffMap[p];
      }
    }

    // 4. Case-insensitive
    if (state.staffMap[s.toLowerCase()]) return state.staffMap[s.toLowerCase()];
    if (state.staffMap[noFloat.toLowerCase()]) return state.staffMap[noFloat.toLowerCase()];

    // 5. Ziffern-Extraktion falls Kombination wie "1773 Althön", "1373 Uwe Altöhn", "Nr. 1773" oder "MA-1773"
    const digitsMatch = noFloat.match(/\b\d+\b/);
    if (digitsMatch) {
      const numVal = parseInt(digitsMatch[0], 10);
      const unpadded = String(numVal);
      if (state.staffMap[unpadded]) return state.staffMap[unpadded];
      for (let len = 1; len <= 10; len++) {
        const p = unpadded.padStart(len, "0");
        if (state.staffMap[p]) return state.staffMap[p];
      }
    }

    // 6. Schreibweisen-Toleranz für Althön / Altöhn
    const sLow = s.toLowerCase();
    if (sLow.includes("althön") || sLow.includes("altöhn") || sLow.includes("althoen")) {
      if (state.staffMap["althön"]) return state.staffMap["althön"];
      if (state.staffMap["altöhn"]) return state.staffMap["altöhn"];
    }

    return null;
  }

  function applyStaffList(list) {
    state.staffList = list;
    state.staffMap = {};
    list.forEach(item => {
      const rawRes = String(item.resource || item.res || "").trim();
      const name = String(item.name || item.mitarbeiter || "").trim();
      if (rawRes && name) {
        registerStaffEntry(rawRes, name);
      }
    });
    updateStaffUI();
  }

  function saveStaffToStorage(list) {
    try {
      localStorage.setItem(STORAGE_KEY_STAFF, JSON.stringify(list));
      applyStaffList(list);
      if (state.timesheetData) {
        renderTimesheetMatrix();
      }
      if (typeof populateFullTableFilters === "function") populateFullTableFilters();
      if (typeof renderFullExcelTable === "function") renderFullExcelTable();
    } catch (e) {
      console.error("Fehler beim Speichern in localStorage:", e);
    }
  }

  function clearStaffStorage() {
    state.staffList = [];
    state.staffMap = {};
    try { localStorage.removeItem(STORAGE_KEY_STAFF); } catch (e) {}
    updateStaffUI();
    if (typeof renderTimesheetMatrix === "function") renderTimesheetMatrix();
    if (typeof populateFullTableFilters === "function") populateFullTableFilters();
    if (typeof renderFullExcelTable === "function") renderFullExcelTable();
    showToast("Mitarbeiter-Stammdaten aus dem Browser gelöscht.");
  }

  function updateStaffUI() {
    const count = state.staffList ? state.staffList.length : 0;
    if (staffFileInfo && staffDropZoneContent) {
      if (count > 0) {
        staffFileInfo.classList.remove("hidden");
        staffDropZoneContent.classList.add("hidden");
        if (staffInfoTitle) staffInfoTitle.textContent = `${count} Mitarbeiter aktiv`;
        if (staffInfoDetail) staffInfoDetail.textContent = `💾 Im Browser gespeichert (dauerhaft erhalten)`;
        if (staffStorageBadge) {
          staffStorageBadge.textContent = "Im Browser gespeichert";
          staffStorageBadge.className = "badge badge-ok";
        }
      } else {
        staffFileInfo.classList.add("hidden");
        staffDropZoneContent.classList.remove("hidden");
        if (staffStorageBadge) {
          staffStorageBadge.textContent = "Optional";
          staffStorageBadge.className = "badge badge-leer-ok";
        }
      }
    }
    if (tsStaffStatusBadge) {
      if (count > 0) {
        tsStaffStatusBadge.className = "badge badge-ok";
        tsStaffStatusBadge.innerHTML = `👤 ${count} Mitarbeiter aktiv`;
      } else {
        tsStaffStatusBadge.className = "badge badge-leer-ok";
        tsStaffStatusBadge.innerHTML = `⚠️ Keine Mitarbeiterdatei geladen (alle Ressourcen werden angezeigt)`;
      }
    }
  }

  function getCleanCellValue(cell) {
    if (!cell) return "";
    let val = WebExcelEngine.extractCellValue(cell.value, cell);
    if (!val && cell.text) val = cell.text;
    if (!val && cell.model && cell.model.value !== undefined) val = String(cell.model.value);
    return String(val || "").normalize("NFC").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
  }

  // --- Mitarbeiter-Datei Drag & Drop Einrichten (Spalte A: Nummer, Spalte B: Name) ---
  if (dropStaff && inputStaff) {
    setupDropZone(dropStaff, inputStaff, async (file) => {
      try {
        showToast(`Lese Mitarbeiterdatei "${file.name}"...`);
        const parsedList = [];
        const seenResKeys = new Set();

        const addStaffItem = (rawRes, rawName, dept = "") => {
          let cleanRes = String(rawRes || "").normalize("NFC").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
          let cleanName = String(rawName || "").normalize("NFC").replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();

          // Spezialfall 1: In einem Feld steht sowohl Nummer als auch Name (z. B. "1773 Althön" oder "1773 - Uwe Altöhn")
          if (cleanRes && !cleanName) {
            const m = cleanRes.match(/^(\d+)\s*[-:–/|]?\s*(.+)$/);
            if (m) {
              cleanRes = m[1];
              cleanName = m[2].trim();
            }
          }
          if (!cleanRes && cleanName) {
            const m = cleanName.match(/^(\d+)\s*[-:–/|]?\s*(.+)$/);
            if (m) {
              cleanRes = m[1];
              cleanName = m[2].trim();
            }
          }

          // Spezialfall 2: Nummer und Name sind vertauscht
          const numOnlyRes = cleanRes.replace(/[,.]0+$/, "");
          const numOnlyName = cleanName.replace(/[,.]0+$/, "");
          if (!/^\d+$/.test(numOnlyRes) && /^\d+$/.test(numOnlyName)) {
            const tmp = cleanRes;
            cleanRes = cleanName;
            cleanName = tmp;
          }

          // Spezialfall 3: Name wiederholt die Nummer ("1773 Althön")
          if (cleanRes && cleanName) {
            const n = cleanRes.replace(/[,.]0+$/, "").replace(/^0+/, "");
            if (cleanName.startsWith(cleanRes)) {
              cleanName = cleanName.substring(cleanRes.length).replace(/^[-:–/|\s]+/, "").trim();
            } else if (n && cleanName.startsWith(n)) {
              cleanName = cleanName.substring(n.length).replace(/^[-:–/|\s]+/, "").trim();
            }
          }

          cleanRes = cleanRes.replace(/[,.]0+$/, "");
          if (!cleanRes) return;

          // Kopfzeilen ignorieren
          const lowerR = cleanRes.toLowerCase();
          const lowerN = cleanName.toLowerCase();
          if (["personalnummer", "persnr", "ressource", "resource", "nummer", "nr"].some(k => lowerR === k) &&
              ["name", "mitarbeiter", "person", "nachname", "bezeichnung"].some(k => lowerN === k)) {
            return;
          }

          if (!cleanName) {
            cleanName = `Mitarbeiter ${cleanRes}`;
          }

          const dKey = cleanRes;
          if (!seenResKeys.has(dKey)) {
            seenResKeys.add(dKey);
            parsedList.push({
              resource: cleanRes,
              name: cleanName,
              dept: dept || ""
            });
          }
        };

        const isCsv = file.name.toLowerCase().endsWith(".csv") || file.name.toLowerCase().endsWith(".txt");
        let loadedViaExcelJS = false;

        if (!isCsv) {
          try {
            const arrayBuffer = await file.arrayBuffer();
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(arrayBuffer);

            wb.worksheets.forEach(ws => {
              if (!ws) return;
              const maxR = Math.max(
                ws.rowCount || 0,
                ws.actualRowCount || 0,
                (ws._rows ? ws._rows.length : 0),
                100
              );

              let emptyCount = 0;
              for (let r = 1; r <= maxR; r++) {
                const row = ws.getRow(r);
                if (!row) {
                  emptyCount++;
                  if (emptyCount > 60 && r > 20) break;
                  continue;
                }

                let cellA = getCleanCellValue(row.getCell(1));
                let cellB = getCleanCellValue(row.getCell(2));
                let cellC = getCleanCellValue(row.getCell(3));
                let cellD = getCleanCellValue(row.getCell(4));

                if (!cellA && !cellB && !cellC && !cellD) {
                  emptyCount++;
                  if (emptyCount > 60 && r > 20) break;
                  continue;
                }
                emptyCount = 0;

                // Fall 1: Spalte A = Nummer, Spalte B = Name
                // Fall 2: Spalte A = Nummer, Spalte B = Vorname, Spalte C = Nachname
                if (/^\d+/.test(cellA) && cellB && cellC && !/^\d+/.test(cellB) && !/^\d+/.test(cellC)) {
                  addStaffItem(cellA, `${cellB} ${cellC}`, ws.name);
                }
                // Fall 3: Spalte A = laufende ID (1,2,3...), Spalte B = Personalnummer (z. B. 1773), Spalte C = Name
                else if (/^\d+$/.test(cellA) && cellA.length <= 2 && /^\d+$/.test(cellB) && cellB.length >= 3 && cellC) {
                  addStaffItem(cellB, cellC, ws.name);
                }
                // Fall 4: Standard Spalte A und B
                else if (cellA || cellB) {
                  addStaffItem(cellA, cellB, ws.name);
                }
              }
            });
            loadedViaExcelJS = true;
          } catch (excelErr) {
            console.warn("ExcelJS konnte Mitarbeiterdatei nicht direkt als XLSX öffnen. Prüfe Text/HTML/CSV-Fallback:", excelErr);
          }
        }

        // Text-/CSV-/HTML-Fallback (falls .csv oder wenn ExcelJS fehlgeschlagen ist, z. B. CSV/TSV oder HTML mit .xls-Endung)
        if (!loadedViaExcelJS || parsedList.length === 0) {
          try {
            const text = await file.text();
            if (text.includes("<table") || text.includes("<tr")) {
              const parser = new DOMParser();
              const doc = parser.parseFromString(text, "text/html");
              const rows = doc.querySelectorAll("tr");
              rows.forEach(tr => {
                const cells = Array.from(tr.querySelectorAll("td, th")).map(c => c.textContent.trim());
                if (cells.length >= 2) {
                  addStaffItem(cells[0], cells[1]);
                } else if (cells.length === 1) {
                  addStaffItem(cells[0], "");
                }
              });
            } else {
              const lines = text.split(/\r?\n/);
              for (let l = 0; l < lines.length; l++) {
                const line = lines[l].trim();
                if (!line) continue;
                let parts = line.split(";");
                if (parts.length < 2) parts = line.split("\t");
                if (parts.length < 2) parts = line.split(",");
                if (parts.length < 2) parts = line.split("|");
                if (parts.length >= 2) {
                  let pA = parts[0];
                  let pB = parts[1];
                  let pC = parts.length > 2 ? parts[2] : "";
                  if (pC && !/^\d+$/.test(pC.trim()) && !/^\d+$/.test(pB.trim())) {
                    pB = `${pB.trim()} ${pC.trim()}`;
                  }
                  addStaffItem(pA, pB, "");
                } else {
                  addStaffItem(parts[0], "", "");
                }
              }
            }
          } catch (textErr) {
            console.warn("Text-Fallback nicht erfolgreich:", textErr);
          }
        }

        if (parsedList.length === 0) {
          const isOldXls = file.name.toLowerCase().endsWith(".xls") && !file.name.toLowerCase().endsWith(".xlsx");
          if (isOldXls) {
            alert(`Die Datei "${file.name}" liegt im alten Excel 97-2003-Format (.xls) vor.\n\nBitte öffnen Sie die Datei kurz in Excel und speichern Sie sie über "Datei > Speichern unter" als modernes "Excel-Arbeitsmappe (.xlsx)" oder als ".csv" ab. Danach kann sie sofort eingelesen werden.`);
          } else {
            alert("In der Mitarbeiterdatei wurden keine Mitarbeiter erkannt. Bitte stellen Sie sicher, dass in Spalte A die Nummern und in Spalte B die Namen stehen.");
          }
          return;
        }

        saveStaffToStorage(parsedList);
        showToast(`✅ ${parsedList.length} Mitarbeiter geladen und dauerhaft im Browser gespeichert!`);

        // Falls Ziel-Datei bereits geladen ist: Spaltenauswahl & Arbeitszeiten sofort aktualisieren
        if (state.targetWorkbook && state.currentTargetSheet) {
          const wsTarget = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
          if (wsTarget) {
            const tgtHeaders = [];
            const headerRow = wsTarget.getRow(1);
            let colCount = wsTarget.columnCount;
            headerRow.eachCell((cell, colNum) => { if (colNum > colCount) colCount = colNum; });
            for (let c = 1; c <= colCount; c++) {
              const hCell = headerRow.getCell(c);
              const val = WebExcelEngine.extractCellValue(hCell.value, hCell).trim();
              tgtHeaders.push({
                colNum: c,
                letter: getColLetter(c),
                name: val || `Spalte ${getColLetter(c)}`
              });
            }
            populateTimesheetColSelects(tgtHeaders);
          }
          if (typeof renderTimesheetMatrix === "function") renderTimesheetMatrix();
          if (typeof populateFullTableFilters === "function") populateFullTableFilters();
          if (typeof renderFullExcelTable === "function") renderFullExcelTable();
        }
      } catch (err) {
        console.error(err);
        alert("Fehler beim Lesen der Mitarbeiterdatei: " + err.message);
      }
    });
  }

  if (btnChangeStaff && inputStaff) {
    btnChangeStaff.addEventListener("click", (e) => {
      e.stopPropagation();
      inputStaff.click();
    });
  }

  if (btnClearStaff) {
    btnClearStaff.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!confirm("Möchten Sie die gespeicherten Mitarbeiter wirklich aus dem Browser löschen?")) return;
      clearStaffStorage();
    });
  }

  if (btnDownloadStaffTemplate) {
    btnDownloadStaffTemplate.addEventListener("click", async (e) => {
      e.stopPropagation();
      try {
        const buf = await generateStaffTemplateExcel();
        downloadBuffer(buf, "Mitarbeiter_Stammdaten_Vorlage.xlsx");
        showToast("📥 Vorlage für Mitarbeiter-Stammdaten heruntergeladen.");
      } catch (err) {
        console.error(err);
        alert("Fehler beim Erstellen der Vorlage: " + err.message);
      }
    });
  }

  // --- Mitarbeiter-Extraktion als optionale Ergänzung aus der Referenzdatei ---
  function extractStaffFromReferenceWorkbook(wb) {
    if (!wb || !wb.worksheets || wb.worksheets.length === 0) return 0;
    const staffList = [...(state.staffList || [])];
    const seenMap = new Map();
    staffList.forEach(s => seenMap.set(s.resource, s.name));

    let addedCount = 0;

    wb.worksheets.forEach(ws => {
      if (!ws) return;
      const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length : 0));
      if (maxRows === 0) return;

      let foundResCol = null;
      let foundNameCol = null;
      let startRow = 2;

      for (let r = 1; r <= Math.min(maxRows, 5); r++) {
        const row = ws.getRow(r);
        if (!row) continue;
        let cRes = null, cName = null;
        row.eachCell((cell, colNum) => {
          const val = getCleanCellValue(cell).toLowerCase();
          if (!cRes && ["personalnummer", "persnr", "pers-nr", "ressource", "resource", "personal-nr", "mitarbeiternr", "mitarbeiter-nr", "mitarbeiter_nr"].some(k => val.includes(k))) {
            cRes = colNum;
          }
          if (!cName && ["name", "mitarbeiter", "person", "nachname", "bezeichnung"].some(k => val.includes(k))) {
            cName = colNum;
          }
        });
        if (cRes && cName) {
          foundResCol = cRes;
          foundNameCol = cName;
          startRow = r + 1;
          break;
        }
      }

      if (!foundResCol || !foundNameCol) return;

      for (let r = startRow; r <= maxRows; r++) {
        const row = ws.getRow(r);
        if (!row) continue;
        let resVal = getCleanCellValue(row.getCell(foundResCol));
        let nameVal = getCleanCellValue(row.getCell(foundNameCol));
        if (!resVal) continue;
        resVal = resVal.replace(/[,.]0+$/, "");
        if (!nameVal) nameVal = `Mitarbeiter ${resVal}`;

        if (!seenMap.has(resVal)) {
          seenMap.set(resVal, nameVal);
          staffList.push({
            resource: resVal,
            name: nameVal,
            dept: ws.name
          });
          addedCount++;
        }
      }
    });

    if (addedCount > 0) {
      applyStaffList(staffList);
      try {
        localStorage.setItem(STORAGE_KEY_STAFF, JSON.stringify(staffList));
      } catch (e) {}
      console.log(`✅ ${addedCount} zusätzliche Mitarbeiter aus Referenzdatei extrahiert.`);
    }

    return addedCount;
  }

  if (btnChangeRef && inputRef) {
    btnChangeRef.addEventListener("click", (e) => {
      e.stopPropagation();
      inputRef.click();
    });
  }

  if (btnClearRef) {
    btnClearRef.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!confirm("Möchten Sie die gespeicherte Referenzdatei wirklich aus dem Browser löschen?")) return;
      await deleteRefFileFromStorage();
      state.refWorkbook = null;
      state.refFileName = "";
      state.currentRefSheet = "";
      if (refFilename) refFilename.textContent = "";
      if (selectRefSheet) selectRefSheet.innerHTML = "";
      if (refFileInfo) refFileInfo.classList.add("hidden");
      const dropContent = dropRef ? dropRef.querySelector(".drop-zone-content") : null;
      if (dropContent) dropContent.classList.remove("hidden");
      if (refStorageBadge) {
        refStorageBadge.textContent = "Nicht geladen";
        refStorageBadge.className = "badge badge-secondary";
      }
      if (configSection) configSection.classList.add("hidden");
      showToast("Referenzdatei aus dem Browser gelöscht.");
    });
  }

  const btnDownloadRef = document.getElementById("btn-download-ref");
  if (btnDownloadRef) {
    btnDownloadRef.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!state.refWorkbook) {
        showToast("Keine Referenzdatei geladen.");
        return;
      }
      try {
        const buf = await state.refWorkbook.xlsx.writeBuffer();
        const baseName = (state.refFileName || "Referenz_Stammdaten").replace(/\.[^/.]+$/, "");
        downloadBuffer(buf, `${baseName}_AKTUALISIERT.xlsx`);
        showToast("📥 Aktualisierte Referenzdatei wird heruntergeladen...");
      } catch (err) {
        console.error("Fehler beim Herunterladen der Referenzdatei:", err);
        alert("Fehler beim Herunterladen der Referenzdatei: " + err.message);
      }
    });
  }

  if (btnChangeTgt && inputTgt) {
    btnChangeTgt.addEventListener("click", (e) => {
      e.stopPropagation();
      inputTgt.click();
    });
  }

  // Gespeicherte Referenzdatei & Mitarbeiter beim Start aus Browser-Speicher laden
  restoreSavedReferenceFile();
  loadStaffFromStorage();

  function setupDropZone(dropZone, fileInput, onFileLoaded) {
    if (!dropZone || !fileInput) return;

    let dragCounter = 0;

    dropZone.addEventListener("click", (e) => {
      // Nicht auslösen, falls ein Button, Select, Label, Link oder das Info-Feld angeklickt wurde
      if (
        e.target.closest("button") ||
        e.target.closest("select") ||
        e.target.closest("option") ||
        e.target.closest("label") ||
        e.target.closest("a") ||
        e.target.closest(".file-info")
      ) {
        return;
      }
      fileInput.click();
    });

    fileInput.addEventListener("change", () => {
      if (fileInput.files && fileInput.files.length > 0) {
        onFileLoaded(fileInput.files[0]);
        fileInput.value = "";
      }
    });

    dropZone.addEventListener("dragenter", (e) => {
      e.preventDefault();
      e.stopPropagation();
      dragCounter++;
      if (e.dataTransfer) {
        e.dataTransfer.dropEffect = "copy";
      }
      dropZone.classList.add("dragover");
    }, false);

    dropZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (e.dataTransfer) {
        e.dataTransfer.dropEffect = "copy";
      }
      dropZone.classList.add("dragover");
    }, false);

    dropZone.addEventListener("dragleave", (e) => {
      e.preventDefault();
      e.stopPropagation();
      dragCounter--;
      if (dragCounter <= 0) {
        dragCounter = 0;
        dropZone.classList.remove("dragover");
      }
    }, false);

    dropZone.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      dragCounter = 0;
      dropZone.classList.remove("dragover");
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        onFileLoaded(e.dataTransfer.files[0]);
      }
    }, false);
  }

  function populateSheetSelect(selectEl, worksheets) {
    selectEl.innerHTML = "";
    worksheets.forEach(ws => {
      const opt = document.createElement("option");
      opt.value = ws.name;
      opt.textContent = `${ws.name} (${ws.rowCount} Zeilen)`;
      selectEl.appendChild(opt);
    });
  }

  selectRefSheet.addEventListener("change", async (e) => {
    state.currentRefSheet = e.target.value;
    await updateRefSheetInStorage(e.target.value);
    extractStaffFromReferenceWorkbook(state.refWorkbook);
    updateMapping();
    if (state.targetWorkbook && state.currentTargetSheet) {
      renderTimesheetMatrix();
      populateFullTableFilters();
      renderFullExcelTable();
    }
  });

  selectTgtSheet.addEventListener("change", (e) => {
    state.currentTargetSheet = e.target.value;
    updateMapping();
  });

  function checkReadyForConfig() {
    if (state.refWorkbook && state.targetWorkbook) {
      updateMapping();
      configSection.classList.remove("hidden");
      configSection.scrollIntoView({ behavior: "smooth" });
    }
  }

  // --- Spalten-Mapping ermitteln ---
  function updateMapping() {
    if (!state.refWorkbook || !state.targetWorkbook) return;
    const refWs = state.refWorkbook.getWorksheet(state.currentRefSheet);
    const tgtWs = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!refWs || !tgtWs) return;

    const refHeaders = [];
    const refHeaderRow = refWs.getRow(1);
    refHeaderRow.eachCell((cell, colNum) => {
      const extracted = WebExcelEngine.extractCellValue(cell.value, cell);
      const val = extracted ? extracted.trim() : `Spalte ${getColLetter(colNum)}`;
      refHeaders.push({ colNum, name: val });
    });

    const tgtHeaders = [];
    const tgtHeaderRow = tgtWs.getRow(1);
    tgtHeaderRow.eachCell((cell, colNum) => {
      const extracted = WebExcelEngine.extractCellValue(cell.value, cell);
      const val = extracted ? extracted.trim() : `Spalte ${getColLetter(colNum)}`;
      tgtHeaders.push({ colNum, letter: getColLetter(colNum), name: val });
    });

    const mapping = [];
    tgtHeaders.forEach(tgt => {
      const tgtClean = tgt.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
      let bestMatch = "";

      for (const ref of refHeaders) {
        const refClean = ref.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
        if (tgtClean === refClean && tgtClean) {
          bestMatch = ref.name;
          break;
        }
      }
      if (!bestMatch) {
        for (const ref of refHeaders) {
          const refClean = ref.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
          if (tgtClean && refClean && (tgtClean.includes(refClean) || refClean.includes(tgtClean))) {
            bestMatch = ref.name;
            break;
          }
        }
      }

      // Spalten B, C und E sind standardmäßig angehakt, alle anderen abwählbar / optional
      const isDefaultCol = ["B", "C", "E"].includes(tgt.letter.toUpperCase()) || [2, 3, 5].includes(tgt.colNum);
      const isRessource = (tgt.letter.toUpperCase() === "C" || tgt.colNum === 3);
      const isLeistung = (tgt.letter.toUpperCase() === "E") || (tgt.colNum === 5) || ["leistung", "leist", "leist-nr", "leist_nr", "leistnr"].some(k => tgt.name.toLowerCase().includes(k));

      // Spezieller Match für Ressource <-> Personalnummer in Referenzdatei
      if (!bestMatch && isRessource) {
        for (const ref of refHeaders) {
          const refClean = ref.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
          if (["personalnummer", "persnr", "personal", "ressourcennummer", "ressource", "mitarbeiternummer", "mitarbeiternr"].some(k => refClean.includes(k))) {
            bestMatch = ref.name;
            break;
          }
        }
      }

      // Falls kein Match über Header-Namen gefunden wurde, nach gleicher Position/Buchstabe suchen
      if (!bestMatch) {
        for (const ref of refHeaders) {
          if (ref.colNum === tgt.colNum || (ref.letter && ref.letter.toUpperCase() === tgt.letter.toUpperCase())) {
            bestMatch = ref.name;
            break;
          }
        }
      }

      mapping.push({
        target_col_idx: tgt.colNum,
        target_col_letter: tgt.letter,
        target_col_name: tgt.name,
        ref_col_name: bestMatch,
        selected: isDefaultCol,
        allow_empty: isLeistung,
        pad_to_4: isRessource
      });
    });

    state.mapping = mapping;
    renderMappingTable(mapping, refHeaders.map(h => h.name));
    populateTimesheetColSelects(tgtHeaders);
    populateFullTableFilters();
    renderFullExcelTable();
  }

  function renderMappingTable(mapping, refHeaderNames) {
    mappingTbody.innerHTML = mapping.map((m, idx) => {
      const isChecked = m.selected ? "checked" : "";
      let optionsHtml = `<option value="">-- Nicht prüfen --</option>`;
      refHeaderNames.forEach(h => {
        const isSel = (h === m.ref_col_name) ? "selected" : "";
        optionsHtml += `<option value="${escapeHtml(h)}" ${isSel}>${escapeHtml(h)}</option>`;
      });

      const isAllowEmptyChecked = Boolean(m.allow_empty) ? "checked" : "";

      return `
        <tr>
          <td>
            <input type="checkbox" class="col-checkbox" data-idx="${idx}" ${isChecked}>
          </td>
          <td>
            <strong>${escapeHtml(m.target_col_name)}</strong>
            <span class="diag-detail"> (Spalte ${m.target_col_letter})</span>
          </td>
          <td style="color: var(--text-muted); font-size: 1.1rem;">➔</td>
          <td>
            <select class="form-select ref-col-select" data-idx="${idx}">
              ${optionsHtml}
            </select>
          </td>
          <td>
            <label class="custom-control">
              <input type="checkbox" class="allow-empty-checkbox" data-idx="${idx}" ${isAllowEmptyChecked}>
              <span>Leere Zellen erlauben</span>
            </label>
          </td>
        </tr>
      `;
    }).join("");

    // Synchronisiere "Alle auswählen"-Checkbox
    checkAllCols.checked = mapping.length > 0 && mapping.every(m => m.selected);

    // Event Listener für Mapping-Tabelle
    mappingTbody.querySelectorAll(".col-checkbox").forEach(cb => {
      cb.addEventListener("change", (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        state.mapping[idx].selected = e.target.checked;
        checkAllCols.checked = state.mapping.length > 0 && state.mapping.every(m => m.selected);
      });
    });

    mappingTbody.querySelectorAll(".ref-col-select").forEach(sel => {
      sel.addEventListener("change", (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        state.mapping[idx].ref_col_name = e.target.value;
        state.mapping[idx].selected = Boolean(e.target.value);
        const rowCb = mappingTbody.querySelector(`.col-checkbox[data-idx="${idx}"]`);
        if (rowCb) rowCb.checked = state.mapping[idx].selected;
        checkAllCols.checked = state.mapping.length > 0 && state.mapping.every(m => m.selected);
      });
    });

    mappingTbody.querySelectorAll(".allow-empty-checkbox").forEach(cb => {
      cb.addEventListener("change", (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        state.mapping[idx].allow_empty = e.target.checked;
      });
    });
  }

  checkAllCols.addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    state.mapping.forEach(m => { m.selected = isChecked; });
    mappingTbody.querySelectorAll(".col-checkbox").forEach(cb => { cb.checked = isChecked; });
  });

  // --- Datumsprüfung Konfiguration & Status-Badge ---
  const chkDateVal = document.getElementById("chk-date-validation");
  const badgeDateStatus = document.getElementById("badge-date-status");
  const inputDateWeeks = document.getElementById("input-date-weeks");
  const dateControls = document.getElementById("date-check-controls");

  function updateDateBadge() {
    if (!chkDateVal) return;
    if (!chkDateVal.checked) {
      if (badgeDateStatus) {
        badgeDateStatus.textContent = "Deaktiviert";
        badgeDateStatus.className = "badge badge-leer-ok";
      }
      if (dateControls) dateControls.style.opacity = "0.5";
    } else {
      const mode = document.querySelector('input[name="date-mode"]:checked')?.value || "weeks";
      const w = inputDateWeeks ? inputDateWeeks.value : 3;
      if (badgeDateStatus) {
        badgeDateStatus.textContent = mode === "weeks" ? `Aktiv (Letzte ${w} Wochen)` : "Aktiv (Fester Zeitraum)";
        badgeDateStatus.className = "badge badge-ok";
      }
      if (dateControls) dateControls.style.opacity = "1";
    }
  }

  if (chkDateVal) {
    chkDateVal.addEventListener("change", updateDateBadge);
  }
  if (inputDateWeeks) {
    inputDateWeeks.addEventListener("input", updateDateBadge);
  }
  document.querySelectorAll('input[name="date-mode"]').forEach(r => {
    r.addEventListener("change", updateDateBadge);
  });

  // --- Prüfung starten ---
  btnStartCheck.addEventListener("click", () => {
    const activeCols = state.mapping.filter(m => m.selected && m.ref_col_name);
    const dateCheckOn = chkDateVal ? chkDateVal.checked : true;
    if (activeCols.length === 0 && !dateCheckOn) {
      alert("Bitte wählen Sie mindestens eine Spalte für den Abgleich aus.");
      return;
    }

    btnStartCheck.disabled = true;
    btnStartCheck.innerHTML = `<span class="icon">⏳</span> Prüfung läuft...`;

    setTimeout(() => {
      runInspection();
      btnStartCheck.disabled = false;
      btnStartCheck.innerHTML = `<span class="icon">🚀</span> Prüfung starten`;
    }, 50);
  });

  function runInspection(options = {}) {
    const refWs = state.refWorkbook.getWorksheet(state.currentRefSheet);
    const tgtWs = state.targetWorkbook.getWorksheet(state.currentTargetSheet);

    // Datumsoptionen auslesen
    const dateCheckEnabled = chkDateVal ? chkDateVal.checked : true;
    const dateRangeMode = document.querySelector('input[name="date-mode"]:checked')?.value || "weeks";
    const weeksVal = parseInt(document.getElementById("input-date-weeks")?.value, 10) || 3;
    const startVal = document.getElementById("input-date-start")?.value || "";
    const endVal = document.getElementById("input-date-end")?.value || "";

    const today = new Date();
    today.setHours(23, 59, 59, 999);
    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);

    let minDate = null;
    let maxDate = today;
    let weeksInfo = `letzte ${weeksVal} Wochen`;

    if (dateRangeMode === "custom" && (startVal || endVal)) {
      minDate = startVal ? WebExcelEngine.parseDateValue(startVal) : new Date(Date.now() - 21 * 86400000);
      maxDate = endVal ? WebExcelEngine.parseDateValue(endVal) : today;
      if (minDate) minDate.setHours(0, 0, 0, 0);
      if (maxDate) maxDate.setHours(23, 59, 59, 999);
      const sDisp = minDate ? WebExcelEngine.formatDate(minDate) : "?";
      const eDisp = maxDate ? WebExcelEngine.formatDate(maxDate) : "?";
      weeksInfo = `Zeitraum von ${sDisp} bis ${eDisp}`;
    } else {
      minDate = new Date(Date.now() - weeksVal * 7 * 86400000);
      minDate.setHours(0, 0, 0, 0);
      weeksInfo = `letzte ${weeksVal} Wochen (ab ${WebExcelEngine.formatDate(minDate)})`;
    }

    // 1. Referenz-Indizes aufbauen
    const colIndices = {};
    state.mapping.forEach(m => {
      if (!m.selected || !m.ref_col_name) return;

      // Finde Spalten-Index in Referenz
      let refColIdx = null;
      refWs.getRow(1).eachCell((cell, colNum) => {
        const hVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
        if (hVal === m.ref_col_name) {
          refColIdx = colNum;
        }
      });

      if (refColIdx) {
        const vals = [];
        for (let r = 2; r <= refWs.rowCount; r++) {
          const cell = refWs.getRow(r).getCell(refColIdx);
          const norm = WebExcelEngine.normalizeCellValue(cell.value, cell);
          if (norm) {
            vals.push(norm);
            if (m.pad_to_4 || /^\d{1,4}$/.test(norm)) {
              vals.push(WebExcelEngine.padNumber(norm, 4));
              const unpadded = norm.replace(/^0+/, "");
              if (unpadded) vals.push(unpadded);
            }
          }
        }
        colIndices[m.target_col_name] = new WebExcelEngine.ColumnIndex(vals);
      }
    });

    // 2. Prüfdatei analysieren
    const results = [];
    const stats = {
      total_cells: 0,
      ok: 0,
      zahlendreher: 0,
      zifferntausch: 0,
      tippfehler: 0,
      ziffer_zuviel_zuwenig: 0,
      nicht_existent: 0,
      leer: 0,
      datum_warnung: 0,
      total_errors: 0
    };

    state.appliedCorrections = state.appliedCorrections || {};

    const colAHeader = tgtWs.getRow(1).getCell(1);
    const colAName = WebExcelEngine.extractCellValue(colAHeader.value, colAHeader).trim() || "Datum";

    for (let r = 2; r <= tgtWs.rowCount; r++) {
      const row = tgtWs.getRow(r);

      // Datumsprüfung Spalte A
      if (dateCheckEnabled) {
        const cellCoordA = `A${r}`;
        const cellKeyA = `${r}_1`;
        const cellA = row.getCell(1);
        const parsedD = WebExcelEngine.parseDateValue(cellA.value, cellA);
        const rawDateStr = WebExcelEngine.extractCellValue(cellA.value, cellA).trim();
        const dispValA = parsedD ? WebExcelEngine.formatDate(parsedD) : (rawDateStr || "");
        stats.total_cells++;

        const todayStr = WebExcelEngine.formatDate(new Date());

        if (cellA.value === null || cellA.value === undefined || rawDateStr === "") {
          stats.total_errors++;
          stats.leer++;
          stats.datum_warnung++;
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: "",
            current_value: "",
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.0,
            suggestion: todayStr,
            detail: "Datum in Spalte A fehlt komplett",
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else if (!parsedD) {
          stats.total_errors++;
          stats.tippfehler++;
          stats.datum_warnung++;
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.0,
            suggestion: todayStr,
            detail: `Ungültiges Datumsformat oder Zahl ('${dispValA}')`,
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else if (minDate && parsedD < minDate) {
          stats.total_errors++;
          stats.tippfehler++;
          stats.datum_warnung++;
          const daysDiff = Math.round((startOfToday - parsedD) / (86400000));
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.5,
            suggestion: todayStr,
            detail: `Datum ${dispValA} liegt ${daysDiff} Tage zurück (erwartet: ${weeksInfo})`,
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else if (maxDate && parsedD > maxDate) {
          stats.total_errors++;
          stats.tippfehler++;
          stats.datum_warnung++;
          const daysDiff = Math.round((parsedD - today) / (86400000));
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "DATUM_WARNUNG",
            status_label: "Datum eventuell falsch",
            badge_class: "badge-tippfehler",
            score: 0.5,
            suggestion: todayStr,
            detail: `Datum ${dispValA} liegt in der Zukunft (+${daysDiff} Tage)`,
            candidates: [{ value: todayStr, detail: "Heutiges Datum" }],
            is_corrected: false
          });
        } else {
          stats.ok++;
          results.push({
            id: cellKeyA,
            row: r,
            col_idx: 1,
            col_letter: "A",
            col_name: colAName,
            ref_col_name: "",
            cell: cellCoordA,
            original_value: dispValA,
            current_value: dispValA,
            status: "OK",
            status_label: "Gültig",
            badge_class: "badge-ok",
            score: 1.0,
            suggestion: dispValA,
            detail: `Aktuelles Datum (${dispValA})`,
            candidates: [],
            is_corrected: false
          });
        }
      }

      state.mapping.forEach(m => {
        if (!m.selected || !m.ref_col_name) return;
        // Wenn Datumsprüfung aktiv ist, Spalte A nicht nochmals gegen Referenzpool abgleichen
        if (dateCheckEnabled && (m.target_col_idx === 1 || m.target_col_letter === "A")) return;

        const index = colIndices[m.target_col_name];
        if (!index) return;

        const cellCoord = `${m.target_col_letter}${r}`;
        const cellKey = `${r}_${m.target_col_idx}`;
        const cell = row.getCell(m.target_col_idx);
        const origRaw = WebExcelEngine.extractCellValue(cell.value, cell);
        const valStr = WebExcelEngine.normalizeCellValue(cell.value, cell);
        const origText = (origRaw && !origRaw.toLowerCase().includes("object")) ? origRaw.trim() : (valStr || "");
        stats.total_cells++;

        const allowEmpty = Boolean(m.allow_empty);

        if (!valStr) {
          if (allowEmpty) {
            stats.ok++;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: "",
              status: "OK_LEER",
              status_label: "Leer (erlaubt)",
              badge_class: "badge-leer-ok",
              score: 1.0,
              suggestion: "",
              detail: "Feld ist leer (optional, kein Fehler)",
              candidates: [],
              is_corrected: false
            });
          } else {
            stats.leer++;
            stats.total_errors++;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: "",
              status: "LEER",
              status_label: "Pflichtfeld leer",
              badge_class: "badge-leer",
              score: 0.0,
              suggestion: "",
              detail: "Zelle ist leer, obwohl Pflichtfeld",
              candidates: [],
              is_corrected: false
            });
          }
          return;
        }

        const targetToCheck = m.pad_to_4 ? WebExcelEngine.padNumber(valStr, 4) : valStr;
        const unpaddedTarget = valStr.replace(/^0+/, "");
        const neededPadding = (targetToCheck !== valStr);
        const isOk = index.contains(targetToCheck) || index.contains(valStr) || (unpaddedTarget && index.contains(unpaddedTarget));

        if (isOk) {
          stats.ok++;
          if (m.pad_to_4 && neededPadding) {
            // Automatisch für den 4-stelligen Export mit 0000 vormerken, aber NICHT als Fehler anzeigen:
            state.appliedCorrections[cellKey] = targetToCheck;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: targetToCheck,
              status: "OK",
              status_label: "Gültig",
              badge_class: "badge-ok",
              score: 1.0,
              suggestion: targetToCheck,
              detail: `Gültige Ressource (wird beim Speichern automatisch 4-stellig als '${targetToCheck}' formatiert)`,
              candidates: [],
              is_corrected: false
            });
          } else {
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: valStr,
              status: "OK",
              status_label: "Gültig",
              badge_class: "badge-ok",
              score: 1.0,
              suggestion: valStr,
              detail: "Exakte Übereinstimmung mit Referenz",
              candidates: [],
              is_corrected: false
            });
          }
        } else {
          stats.total_errors++;
          const candidates = index.findCandidates(targetToCheck);

          if (candidates.length > 0) {
            const best = candidates[0];
            let lbl = "Tippfehler";
            let bclass = "badge-tippfehler";

            if (best.type === "ZAHLENDREHER") {
              stats.zahlendreher++;
              lbl = "Zahlendreher";
              bclass = "badge-zahlendreher";
            } else if (best.type === "ZIFFERENTAUSCH") {
              stats.zifferntausch++;
              lbl = "Zifferntausch";
              bclass = "badge-zifferntausch";
            } else if (best.type === "TIPPFEHLER") {
              stats.tippfehler++;
              lbl = "Tippfehler (1 Ziffer)";
              bclass = "badge-tippfehler";
            } else if (best.type === "ZIFFER_ZUVIEL" || best.type === "ZIFFER_FEHLT") {
              stats.ziffer_zuviel_zuwenig++;
              lbl = "Ziffer zuviel/fehlt";
              bclass = "badge-tippfehler";
            } else if (best.type === "NAECHSTE_ALTERNATIVE") {
              stats.nicht_existent++;
              lbl = "Falscher Wert (Alternative)";
              bclass = "badge-fehler";
            }

            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: valStr,
              status: best.type,
              status_label: lbl,
              badge_class: bclass,
              score: best.score,
              suggestion: best.value,
              detail: best.detail,
              candidates: candidates,
              is_corrected: false
            });
          } else {
            stats.nicht_existent++;
            results.push({
              id: cellKey,
              row: r,
              col_idx: m.target_col_idx,
              col_letter: m.target_col_letter,
              col_name: m.target_col_name,
              ref_col_name: m.ref_col_name,
              cell: cellCoord,
              original_value: origText,
              current_value: valStr,
              status: "NICHT_EXISTENT",
              status_label: "Nicht in Stammdaten",
              badge_class: "badge-fehler",
              score: 0.0,
              suggestion: "",
              detail: "Nummer existiert nicht in der Referenzspalte",
              candidates: [],
              is_corrected: false
            });
          }
        }
      });
    }

    state.allResults = results;
    updateStats(stats);
    updateFilterCounts(results);
    renderResultsTable();
    renderTimesheetMatrix();
    populateFullTableFilters();
    renderFullExcelTable();

    resultsSection.classList.remove("hidden");
    if (!options.skipScroll) {
      resultsSection.scrollIntoView({ behavior: "smooth" });
    }
    if (!options.skipToast) {
      showToast(`Prüfung abgeschlossen: ${stats.total_errors} Abweichungen gefunden.`);
    }
  }

  // --- Statistiken aktualisieren ---
  function updateStats(stats) {
    document.getElementById("stat-total").textContent = stats.total_cells.toLocaleString("de-DE");
    document.getElementById("stat-ok").textContent = stats.ok.toLocaleString("de-DE");
    document.getElementById("stat-zahlendreher").textContent = (stats.zahlendreher + stats.zifferntausch).toLocaleString("de-DE");
    document.getElementById("stat-tippfehler").textContent = (stats.tippfehler + stats.ziffer_zuviel_zuwenig).toLocaleString("de-DE");
    document.getElementById("stat-fehler").textContent = (stats.nicht_existent + stats.leer).toLocaleString("de-DE");
  }

  function updateFilterCounts(results) {
    const isError = (r) => (r.status !== "OK" && r.status !== "OK_LEER");
    const total = results.length;
    const errors = results.filter(isError).length;
    const zd = results.filter(r => r.status === "ZAHLENDREHER" || r.status === "ZIFFERENTAUSCH").length;
    const tp = results.filter(r => r.status === "TIPPFEHLER" || r.status.startsWith("ZIFFER_")).length;
    const ne = results.filter(r => r.status === "NICHT_EXISTENT" || r.status === "LEER" || r.status === "NAECHSTE_ALTERNATIVE").length;
    const datum = results.filter(r => r.status === "DATUM_WARNUNG").length;
    const corr = results.filter(r => r.is_corrected).length;

    document.getElementById("count-all").textContent = total;
    document.getElementById("count-errors").textContent = errors;
    document.getElementById("count-zd").textContent = zd;
    document.getElementById("count-tp").textContent = tp;
    document.getElementById("count-ne").textContent = ne;
    const countDatumEl = document.getElementById("count-datum");
    if (countDatumEl) countDatumEl.textContent = datum;
    document.getElementById("count-corr").textContent = corr;

    if (tabBadgeErrors) {
      tabBadgeErrors.textContent = `${errors} Abweichungen`;
    }
    if (tabBadgeTotalRows && state.targetWorkbook && state.currentTargetSheet) {
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      const totalDataRows = ws ? Math.max(0, ws.rowCount - 1) : 0;
      tabBadgeTotalRows.textContent = `${totalDataRows} Zeilen`;
    }
  }

  // --- Filter Event Listener ---
  filterPills.querySelectorAll(".pill").forEach(p => {
    p.addEventListener("click", () => {
      filterPills.querySelectorAll(".pill").forEach(el => el.classList.remove("active"));
      p.classList.add("active");
      state.currentFilter = p.dataset.filter;
      renderResultsTable();
    });
  });

  tableSearchInput.addEventListener("input", (e) => {
    state.searchQuery = e.target.value.toLowerCase().trim();
    renderResultsTable();
  });

  // --- Referenzdatei: Wert-Prüfung & Dynamisches Hinzufügen neuer Nummern ---
  function isValueInReferenceColumn(refColName, val) {
    if (!state.refWorkbook || !val) return false;
    const clean = String(val).trim();
    if (!clean) return false;

    let targetWs = null;
    let targetColIdx = null;

    if (state.currentRefSheet) {
      targetWs = state.refWorkbook.getWorksheet(state.currentRefSheet);
    }
    if (!targetWs && state.refWorkbook.worksheets.length > 0) {
      targetWs = state.refWorkbook.worksheets[0];
    }

    if (targetWs) {
      targetWs.getRow(1).eachCell((cell, colNum) => {
        const hVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
        if (refColName && hVal.toLowerCase() === refColName.toLowerCase()) {
          targetColIdx = colNum;
        }
      });
    }

    if (!targetColIdx) {
      for (const ws of state.refWorkbook.worksheets) {
        ws.getRow(1).eachCell((cell, colNum) => {
          const hVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
          if (refColName && hVal.toLowerCase() === refColName.toLowerCase()) {
            targetWs = ws;
            targetColIdx = colNum;
          }
        });
        if (targetColIdx) break;
      }
    }

    if (!targetWs || !targetColIdx) return false;

    const maxRows = Math.max(targetWs.rowCount || 0, targetWs.actualRowCount || 0);
    const isNum = /^\d+$/.test(clean);
    const numInt = isNum ? parseInt(clean, 10) : null;

    for (let r = 2; r <= maxRows; r++) {
      const cell = targetWs.getRow(r).getCell(targetColIdx);
      const cv = WebExcelEngine.extractCellValue(cell.value, cell).trim();
      if (!cv) continue;
      if (cv === clean) return true;
      if (isNum && /^\d+$/.test(cv) && parseInt(cv, 10) === numInt) return true;
    }
    return false;
  }

  async function addNumberToReferenceFile(refColName, newNumber, optStaffName = null) {
    try {
      if (!state.refWorkbook) {
        alert("Keine Referenzdatei geladen. Bitte laden Sie zuerst eine Referenzdatei in Schritt 1.");
        return false;
      }

      const cleanNum = String(newNumber || "").trim();
      if (!cleanNum) {
        alert("Bitte geben Sie eine gültige Nummer ein.");
        return false;
      }

      // 1. Arbeitsblatt und Spalte ermitteln
      let targetWs = null;
      let targetColIdx = null;
      let targetNameColIdx = null;

      if (state.currentRefSheet) {
        targetWs = state.refWorkbook.getWorksheet(state.currentRefSheet);
      }
      if (!targetWs && state.refWorkbook.worksheets.length > 0) {
        targetWs = state.refWorkbook.worksheets[0];
      }

      if (targetWs) {
        targetWs.getRow(1).eachCell((cell, colNum) => {
          const hVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
          if (refColName && hVal.toLowerCase() === refColName.toLowerCase()) {
            targetColIdx = colNum;
          }
          if (["name", "mitarbeiter", "person", "bezeichnung"].some(k => hVal.toLowerCase().includes(k))) {
            targetNameColIdx = colNum;
          }
        });
      }

      if (!targetColIdx) {
        for (const ws of state.refWorkbook.worksheets) {
          ws.getRow(1).eachCell((cell, colNum) => {
            const hVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
            if (refColName && hVal.toLowerCase() === refColName.toLowerCase()) {
              targetWs = ws;
              targetColIdx = colNum;
            }
          });
          if (targetColIdx) break;
        }
      }

      if (!targetColIdx && targetWs) {
        const headerCount = Math.max(targetWs.columnCount || 0, targetWs.actualColumnCount || 0, 1);
        targetColIdx = headerCount + 1;
        targetWs.getRow(1).getCell(targetColIdx).value = refColName || "Stammdaten";
      }

      if (!targetWs || !targetColIdx) {
        alert(`Konnte die Spalte "${refColName}" in der Referenzdatei nicht lokalisieren.`);
        return false;
      }

      const isResourceCol = /ressource|personal|pers|mitarbeiter|worker/i.test(refColName || "");

      // 2. Prüfen, ob die Nummer schon in dieser Spalte existiert
      const maxRows = Math.max(targetWs.rowCount || 0, targetWs.actualRowCount || 0, (targetWs._rows ? targetWs._rows.length : 0));
      let alreadyExists = false;
      let emptyRowIdx = null;

      for (let r = 2; r <= maxRows; r++) {
        const cell = targetWs.getRow(r).getCell(targetColIdx);
        const cellVal = WebExcelEngine.extractCellValue(cell.value, cell).trim();
        if (!cellVal && !emptyRowIdx) {
          emptyRowIdx = r;
        }
        if (cellVal === cleanNum || (/^\d+$/.test(cellVal) && /^\d+$/.test(cleanNum) && parseInt(cellVal, 10) === parseInt(cleanNum, 10))) {
          alreadyExists = true;
          break;
        }
      }

      // 3. Eintragen
      if (!alreadyExists) {
        const insertRowIdx = emptyRowIdx || (maxRows + 1);
        const row = targetWs.getRow(insertRowIdx);
        const cell = row.getCell(targetColIdx);

        if (/^\d{1,4}$/.test(cleanNum) && isResourceCol) {
          cell.value = WebExcelEngine.padNumber(cleanNum, 4);
          cell.numFmt = "0000";
        } else if (/^\d+$/.test(cleanNum) && cleanNum.length < 10) {
          cell.value = parseInt(cleanNum, 10);
        } else {
          cell.value = cleanNum;
        }
        cell.alignment = { horizontal: "center", vertical: "middle" };

        if (optStaffName && targetNameColIdx) {
          const nameCell = row.getCell(targetNameColIdx);
          nameCell.value = optStaffName;
        }
      }

      // 4. Falls Mitarbeiter-Spalte oder Mitarbeiter-Name angegeben:
      if (isResourceCol || optStaffName) {
        const staffWs = state.refWorkbook.worksheets.find(w => /personal|mitarbeiter|stamm/i.test(w.name) && w !== targetWs);
        if (staffWs) {
          let sResCol = 1, sNameCol = 2;
          staffWs.getRow(1).eachCell((cell, colNum) => {
            const hVal = WebExcelEngine.extractCellValue(cell.value, cell).toLowerCase();
            if (/personal|pers|ressource|mitarbeiter/i.test(hVal)) sResCol = colNum;
            if (/name|bezeichnung/i.test(hVal)) sNameCol = colNum;
          });
          const sMax = Math.max(staffWs.rowCount || 0, staffWs.actualRowCount || 0);
          let sFound = false;
          for (let r = 2; r <= sMax; r++) {
            const c = staffWs.getRow(r).getCell(sResCol);
            const cv = WebExcelEngine.extractCellValue(c.value, c).trim();
            if (cv === cleanNum || (/^\d+$/.test(cv) && /^\d+$/.test(cleanNum) && parseInt(cv, 10) === parseInt(cleanNum, 10))) {
              sFound = true;
              break;
            }
          }
          if (!sFound) {
            const newStaffRow = staffWs.getRow(sMax + 1);
            newStaffRow.getCell(sResCol).value = /^\d{1,4}$/.test(cleanNum) ? WebExcelEngine.padNumber(cleanNum, 4) : cleanNum;
            if (optStaffName) newStaffRow.getCell(sNameCol).value = optStaffName;
          }
        }

        const staffDisplayName = optStaffName || getStaffName(cleanNum) || `Mitarbeiter ${cleanNum}`;
        registerStaffEntry(cleanNum, staffDisplayName);
        if (!state.staffList.some(s => s.resource === cleanNum || (parseInt(s.resource, 10) === parseInt(cleanNum, 10)))) {
          state.staffList.push({
            resource: /^\d{1,4}$/.test(cleanNum) ? WebExcelEngine.padNumber(cleanNum, 4) : cleanNum,
            name: staffDisplayName,
            dept: "Referenzdatei"
          });
          saveStaffToStorage(state.staffList);
        }
      }

      // 5. In IndexedDB dauerhaft speichern
      const refBuf = await state.refWorkbook.xlsx.writeBuffer();
      state.lastRefBuffer = refBuf;
      await saveRefFileToStorage(state.refFileName || "Referenz_Stammdaten.xlsx", refBuf, state.currentRefSheet);

      // 6. Neu analysieren & Oberflächen aktualisieren
      extractStaffFromReferenceWorkbook(state.refWorkbook);
      runInspection({ skipScroll: true, skipToast: true });

      showToast(`✅ "${cleanNum}" wurde dauerhaft in die Referenzdatei (${refColName || 'Stammdaten'}) übernommen!`);
      return true;
    } catch (err) {
      console.error("Fehler beim Hinzufügen der Nummer zur Referenzdatei:", err);
      alert("Fehler beim Aktualisieren der Referenzdatei: " + err.message);
      return false;
    }
  }

  // Behandelt das Speichern einer manuellen Korrektur inklusive Referenz-Option
  async function handleManualSaveWithRefOption(targetItem, val, isChecked) {
    if (!targetItem) return;
    const refCol = targetItem.ref_col_name || targetItem.col_name;
    const isResource = /ressource|personal|pers|mitarbeiter|worker/i.test(refCol || "");

    const isInRef = isValueInReferenceColumn(refCol, val);
    let shouldAddToRef = isChecked;

    if (!shouldAddToRef && !isInRef && val) {
      const ask = confirm(
        `Die Nummer "${val}" (${targetItem.col_name}) existiert noch nicht in der Referenzdatei (${refCol}).\n\n` +
        `Möchten Sie diese Nummer jetzt dauerhaft in die Referenzdatei schreiben, damit sie als gültig anerkannt wird?`
      );
      if (ask) {
        shouldAddToRef = true;
      }
    }

    if (shouldAddToRef && val) {
      let optStaffName = null;
      if (isResource) {
        const existingName = getStaffName(val);
        const enteredName = prompt(
          `Mitarbeiter-Name für Nummer "${val}" (optional):`,
          existingName && !existingName.startsWith("Mitarbeiter") ? existingName : ""
        );
        if (enteredName !== null && enteredName.trim() !== "") {
          optStaffName = enteredName.trim();
        }
      }
      await addNumberToReferenceFile(refCol, val, optStaffName);
    }

    applySingleCorrection(targetItem.id, val);
  }

  // --- Ergebnistabelle Rendern ---
  function renderResultsTable() {
    const query = state.searchQuery;
    const filter = state.currentFilter;

    const filtered = state.allResults.filter(r => {
      const isError = (r.status !== "OK" && r.status !== "OK_LEER");
      if (filter === "errors" && !isError) return false;
      if (filter === "zahlendreher" && r.status !== "ZAHLENDREHER" && r.status !== "ZIFFERENTAUSCH") return false;
      if (filter === "tippfehler" && r.status !== "TIPPFEHLER" && !r.status.startsWith("ZIFFER_")) return false;
      if (filter === "datum" && r.status !== "DATUM_WARNUNG") return false;
      if (filter === "nicht_existent" && r.status !== "NICHT_EXISTENT" && r.status !== "LEER" && r.status !== "NAECHSTE_ALTERNATIVE") return false;
      if (filter === "corrected" && !r.is_corrected) return false;

      if (query) {
        const rowStr = `${r.cell} ${r.col_name} ${r.row} ${r.original_value} ${r.current_value} ${r.status_label} ${r.detail} ${r.suggestion}`.toLowerCase();
        if (!rowStr.includes(query)) return false;
      }
      return true;
    });

    if (filtered.length === 0) {
      resultsTbody.innerHTML = `<tr><td colspan="7" style="text-align: center; padding: 2rem; color: var(--text-muted);">Keine Einträge für die aktuellen Filterkriterien gefunden.</td></tr>`;
      return;
    }

    resultsTbody.innerHTML = filtered.map(r => {
      const isCorrected = r.is_corrected;
      const isZahlendreher = (r.status === "ZAHLENDREHER" || r.status === "ZIFFERENTAUSCH");
      const isError = (r.status !== "OK" && r.status !== "OK_LEER");
      const isRefCandidate = (r.status === "NICHT_EXISTENT" || r.status === "LEER" || r.status === "ZAHLENDREHER" || r.status === "TIPPFEHLER" || r.status.startsWith("ZIFFER_"));
      const refColDisp = r.ref_col_name || r.col_name;
      const valForRef = (r.current_value || r.original_value || "").trim();

      let rowClass = "";
      if (isCorrected) rowClass = "row-corrected";
      else if (isZahlendreher) rowClass = "row-zahlendreher";
      else if (isError) rowClass = "row-error";

      const strikedClass = isCorrected ? "is-striked" : "";
      const inputCorrClass = isCorrected ? "is-corrected" : "";

      let suggHtml = "-";
      if (r.suggestion) {
        let otherChips = "";
        if (r.candidates && r.candidates.length > 1) {
          const others = r.candidates.filter(c => c.value !== r.suggestion).slice(0, 3);
          if (others.length > 0) {
            otherChips = `
              <div class="alt-chips">
                <span>Weitere:</span>
                ${others.map(c => `
                  <button type="button" class="alt-chip" data-id="${r.id}" data-val="${escapeHtml(c.value)}" title="${escapeHtml(c.detail)}">
                    ${escapeHtml(c.value)}
                  </button>
                `).join("")}
              </div>
            `;
          }
        }

        suggHtml = `
          <div class="alt-chips-container">
            <button class="btn-apply-sugg" data-id="${r.id}" data-val="${escapeHtml(r.suggestion)}" title="Diesen Vorschlag übernehmen">
              <span class="icon">✓</span> <strong>${escapeHtml(r.suggestion)}</strong> übernehmen
            </button>
            ${otherChips}
          </div>
        `;
      } else if (isRefCandidate && valForRef && valForRef !== "—") {
        suggHtml = `
          <button type="button" class="btn-quick-add-ref" data-id="${r.id}" data-val="${escapeHtml(valForRef)}" data-col="${escapeHtml(refColDisp)}" title="Diese Nummer als neue gültige Nummer dauerhaft in die Referenzdatei schreiben">
            <span class="icon">➕</span> In Referenz (${escapeHtml(refColDisp)}) aufnehmen
          </button>
        `;
      }

      const manualHtml = `
        <div class="manual-edit-box" style="display: flex; flex-direction: column; align-items: flex-start; gap: 0.3rem;">
          <div style="display: flex; align-items: center; gap: 0.35rem; width: 100%;">
            <input type="text" class="manual-input ${inputCorrClass}" 
              data-id="${r.id}" 
              value="${escapeHtml(r.current_value)}" 
              placeholder="Nummer eintragen..."
              title="Nummer manuell anpassen und Enter drücken">
            <button class="btn-save-manual" data-id="${r.id}" title="Händisch speichern">💾 Speichern</button>
          </div>
          ${isRefCandidate ? `
            <label class="cb-ref-label" title="Diesen Wert zusätzlich dauerhaft in die Referenzdatei (${escapeHtml(refColDisp)}) schreiben, sodass er sofort als gültig anerkannt wird">
              <input type="checkbox" class="cb-add-to-ref" data-id="${r.id}">
              <span class="cb-ref-text">In Referenz übernehmen</span>
            </label>
          ` : ''}
        </div>
        ${isCorrected ? '<div style="margin-top: 4px;"><span class="badge badge-corrected">✓ Geändert</span></div>' : ''}
      `;

      const displayOrig = (r.original_value && !String(r.original_value).toLowerCase().includes("object")) ? r.original_value : "—";

      return `
        <tr class="${rowClass}">
          <td><span class="cell-coord">${escapeHtml(r.cell)}</span></td>
          <td><strong>${escapeHtml(r.col_name)}</strong></td>
          <td style="color: var(--text-muted);">Zeile ${r.row}</td>
          <td><span class="val-old ${strikedClass}">${escapeHtml(displayOrig)}</span></td>
          <td>
            <span class="badge ${r.badge_class}">${escapeHtml(r.status_label)}</span>
            <div class="diag-detail" style="margin-top: 3px;">${escapeHtml(r.detail)}</div>
          </td>
          <td>${suggHtml}</td>
          <td>${manualHtml}</td>
        </tr>
      `;
    }).join("");

    // Event Listener für Vorschläge
    resultsTbody.querySelectorAll(".btn-apply-sugg").forEach(btn => {
      btn.addEventListener("click", () => {
        applySingleCorrection(btn.dataset.id, btn.dataset.val);
      });
    });

    resultsTbody.querySelectorAll(".alt-chip").forEach(chip => {
      chip.addEventListener("click", () => {
        applySingleCorrection(chip.dataset.id, chip.dataset.val);
      });
    });

    // Event Listener für 1-Klick-Übernahme in Referenzdatei
    resultsTbody.querySelectorAll(".btn-quick-add-ref").forEach(btn => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.id;
        const val = btn.dataset.val;
        const col = btn.dataset.col;
        const targetItem = state.allResults.find(r => r.id === id);
        if (!targetItem || !val) return;

        const isResource = /ressource|personal|pers|mitarbeiter|worker/i.test(col || "");
        let optStaffName = null;
        if (isResource) {
          const existingName = getStaffName(val);
          const enteredName = prompt(
            `Nummer "${val}" als neu in die Referenzdatei (${col}) aufnehmen.\n\nName des Mitarbeiters (optional):`,
            existingName && !existingName.startsWith("Mitarbeiter") ? existingName : ""
          );
          if (enteredName === null) return;
          if (enteredName.trim() !== "") optStaffName = enteredName.trim();
        } else {
          const confirmAdd = confirm(`Möchten Sie die Nummer "${val}" als neue gültige Nummer dauerhaft in die Referenzdatei (${col}) übernehmen?`);
          if (!confirmAdd) return;
        }

        await addNumberToReferenceFile(col, val, optStaffName);
      });
    });

    // Event Listener für manuelles Speichern (inkl. Referenzdatei-Option)
    resultsTbody.querySelectorAll(".btn-save-manual").forEach(btn => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.id;
        const input = resultsTbody.querySelector(`.manual-input[data-id="${id}"]`);
        const chkRef = resultsTbody.querySelector(`.cb-add-to-ref[data-id="${id}"]`);
        if (input) {
          const val = input.value.trim();
          const targetItem = state.allResults.find(r => r.id === id);
          const isChecked = chkRef ? chkRef.checked : false;
          await handleManualSaveWithRefOption(targetItem, val, isChecked);
        }
      });
    });

    resultsTbody.querySelectorAll(".manual-input").forEach(input => {
      const handleInputLive = () => {
        const id = input.dataset.id;
        const val = input.value.trim();
        const targetItem = state.allResults ? state.allResults.find(r => r.id === id) : null;
        if (targetItem) {
          targetItem.current_value = val;
          targetItem.is_corrected = (val !== targetItem.original_value);
          if (val !== targetItem.original_value) {
            state.appliedCorrections[id] = val;
          } else {
            delete state.appliedCorrections[id];
          }

          // Wert auch sofort direkt in die Arbeitsmappe schreiben
          if (state.targetWorkbook && state.currentTargetSheet) {
            const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
            if (ws) {
              const row = ws.getRow(targetItem.row);
              if (row) {
                const cell = row.getCell(targetItem.col_idx);
                if (val === "") {
                  cell.value = null;
                } else if (targetItem.col_idx === 1) {
                  const dParts = val.split("-");
                  let dispDate = val;
                  if (dParts.length === 3) dispDate = `${dParts[2]}.${dParts[1]}.${dParts[0]}`;
                  cell.value = dispDate;
                  cell.numFmt = "DD.MM.YYYY";
                } else if (/^\d{1,4}$/.test(val) && (targetItem.col_idx === 3 || (targetItem.col_letter && targetItem.col_letter.toUpperCase() === "C"))) {
                  cell.value = WebExcelEngine.padNumber(val, 4);
                  cell.numFmt = "0000";
                } else if (targetItem.col_idx === 4 || targetItem.col_idx === 5) {
                  let unpadded = val.replace(/^0+/, "") || "0";
                  cell.value = parseInt(unpadded, 10);
                  cell.numFmt = "0";
                } else if (/^\d+$/.test(val)) {
                  cell.value = parseInt(val, 10);
                } else if (/^-?\d+(\.\d+)?$/.test(val.replace(",", "."))) {
                  const num = parseFloat(val.replace(",", "."));
                  cell.value = isNaN(num) ? val : num;
                } else {
                  cell.value = val;
                }
              }
            }
          }
        }
      };

      input.addEventListener("input", handleInputLive);
      input.addEventListener("change", handleInputLive);
      input.addEventListener("blur", handleInputLive);

      input.addEventListener("keydown", async (e) => {
        if (e.key === "Enter") {
          const id = input.dataset.id;
          const val = input.value.trim();
          const targetItem = state.allResults.find(r => r.id === id);
          const chkRef = resultsTbody.querySelector(`.cb-add-to-ref[data-id="${id}"]`);
          const isChecked = chkRef ? chkRef.checked : false;
          input.blur();
          await handleManualSaveWithRefOption(targetItem, val, isChecked);
        }
      });
    });
  }

  function applySingleCorrection(cellId, newVal) {
    const targetItem = state.allResults.find(r => r.id === cellId);
    if (!targetItem) return;

    const origVal = targetItem.original_value;
    const colIdx = targetItem.col_idx;
    let count = 0;

    let tgtWs = null;
    if (state.targetWorkbook && state.currentTargetSheet) {
      tgtWs = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    }

    state.allResults.forEach(r => {
      const isSameMatch = (targetItem && origVal !== "" && r.col_idx === colIdx && r.original_value === origVal);
      if (r.id === cellId || isSameMatch) {
        r.current_value = newVal;
        r.is_corrected = (newVal !== r.original_value);
        if (newVal !== r.original_value) {
          state.appliedCorrections[r.id] = newVal;
        } else {
          delete state.appliedCorrections[r.id];
        }
        count++;

        if (tgtWs) {
          const cell = tgtWs.getRow(r.row).getCell(r.col_idx);
          const valStr = String(newVal).trim();
          if (valStr === "") {
            cell.value = null;
          } else if (/^\d{1,4}$/.test(valStr) && (r.col_idx === 3 || (targetItem.col_letter && targetItem.col_letter.toUpperCase() === "C"))) {
            cell.value = WebExcelEngine.padNumber(valStr, 4);
            cell.numFmt = "0000";
          } else if (r.col_idx === 4 || r.col_idx === 5) {
            let unpadded = valStr.replace(/^0+/, "") || "0";
            cell.value = parseInt(unpadded, 10);
            cell.numFmt = "0";
          } else if (/^\d+$/.test(valStr) && valStr.length < 10) {
            cell.value = parseInt(valStr, 10);
          } else if (/^-?\d+(\.\d+)?$/.test(valStr.replace(",", "."))) {
            const num = parseFloat(valStr.replace(",", "."));
            cell.value = isNaN(num) ? valStr : num;
          } else {
            cell.value = valStr;
          }
        }
      }
    });

    updateFilterCounts(state.allResults);
    renderResultsTable();
    if (state.timesheetData) renderTimesheetMatrix();
    renderFullExcelTable();

    if (count > 1) {
      showToast(`Korrektur "${newVal}" für alle ${count} Zellen mit Wert "${origVal}" übernommen!`);
    } else {
      showToast(`Zelle ${targetItem.cell} aktualisiert auf "${newVal}".`);
    }
  }

  // --- GENERAL-BUTTON: Datei erstellen & Änderungen übernehmen ---
  if (btnGeneralApply) {
    btnGeneralApply.addEventListener("click", async () => {
      if (!state.targetWorkbook || !state.currentTargetSheet) {
        alert("Bitte laden Sie zuerst eine zu prüfende Excel-Datei.");
        return;
      }

      // Falls die Prüfung noch nicht lief, automatisch einmal im Hintergrund ausführen
      if (!state.allResults || state.allResults.length === 0) {
        runInspection({ skipScroll: true, skipToast: true });
      }

      btnGeneralApply.disabled = true;
      btnGeneralApply.innerHTML = `<span class="icon">⏳</span> Erstelle korrigierte Excel-Datei...`;
      showToast("General-Button: Erstelle korrigierte Excel-Datei...");
      console.log("[GENERAL-BTN] Gestartet. appliedCorrections:", JSON.stringify(state.appliedCorrections));

      // 1. Alle im Formular manuell eingegebenen Werte sammeln
      if (resultsTbody) {
        resultsTbody.querySelectorAll(".manual-input").forEach(input => {
          const id = input.dataset.id;
          const val = input.value.trim();
          const item = state.allResults ? state.allResults.find(r => r.id === id) : null;
          if (item && val !== item.original_value) {
            item.current_value = val;
            item.is_corrected = true;
            state.appliedCorrections[id] = val;
          }
        });
      }

      // 2. Excel-Datei 1:1 klonen und aufbauen
      try {
        const outBuffer = await generateCorrectedExcel();
        const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
        const outFileName = `${baseName}_KORRIGIERT.xlsx`;

        // 3. Browser-Download anstoßen
        downloadBuffer(outBuffer, outFileName);

        btnGeneralApply.disabled = false;
        btnGeneralApply.innerHTML = `<span class="icon">🚀</span> <strong>General-Button: Datei erstellen &amp; Änderungen übernehmen</strong>`;

        // Erfolgsbanner einblenden
        const banner = document.getElementById("download-success-banner");
        const nameEl = document.getElementById("download-file-name");
        if (banner && nameEl) {
          nameEl.textContent = outFileName;
          banner.classList.remove("hidden");
          banner.scrollIntoView({ behavior: "smooth" });
        }

        if (btnExportExcel) btnExportExcel.classList.remove("hidden");
        if (btnExportReport) btnExportReport.classList.remove("hidden");

        renderResultsTable();
        updateFilterCounts(state.allResults);
        if (state.timesheetData) renderTimesheetMatrix();
        renderFullExcelTable();
        showToast(`✅ Datei "${outFileName}" erfolgreich im Download-Ordner gespeichert!`);
      } catch (err) {
        console.error(err);
        btnGeneralApply.disabled = false;
        btnGeneralApply.innerHTML = `<span class="icon">🚀</span> <strong>General-Button: Datei erstellen &amp; Änderungen übernehmen</strong>`;
        alert("Fehler beim Erstellen der Datei: " + err.message);
      }
    });
  }

  // --- Nur Zahlendreher korrigieren ---
  if (btnFixZahlendreher) {
    btnFixZahlendreher.addEventListener("click", () => {
      let count = 0;
      if (state.allResults) {
        state.allResults.forEach(r => {
          if ((r.status === "ZAHLENDREHER" || r.status === "ZIFFERENTAUSCH") && r.suggestion) {
            r.current_value = r.suggestion;
            r.is_corrected = true;
            state.appliedCorrections[r.id] = r.suggestion;
            count++;
          }
        });
      }
      updateFilterCounts(state.allResults);
      renderResultsTable();
      if (state.timesheetData) renderTimesheetMatrix();
      renderFullExcelTable();
      showToast(`⚡ ${count} Zahlendreher wurden automatisch korrigiert!`);
    });
  }

  // --- Zurücksetzen ---
  if (btnResetFixes) {
    btnResetFixes.addEventListener("click", () => {
      if (!confirm("Alle Korrekturen rückgängig machen?")) return;
      if (state.allResults) {
        state.allResults.forEach(r => {
          r.current_value = r.original_value;
          r.is_corrected = false;
        });
      }
      state.appliedCorrections = {};
      updateFilterCounts(state.allResults);
      renderResultsTable();
      if (state.timesheetData) renderTimesheetMatrix();
      renderFullExcelTable();
      showToast("Alle Korrekturen wurden zurückgesetzt.");
    });
  }

  // --- Nochmals herunterladen ---
  const btnDownloadAgain = document.getElementById("btn-download-again");
  if (btnDownloadAgain) {
    btnDownloadAgain.addEventListener("click", async () => {
      const outBuffer = await generateCorrectedExcel();
      const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
      downloadBuffer(outBuffer, `${baseName}_KORRIGIERT.xlsx`);
      showToast("📥 Datei erneut heruntergeladen.");
    });
  }

  if (btnExportExcel) {
    btnExportExcel.addEventListener("click", async () => {
      const outBuffer = await generateCorrectedExcel();
      const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
      downloadBuffer(outBuffer, `${baseName}_KORRIGIERT.xlsx`);
      showToast("📥 Korrigierte Excel wird heruntergeladen...");
    });
  }

  if (btnExportReport) {
    btnExportReport.addEventListener("click", async () => {
      const reportBuffer = await generateReportExcel();
      const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
      downloadBuffer(reportBuffer, `Pruefbericht_${baseName}.xlsx`);
      showToast("📊 Prüfbericht wird heruntergeladen...");
    });
  }

  // --- Excel-Erstellung im 1:1 Original-Layout mit 0000 Formatierung ---
  async function generateCorrectedExcel() {
    if (!state.targetWorkbook || !state.currentTargetSheet) {
      throw new Error("Keine Prüfdatei geladen.");
    }
    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) {
      throw new Error("Arbeitsblatt in der Prüfdatei nicht gefunden.");
    }

    console.log("[generateCorrectedExcel] START");
    console.log("[generateCorrectedExcel] appliedCorrections KEYS:", Object.keys(state.appliedCorrections));
    console.log("[generateCorrectedExcel] appliedCorrections:", JSON.stringify(state.appliedCorrections));

    // 1. Manuelle Eingaben aus geöffneten Eingabefeldern in der Prüfliste einsammeln
    if (resultsTbody) {
      const manualInputs = Array.from(resultsTbody.querySelectorAll(".manual-input"));
      for (const input of manualInputs) {
        const id = input.dataset.id;
        const val = input.value.trim();
        const item = state.allResults ? state.allResults.find(r => r.id === id) : null;
        if (item && val !== item.original_value) {
          item.current_value = val;
          item.is_corrected = true;
          state.appliedCorrections[id] = val;
        }
      }
    }

    let dateColIdx = parseInt(selectTsDateCol?.value, 10) || 1;
    let resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;
    let orderColIdx = parseInt(document.getElementById("select-ts-order-col")?.value, 10) || 4;

    // Helper: Schreibt einen Wert mit korrekter Formatierung in eine Zelle
    function writeValueToCell(rowNum, colNum, rawVal) {
      const row = ws.getRow(rowNum);
      if (!row) return;
      const cell = row.getCell(colNum);

      if (rawVal === null || rawVal === undefined || rawVal === "") {
        cell.value = null;
        return;
      }

      const valStr = String(rawVal).trim();

      if (colNum === dateColIdx || colNum === 1) {
        const dParts = valStr.split("-");
        let dispDate = valStr;
        if (dParts.length === 3) {
          dispDate = `${dParts[2]}.${dParts[1]}.${dParts[0]}`;
        }
        cell.value = dispDate;
        cell.numFmt = "DD.MM.YYYY";
      } else if (colNum === 3) {
        // Spalte C: AUSSCHLIESSLICH 0000 Formatierung (mit führender Null, 4-stellig)
        let cleanRes = valStr.replace(/[,.]0+$/, "").trim();
        if (/^\d{1,4}$/.test(cleanRes)) {
          cell.value = WebExcelEngine.padNumber(cleanRes, 4);
          cell.numFmt = "0000";
        } else {
          cell.value = cleanRes;
          cell.numFmt = "0000";
        }
      } else if (colNum === 4 || colNum === 5) {
        // Spalte D & Spalte E: Einfache Zahlen OHNE führende Null (z.B. 8 statt 0008)
        let cleanVal = valStr.replace(/[,.]0+$/, "").trim();
        let unpadded = cleanVal.replace(/^0+/, "") || "0";
        if (/^-?\d+$/.test(unpadded)) {
          cell.value = parseInt(unpadded, 10);
          cell.numFmt = "0";
        } else if (/^-?\d+([.,]\d+)?$/.test(cleanVal)) {
          const num = Math.round(parseFloat(cleanVal.replace(",", ".")));
          cell.value = num;
          cell.numFmt = "0";
        } else {
          cell.value = cleanVal;
          cell.numFmt = "0";
        }
      } else if (/^-?\d+(\.\d+)?$/.test(valStr.replace(",", "."))) {
        const num = parseFloat(valStr.replace(",", "."));
        if (!isNaN(num)) {
          cell.value = num;
        } else {
          cell.value = valStr;
        }
      } else {
        cell.value = valStr;
      }
    }

    // 2. Eindeutige Korrekturvorschläge aus den Prüfergebnissen verarbeiten
    // 2. Automatisches Übernehmen von klaren Empfehlungen (Zahlendreher)
    if (state.allResults && Array.isArray(state.allResults)) {
      state.allResults.forEach(r => {
        if (state.appliedCorrections && state.appliedCorrections[r.id] !== undefined) {
          // Bereits in appliedCorrections vorhanden (z. B. vom Nutzer manuell geändert)
        } else if (r.current_value && r.current_value !== r.original_value) {
          applySingleCorrection(r.id, r.current_value);
        } else if (r.suggestion && r.status !== "DATUM_WARNUNG" && r.status !== "NICHT_EXISTENT" && r.status !== "LEER") {
          applySingleCorrection(r.id, r.suggestion);
        }
      });
    }

    // 3. ALLE Korrekturen aus state.appliedCorrections (Wochenübersicht, Modale, Inline, Ergebnisse) in die Arbeitsmappe schreiben
    if (state.appliedCorrections) {
      console.log("[generateCorrectedExcel] Schritt 3: Schreibe", Object.keys(state.appliedCorrections).length, "Korrekturen in Arbeitsmappe");
      Object.entries(state.appliedCorrections).forEach(([key, valToApply]) => {
        if (key.includes("_")) {
          const parts = key.split("_");
          if (parts.length === 2) {
            const rNum = parseInt(parts[0], 10);
            const cNum = parseInt(parts[1], 10);
            if (!isNaN(rNum) && !isNaN(cNum) && rNum > 0 && cNum > 0) {
              writeValueToCell(rNum, cNum, valToApply);
              console.log(`[generateCorrectedExcel] Schreibe Zeile ${rNum}, Spalte ${cNum} = "${valToApply}"`);
            }
          }
        }
      });
    }

    // 4. Spalten-Formatierungen setzen (Datum DD.MM.YYYY, Ressource 0000 in C, Spalte D & E ganze Zahlen)
    const maxCols = Math.max(ws.columnCount || 0, ws.actualRowCount ? ws.actualColumnCount || 0 : 0, 10);
    const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length : 0));

    for (let colIdx = 1; colIdx <= maxCols; colIdx++) {
      const col = ws.getColumn(colIdx);
      const isDateCol = (colIdx === dateColIdx || colIdx === 1);
      const isResCol = (colIdx === 3); // EXKLUSIV Spalte C!
      const isSimpleNumCol = (colIdx === 4 || colIdx === 5); // Spalte D und Spalte E!

      if (isDateCol) {
        col.numFmt = "DD.MM.YYYY";
        for (let r = 2; r <= maxRows; r++) {
          const cell = ws.getRow(r).getCell(colIdx);
          if (typeof cell.value === "number") {
            cell.numFmt = "DD.MM.YYYY";
          }
        }
      } else if (isResCol) {
        // EXKLUSIV Spalte C: 0000 Formatierung (4-stellig mit führenden Nullen)
        col.numFmt = "0000";
        for (let r = 2; r <= maxRows; r++) {
          const cell = ws.getRow(r).getCell(colIdx);
          cell.numFmt = "0000";
          if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
            const s = WebExcelEngine.extractCellValue(cell.value, cell).replace(/[,.]0+$/, "").trim();
            if (/^\d{1,4}$/.test(s)) {
              cell.value = WebExcelEngine.padNumber(s, 4);
            } else if (/^\d+$/.test(s)) {
              cell.value = parseInt(s, 10);
            }
          }
        }
      } else if (isSimpleNumCol) {
        // Spalte D und Spalte E: Einfache Zahlen OHNE führende Nullen
        col.numFmt = "0";
        for (let r = 2; r <= maxRows; r++) {
          const cell = ws.getRow(r).getCell(colIdx);
          cell.numFmt = "0";
          if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
            const rawVal = WebExcelEngine.extractCellValue(cell.value, cell);
            let s = String(rawVal).trim().replace(",", ".");
            s = s.replace(/\.0+$/, "");
            let unpadded = s.replace(/^0+/, "") || "0";
            if (/^-?\d+$/.test(unpadded)) {
              cell.value = parseInt(unpadded, 10); // Wandelt "0008" oder "08" garantiert in 8 um!
            } else if (/^-?\d+\.\d+$/.test(s)) {
              const num = Math.round(parseFloat(s));
              if (!isNaN(num)) cell.value = num;
            }
          }
        }
      }
    }

    // Vor dem Schreiben: Prüfe ob Werte tatsächlich in der Arbeitsmappe stehen
    const verifyWs = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (verifyWs && Object.keys(state.appliedCorrections).length > 0) {
      const firstKey = Object.keys(state.appliedCorrections).find(k => k.includes("_"));
      if (firstKey) {
        const [vr, vc] = firstKey.split("_").map(Number);
        if (vr > 0 && vc > 0) {
          const verifyCell = verifyWs.getRow(vr).getCell(vc);
          console.log(`[generateCorrectedExcel] VERIFY: Zelle ${firstKey} hat Wert "${verifyCell.value}" (soll: "${state.appliedCorrections[firstKey]}"`);
        }
      }
    }
    console.log("[generateCorrectedExcel] Rufe writeBuffer() auf...");
    const buffer = await state.targetWorkbook.xlsx.writeBuffer();
    console.log(`[generateCorrectedExcel] writeBuffer() fertig: ${buffer.byteLength} bytes`);
    return buffer;
  }

  // --- Detaillierten Prüfbericht als Excel erstellen ---
  async function generateReportExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Prüfbericht");

    ws.columns = [
      { header: "Zelle", key: "cell", width: 10 },
      { header: "Zeile", key: "row", width: 8 },
      { header: "Spalte", key: "col_name", width: 22 },
      { header: "Originalwert", key: "orig", width: 18 },
      { header: "Aktueller Wert", key: "curr", width: 18 },
      { header: "Status", key: "status_lbl", width: 25 },
      { header: "Vorschlag", key: "sugg", width: 18 },
      { header: "Diagnosedetails", key: "detail", width: 45 },
      { header: "Wurde Korrigiert?", key: "is_corr", width: 18 }
    ];

    const hdrRow = ws.getRow(1);
    hdrRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    hdrRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E293B" }
    };

    state.allResults.forEach(r => {
      ws.addRow({
        cell: r.cell,
        row: r.row,
        col_name: r.col_name,
        orig: r.original_value,
        curr: r.current_value,
        status_lbl: r.status_label,
        sugg: r.suggestion || "—",
        detail: r.detail,
        is_corr: r.is_corrected ? "Ja" : "Nein"
      });
    });

    return await wb.xlsx.writeBuffer();
  }

  function downloadBuffer(buffer, fileName) {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const isFileProtocol = (window.location.protocol === "file:");

    if (!isFileProtocol) {
      // Standard-Methode: Blob-URL (funktioniert zuverlässig über http/https)
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
      console.log("[downloadBuffer] Blob-URL Download für:", fileName);
    } else {
      // Fallback für file:// Protokoll: Data-URL
      const reader = new FileReader();
      reader.onloadend = function() {
        const a = document.createElement("a");
        a.href = reader.result;
        a.download = fileName;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        setTimeout(() => a.remove(), 2000);
        console.log("[downloadBuffer] Data-URL Download für:", fileName);
      };
      reader.readAsDataURL(blob);
    }
  }

  // ==============================================================================
  // Arbeitszeiten-Übersicht & Mitarbeiter-Matrix (Schritt 4)
  // ==============================================================================

  function populateTimesheetColSelects(tgtHeaders) {
    if (!selectTsDateCol || !selectTsHoursCol || !selectTsResourceCol) return;

    selectTsDateCol.innerHTML = "";
    selectTsHoursCol.innerHTML = "";
    selectTsResourceCol.innerHTML = "";

    // 1. Suche nach Datumsspalte
    let bestDateCol = null;
    tgtHeaders.forEach(h => {
      const lower = h.name.toLowerCase();
      if (!bestDateCol && (lower.includes("datum") || lower.includes("date") || lower.includes("tag") || lower.includes("zeitpunkt") || lower.includes("buchung"))) {
        bestDateCol = h.colNum;
      }
    });
    if (!bestDateCol && tgtHeaders.length > 0) {
      bestDateCol = tgtHeaders[0].colNum;
    }

    // 2. Suche nach Stundenspalte / Menge
    let bestHoursCol = null;
    tgtHeaders.forEach(h => {
      const lower = h.name.toLowerCase();
      if (!bestHoursCol && (lower.includes("stunde") || lower.includes("zeit") || lower.includes("dauer") || lower.includes("ist-stunden") || lower.includes("menge") || lower.includes("std") || lower === "h")) {
        bestHoursCol = h.colNum;
      }
    });
    if (!bestHoursCol) {
      const mengeCol = tgtHeaders.find(h => h.name.toLowerCase().includes("menge") || h.letter === "D" || h.letter === "G");
      if (mengeCol) bestHoursCol = mengeCol.colNum;
      else if (tgtHeaders.length >= 4) bestHoursCol = tgtHeaders[3].colNum;
    }

    // 3. Suche nach Ressourcenspalte
    let bestResCol = null;

    // Priorität 1: Falls Mitarbeiter geladen sind, prüfe die ersten Zeilen der Ziel-Tabelle auf die meisten Treffer mit Mitarbeiternummern!
    if (state.staffList && state.staffList.length > 0 && state.targetWorkbook && state.currentTargetSheet) {
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      if (ws) {
        const colMatchCounts = {};
        const maxSampleRows = Math.min(Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0)), 80);
        for (let r = 2; r <= maxSampleRows; r++) {
          const sampleRow = ws.getRow(r);
          if (!sampleRow) continue;
          tgtHeaders.forEach(h => {
            const val = WebExcelEngine.extractCellValue(sampleRow.getCell(h.colNum).value, sampleRow.getCell(h.colNum));
            if (val && getStaffName(val)) {
              colMatchCounts[h.colNum] = (colMatchCounts[h.colNum] || 0) + 1;
            }
          });
        }
        let maxMatches = 0;
        let matchedColNum = null;
        for (const [colNumStr, count] of Object.entries(colMatchCounts)) {
          if (count > maxMatches) {
            maxMatches = count;
            matchedColNum = parseInt(colNumStr, 10);
          }
        }
        if (maxMatches >= 1) {
          bestResCol = matchedColNum;
        }
      }
    }

    // Priorität 2: Suche nach eindeutigen Spalten-Überschriften
    if (!bestResCol) {
      tgtHeaders.forEach(h => {
        const lower = h.name.toLowerCase();
        if (!bestResCol && (
          lower.includes("ressource") || lower.includes("resource") || lower.includes("personalnummer") ||
          lower.includes("persnr") || lower.includes("pers-nr") || lower.includes("personal-nr") ||
          lower.includes("mitarbeiternr") || lower.includes("mitarbeiter-nr") || lower.includes("mitarbeiter") ||
          lower.includes("monteur") || lower.includes("techniker") || lower.includes("arbeiter") ||
          lower.includes("personal") || lower.includes("manr") || lower.includes("ma-nr") || lower.includes("kollege")
        )) {
          bestResCol = h.colNum;
        }
      });
    }

    // Priorität 3: Mapping-Prüfung
    if (!bestResCol && state.mapping) {
      const mappedRes = state.mapping.find(m => m.selected && m.ref_col_name &&
        ["ressource", "resource", "personalnummer", "persnr"].some(k => m.ref_col_name.toLowerCase().includes(k))
      );
      if (mappedRes) bestResCol = mappedRes.target_col_idx;
    }

    // Priorität 4: Fallback auf Spalte C / 3
    if (!bestResCol) {
      const colC = tgtHeaders.find(h => h.letter === "C" || h.colNum === 3);
      if (colC) bestResCol = colC.colNum;
      else if (tgtHeaders.length >= 3) bestResCol = tgtHeaders[2].colNum;
    }

    // Optionen für Datum
    tgtHeaders.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h.colNum;
      opt.textContent = `Spalte ${h.letter}: ${h.name}`;
      if (h.colNum === bestDateCol) opt.selected = true;
      selectTsDateCol.appendChild(opt);
    });
    const optNoDate = document.createElement("option");
    optNoDate.value = "none";
    optNoDate.textContent = "-- Ohne Datum (Gesamtüberblick) --";
    selectTsDateCol.appendChild(optNoDate);

    // Optionen für Stunden
    tgtHeaders.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h.colNum;
      opt.textContent = `Spalte ${h.letter}: ${h.name}`;
      if (h.colNum === bestHoursCol) opt.selected = true;
      selectTsHoursCol.appendChild(opt);
    });
    const optFixedOne = document.createElement("option");
    optFixedOne.value = "fixed_one";
    optFixedOne.textContent = "-- Pauschal 1 Std pro Buchung --";
    selectTsHoursCol.appendChild(optFixedOne);

    // Optionen für Ressourcen
    tgtHeaders.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h.colNum;
      opt.textContent = `Spalte ${h.letter}: ${h.name}`;
      if (h.colNum === bestResCol) opt.selected = true;
      selectTsResourceCol.appendChild(opt);
    });
  }

  // Event Listener für Spaltenauswahl & Filter der Arbeitszeiten
  [selectTsDateCol, selectTsHoursCol, selectTsResourceCol, chkTsOnlyStaff].forEach(sel => {
    if (sel) {
      sel.addEventListener("change", () => {
        if (state.targetWorkbook && state.currentTargetSheet) {
          renderTimesheetMatrix();
          populateFullTableFilters();
          renderFullExcelTable();
        }
      });
    }
  });

  function parseRowDate(cellVal, cell) {
    if (!cellVal && cellVal !== 0) return { isoKey: "ohne_datum", displayDate: "Ohne Datum", weekday: "", label: "Ohne Datum" };

    let d = null;
    if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
      d = cellVal;
    } else if (typeof cellVal === "number" && cellVal > 1000 && cellVal < 100000) {
      // Excel-Seriennummer (25569 Tage zwischen 01.01.1900 und 01.01.1970)
      const ms = Math.round((cellVal - 25569) * 86400 * 1000);
      d = new Date(ms);
    } else {
      let str = "";
      if (typeof cellVal === "object" && cellVal) {
        str = String(cellVal.result || cellVal.text || "").trim();
      } else {
        str = String(cellVal).trim();
      }

      // Format TT.MM.JJJJ oder TT.MM.JJ
      const deMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
      if (deMatch) {
        let day = parseInt(deMatch[1], 10);
        let month = parseInt(deMatch[2], 10) - 1;
        let year = parseInt(deMatch[3], 10);
        if (year < 100) year += 2000;
        d = new Date(year, month, day);
      } else {
        // Format JJJJ-MM-TT
        const isoMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (isoMatch) {
          d = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
        } else {
          const parsed = Date.parse(str);
          if (!isNaN(parsed)) {
            d = new Date(parsed);
          }
        }
      }
    }

    if (!d || isNaN(d.getTime())) {
      const clean = String(cellVal).trim() || "Ohne Datum";
      return { isoKey: clean, displayDate: clean, weekday: "", label: clean };
    }

    const pad2 = (n) => String(n).padStart(2, "0");
    const yyyy = d.getFullYear();
    const mm = pad2(d.getMonth() + 1);
    const dd = pad2(d.getDate());
    const isoKey = `${yyyy}-${mm}-${dd}`;
    const displayDate = `${dd}.${mm}.${yyyy}`;
    const WEEKDAYS = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
    const weekday = WEEKDAYS[d.getDay()] || "";
    const label = weekday ? `${weekday}, ${dd}.${mm}.` : displayDate;

    return { isoKey, displayDate, weekday, label };
  }

  function parseHours(cellVal) {
    if (cellVal === null || cellVal === undefined || cellVal === "") return 0;
    if (typeof cellVal === "number") return isNaN(cellVal) ? 0 : cellVal;

    let str = "";
    if (typeof cellVal === "object") {
      str = String(cellVal.result || cellVal.text || "").trim();
    } else {
      str = String(cellVal).trim();
    }
    if (!str) return 0;

    // Zeitformat z.B. "08:30" oder "8:15"
    const timeMatch = str.match(/^(\d{1,2}):(\d{2})$/);
    if (timeMatch) {
      const h = parseInt(timeMatch[1], 10);
      const m = parseInt(timeMatch[2], 10);
      return h + (m / 60);
    }

    // Komma durch Punkt ersetzen und Zahlen filtern
    str = str.replace(/[^0-9,.-]/g, "").replace(",", ".");
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  function formatHoursDe(hours) {
    if (hours === 0) return "0,0";
    return Number(hours).toLocaleString("de-DE", { minimumFractionDigits: 1, maximumFractionDigits: 2 });
  }

  function renderTimesheetMatrix() {
    if (!timesheetSection || !state.targetWorkbook || !state.currentTargetSheet) {
      if (timesheetSection) timesheetSection.classList.add("hidden");
      return;
    }

    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) {
      timesheetSection.classList.add("hidden");
      return;
    }

    const dateColVal = selectTsDateCol?.value;
    const isNoDate = (dateColVal === "none");
    const dateColIdx = isNoDate ? null : parseInt(dateColVal, 10);

    const hoursColVal = selectTsHoursCol?.value;
    const isFixedHours = (hoursColVal === "fixed_one");
    const hoursColIdx = isFixedHours ? null : parseInt(hoursColVal, 10);

    const resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;

    const onlyStaff = chkTsOnlyStaff ? chkTsOnlyStaff.checked : true;
    const hasStaffList = state.staffList && state.staffList.length > 0;

    let skippedNonStaffRows = 0;
    const skippedNonStaffResources = new Set();

    // Maps für Aggregation
    const dateMap = new Map();
    const empMap = new Map();

    let grandTotalHours = 0;
    let grandTotalEntries = 0;

    const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0));
    for (let r = 2; r <= maxRows; r++) {
      const row = ws.getRow(r);
      if (!row) continue;

      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;

      // 1. Ressource ermitteln (inklusive etwaiger Benutzer-Korrekturen)
      const cellKey = `${r}_${resColIdx}`;
      let rawRes = (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined)
        ? String(state.appliedCorrections[cellKey]).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
      rawRes = rawRes.replace(/[,.]0+$/, "");

      if (!rawRes) continue; // Zeilen ohne Ressourcennummer überspringen

      // Mitarbeiter-Name aus Stammdaten ermitteln
      const empNameFromMap = getStaffName(rawRes);
      const isKnownWorker = Boolean(empNameFromMap);

      // WICHTIG: Wenn Mitarbeiterdatei hinterlegt ist und Filter aktiv ist,
      // ausschließlich echte Arbeiter anzeigen! Fremd-/Maschinen-Ressourcen überspringen.
      if (hasStaffList && onlyStaff && !isKnownWorker) {
        skippedNonStaffRows++;
        skippedNonStaffResources.add(rawRes);
        continue;
      }

      const paddedRes = (/^\d+$/.test(rawRes) && rawRes.length < 4) ? WebExcelEngine.padNumber(rawRes, 4) : rawRes;
      const displayRes = paddedRes;
      const empName = empNameFromMap || `Ressource ${displayRes}`;
      const hasCustomName = Boolean(empNameFromMap);

      // 2. Datum ermitteln
      let dateObj = null;
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        dateObj = parseRowDate(cell.value, cell);
      } else {
        dateObj = { isoKey: "gesamt", displayDate: "Gesamtzeit", weekday: "", label: "Gesamtzeit" };
      }

      // 3. Stunden ermitteln
      let hours = 0;
      if (isFixedHours) {
        hours = 1;
      } else if (hoursColIdx) {
        const cell = row.getCell(hoursColIdx);
        hours = parseHours(cell.value);
      } else {
        hours = 1;
      }

      // Datum aggregieren
      if (!dateMap.has(dateObj.isoKey)) {
        dateMap.set(dateObj.isoKey, {
          isoKey: dateObj.isoKey,
          displayDate: dateObj.displayDate,
          weekday: dateObj.weekday,
          label: dateObj.label,
          totalHours: 0
        });
      }
      dateMap.get(dateObj.isoKey).totalHours += hours;

      // Mitarbeiter aggregieren (Eindeutiger Schlüssel: empName falls bekannt, sonst displayRes)
      const empKey = isKnownWorker ? empName : displayRes;
      if (!empMap.has(empKey)) {
        empMap.set(empKey, {
          resource: displayRes,
          name: empName,
          hasCustomName: hasCustomName,
          hoursByDate: {},
          totalHours: 0,
          totalEntries: 0
        });
      }
      const empRecord = empMap.get(empKey);
      empRecord.hoursByDate[dateObj.isoKey] = (empRecord.hoursByDate[dateObj.isoKey] || 0) + hours;
      empRecord.totalHours += hours;
      empRecord.totalEntries += 1;

      grandTotalHours += hours;
      grandTotalEntries += 1;
    }

    // Sortierung der Kalendertage
    const sortedDates = Array.from(dateMap.values()).sort((a, b) => a.isoKey.localeCompare(b.isoKey));

    // Sortierung der Mitarbeiter nach Ressourcennummer aufsteigend
    const sortedEmployees = Array.from(empMap.values()).sort((a, b) => {
      const numA = parseInt(a.resource, 10);
      const numB = parseInt(b.resource, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.resource.localeCompare(b.resource);
    });

    state.timesheetData = {
      sortedDates,
      employees: sortedEmployees,
      grandTotalHours,
      grandTotalEntries
    };

    // Status-Badge aktualisieren
    if (tsStaffStatusBadge) {
      if (hasStaffList) {
        const workerCount = sortedEmployees.length;
        const totalStaffInList = state.staffList.length;
        if (onlyStaff) {
          tsStaffStatusBadge.className = "badge badge-ok";
          let badgeHtml = `👷 ${workerCount} von ${totalStaffInList} Arbeiter erfasst`;
          if (skippedNonStaffRows > 0) {
            const resListStr = Array.from(skippedNonStaffResources).sort().join(", ");
            badgeHtml += ` • <span title="Ausgeblendete Ressourcen: ${escapeHtml(resListStr)}">🚫 ${skippedNonStaffRows} Buchung(en) Fremd-Ressourcen ausgeblendet (${skippedNonStaffResources.size} Nr.: ${escapeHtml(resListStr)})</span>`;
          }
          tsStaffStatusBadge.innerHTML = badgeHtml;
        } else {
          tsStaffStatusBadge.className = "badge badge-info";
          tsStaffStatusBadge.innerHTML = `⚠️ Alle ${sortedEmployees.length} Ressourcen aktiv (inkl. Fremd/Maschinen)`;
        }
      } else {
        tsStaffStatusBadge.className = "badge badge-leer-ok";
        tsStaffStatusBadge.innerHTML = `⚠️ Keine Mitarbeiterdatei geladen (alle Ressourcen werden angezeigt)`;
      }
    }

    if (sortedEmployees.length === 0) {
      timesheetThead.innerHTML = `<tr><th class="ts-col-emp">Mitarbeiter / Ressource</th><th>Status</th></tr>`;
      let emptyMsg = "Keine Buchungen in der ausgewählten Ressourcenspalte gefunden.";
      if (hasStaffList && onlyStaff && skippedNonStaffRows > 0) {
        const resListStr = Array.from(skippedNonStaffResources).sort().join(", ");
        emptyMsg = `Es wurden keine Arbeiter aus der Mitarbeiterdatei gefunden (${skippedNonStaffRows} Buchungen anderer Ressourcen [${resListStr}] ausgeblendet).`;
      }
      timesheetTbody.innerHTML = `<tr><td colspan="2" style="text-align: center; padding: 1.5rem; color: var(--text-muted);">${escapeHtml(emptyMsg)}</td></tr>`;
      timesheetTfoot.innerHTML = "";
      timesheetSection.classList.remove("hidden");
      return;
    }

    // THEAD
    timesheetThead.innerHTML = `
      <tr>
        <th class="ts-col-emp">Mitarbeiter / Ressource</th>
        ${sortedDates.map(d => `<th title="${escapeHtml(d.displayDate)}">${escapeHtml(d.label)}</th>`).join("")}
        <th class="ts-col-total">Gesamtstunden</th>
        <th>Buchungen</th>
      </tr>
    `;

    // TBODY
    timesheetTbody.innerHTML = sortedEmployees.map(emp => {
      const dateCells = sortedDates.map(d => {
        const h = emp.hoursByDate[d.isoKey] || 0;
        if (h > 0) {
          return `<td><span class="ts-hour-badge ts-hour-clickable" data-res="${escapeHtml(emp.resource)}" data-date="${escapeHtml(d.isoKey)}" data-display-date="${escapeHtml(d.displayDate)}" title="✏️ Klicken, um Buchung (${formatHoursDe(h)} Std) von ${escapeHtml(emp.name)} am ${escapeHtml(d.displayDate)} zu bearbeiten oder neu einzutragen">${formatHoursDe(h)}</span></td>`;
        } else {
          return `<td><span class="ts-hour-empty-clickable" data-res="${escapeHtml(emp.resource)}" data-date="${escapeHtml(d.isoKey)}" data-display-date="${escapeHtml(d.displayDate)}" title="➕ Klicken, um neue Buchung für ${escapeHtml(emp.name)} am ${escapeHtml(d.displayDate)} einzutragen">-</span></td>`;
        }
      }).join("");

      const isTargetEmp = Boolean(state.highlightEmployeeRes && (
        emp.resource === state.highlightEmployeeRes ||
        (/^\d+$/.test(emp.resource) && /^\d+$/.test(state.highlightEmployeeRes) && parseInt(emp.resource, 10) === parseInt(state.highlightEmployeeRes, 10))
      ));
      const rowClass = isTargetEmp ? "ts-row-highlight" : "";

      return `
        <tr class="${rowClass}">
          <td class="ts-cell-emp">
            <div style="display: flex; justify-content: space-between; align-items: center; gap: 0.5rem;">
              <span class="ts-emp-clickable" data-res="${escapeHtml(emp.resource)}" style="font-weight: 600; color: #1e293b;" title="🔍 Klicken, um alle Zeilen von ${escapeHtml(emp.name)} in der Gesamttabelle anzuzeigen">${escapeHtml(emp.name)}</span>
              <span class="badge ${emp.hasCustomName ? 'badge-ok' : 'badge-leer-ok'}">Nr. ${escapeHtml(emp.resource)}</span>
            </div>
          </td>
          ${dateCells}
          <td class="ts-cell-total"><span class="ts-hour-badge ts-hour-total">${formatHoursDe(emp.totalHours)} h</span></td>
          <td>${emp.totalEntries}</td>
        </tr>
      `;
    }).join("");

    if (state.highlightEmployeeRes) {
      setTimeout(() => {
        state.highlightEmployeeRes = null;
      }, 3000);
    }

    // TFOOT
    timesheetTfoot.innerHTML = `
      <tr>
        <td class="ts-cell-emp"><strong>Gesamtsumme Tag:</strong></td>
        ${sortedDates.map(d => `<td><strong>${formatHoursDe(d.totalHours)} h</strong></td>`).join("")}
        <td class="ts-cell-total"><span class="ts-grand-total">${formatHoursDe(grandTotalHours)} h</span></td>
        <td><strong>${grandTotalEntries}</strong></td>
      </tr>
    `;

    timesheetSection.classList.remove("hidden");

    // Event Listener für interaktiven Klick auf ZAHLEN (Bearbeiten / Wählen)
    timesheetTbody.querySelectorAll(".ts-hour-clickable").forEach(el => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        state.modalCaller = "timesheet";
        const res = el.dataset.res;
        const dateIso = el.dataset.date;
        const matchingRows = findRowsForResourceAndDate(res, dateIso);

        if (matchingRows.length === 1) {
          openEditRowModal(matchingRows[0].rowIdx);
        } else if (matchingRows.length > 1) {
          openChooseBookingModal(res, dateIso, matchingRows);
        } else {
          openInsertRowModal(null, { res: res, dateIso: dateIso });
        }
      });
    });

    // Event Listener für interaktiven Klick auf FREIFELDER (Neue Zeile nachtragen)
    timesheetTbody.querySelectorAll(".ts-hour-empty-clickable").forEach(el => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        state.modalCaller = "timesheet";
        const res = el.dataset.res;
        const dateIso = el.dataset.date;
        openInsertRowModal(null, { res: res, dateIso: dateIso });
      });
    });

    timesheetTbody.querySelectorAll(".ts-emp-clickable").forEach(el => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        drillDownToFullTable(el.dataset.res, null);
      });
    });
  }

  async function generateTimesheetExcel() {
    if (!state.timesheetData || state.timesheetData.employees.length === 0) {
      alert("Es liegen noch keine Daten für die Arbeitszeiten-Übersicht vor.");
      return null;
    }

    const { sortedDates, employees, grandTotalHours, grandTotalEntries } = state.timesheetData;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Arbeitszeiten_Uebersicht");

    const cols = [
      { header: "Mitarbeiter", key: "name", width: 28 },
      { header: "Ressource", key: "res", width: 14 }
    ];

    sortedDates.forEach((d, i) => {
      cols.push({
        header: d.label || d.displayDate,
        key: `d_${i}`,
        width: 14
      });
    });

    cols.push({ header: "Gesamt (Std)", key: "total", width: 16 });
    cols.push({ header: "Buchungen", key: "entries", width: 12 });

    ws.columns = cols;

    // Header styling
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell, colNum) => {
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0F766E" } // Teal header
      };
      cell.alignment = { vertical: "middle", horizontal: (colNum <= 1 ? "left" : "center"), wrapText: true };
      cell.border = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "medium", color: { argb: "FF0D5F58" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
      };
    });

    // Data rows
    employees.forEach(emp => {
      const rowValues = [emp.name, emp.resource];
      sortedDates.forEach(d => {
        const h = emp.hoursByDate[d.isoKey] || 0;
        rowValues.push(h > 0 ? h : null);
      });
      rowValues.push(emp.totalHours);
      rowValues.push(emp.totalEntries);

      const row = ws.addRow(rowValues);
      row.height = 22;

      // Spalte B (Ressource) als 0000 formatieren
      const resCell = row.getCell(2);
      resCell.numFmt = "0000";
      resCell.alignment = { horizontal: "center", vertical: "middle" };

      // Datumsspalten formatieren
      for (let c = 3; c < 3 + sortedDates.length; c++) {
        const cell = row.getCell(c);
        cell.numFmt = '#,##0.0 "h"';
        cell.alignment = { horizontal: "right", vertical: "middle" };
      }

      // Gesamtstunden
      const totalCell = row.getCell(3 + sortedDates.length);
      totalCell.numFmt = '#,##0.0 "h"';
      totalCell.font = { bold: true };
      totalCell.alignment = { horizontal: "right", vertical: "middle" };
      totalCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF0FDFA" }
      };

      // Buchungen
      const entriesCell = row.getCell(4 + sortedDates.length);
      entriesCell.numFmt = '0';
      entriesCell.alignment = { horizontal: "center", vertical: "middle" };

      row.eachCell(cell => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } }
        };
      });
    });

    // Summenzeile im Footer
    const footValues = ["Gesamtsumme Tag", ""];
    sortedDates.forEach(d => {
      footValues.push(d.totalHours);
    });
    footValues.push(grandTotalHours);
    footValues.push(grandTotalEntries);

    const footRow = ws.addRow(footValues);
    footRow.height = 26;
    footRow.eachCell((cell, colNum) => {
      cell.font = { bold: true, color: { argb: "FF0F172A" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE2E8F0" }
      };
      cell.border = {
        top: { style: "double", color: { argb: "FF475569" } },
        bottom: { style: "medium", color: { argb: "FF475569" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
      };

      if (colNum >= 3 && colNum <= 3 + sortedDates.length) {
        cell.numFmt = '#,##0.0 "h"';
        cell.alignment = { horizontal: "right", vertical: "middle" };
      } else if (colNum === 4 + sortedDates.length) {
        cell.numFmt = '0';
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      }
    });

    // Grand Total hervorheben
    const grandCell = footRow.getCell(3 + sortedDates.length);
    grandCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFCCFBF1" }
    };
    grandCell.font = { bold: true, color: { argb: "FF0F766E" }, size: 11 };

    return await wb.xlsx.writeBuffer();
  }

  if (btnExportTimesheet) {
    btnExportTimesheet.addEventListener("click", async () => {
      try {
        const buf = await generateTimesheetExcel();
        if (!buf) return;
        const baseName = (state.targetFileName || "Datei").replace(/\.[^/.]+$/, "");
        downloadBuffer(buf, `Arbeitszeiten_Uebersicht_${baseName}.xlsx`);
        showToast("✅ Arbeitszeiten-Matrix erfolgreich als Excel exportiert!");
      } catch (err) {
        console.error("Fehler beim Exportieren der Arbeitszeiten:", err);
        alert("Fehler beim Erstellen der Arbeitszeiten-Excel: " + err.message);
      }
    });
  }

  async function generateStaffTemplateExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Mitarbeiter");

    ws.columns = [
      { header: "Ressourcennummer", key: "res", width: 20 },
      { header: "Nachname", key: "nachname", width: 22 },
      { header: "Vorname", key: "vorname", width: 20 },
      { header: "Abteilung", key: "abt", width: 24 },
      { header: "Bemerkung", key: "bem", width: 28 }
    ];

    const headerRow = ws.getRow(1);
    headerRow.height = 26;
    headerRow.eachCell(cell => {
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0284C7" } // Blue header
      };
      cell.alignment = { vertical: "middle", horizontal: "left" };
      cell.border = {
        bottom: { style: "medium", color: { argb: "FF0369A1" } }
      };
    });

    const sampleData = [
      ["0045", "Mustermann", "Max", "Montage", "Leitender Monteur"],
      ["0120", "Schmidt", "Anna", "Kundendienst", "Technikerin"],
      ["0300", "Weber", "Michael", "Logistik", "Lagerleitung"],
      ["0400", "Fischer", "Sarah", "Projektleitung", "Bauleiterin"],
      ["0500", "Becker", "Thomas", "Qualitätssicherung", "Prüftechniker"],
      ["0600", "Wagner", "Julia", "Service", "Außendienst"],
      ["0700", "Hoffmann", "Stefan", "Instandhaltung", "Elektroniker"]
    ];

    sampleData.forEach(rowArr => {
      const row = ws.addRow(rowArr);
      row.height = 20;
      const resCell = row.getCell(1);
      resCell.numFmt = "0000";
      resCell.alignment = { horizontal: "center", vertical: "middle" };
    });

    return await wb.xlsx.writeBuffer();
  }

  // ==============================================================================
  // Gesamte Excel-Datei: Vollständige Tabellenansicht mit Filtern & Live-Editing
  // ==============================================================================

  function activateTab(tabName) {
    if (tabName === "fulltable") {
      if (tabBtnDiagnostics) tabBtnDiagnostics.classList.remove("active");
      if (tabBtnFulltable) tabBtnFulltable.classList.add("active");
      if (viewDiagnostics) viewDiagnostics.classList.add("hidden");
      if (viewFulltable) viewFulltable.classList.remove("hidden");
      renderFullExcelTable();
    } else {
      if (tabBtnFulltable) tabBtnFulltable.classList.remove("active");
      if (tabBtnDiagnostics) tabBtnDiagnostics.classList.add("active");
      if (viewFulltable) viewFulltable.classList.add("hidden");
      if (viewDiagnostics) viewDiagnostics.classList.remove("hidden");
    }
  }

  if (tabBtnDiagnostics) {
    tabBtnDiagnostics.addEventListener("click", () => activateTab("diagnostics"));
  }
  if (tabBtnFulltable) {
    tabBtnFulltable.addEventListener("click", () => activateTab("fulltable"));
  }

  function populateFullTableFilters() {
    if (!filterFullStaff || !filterFullDate || !state.targetWorkbook || !state.currentTargetSheet) return;

    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) return;

    const resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;
    const dateColIdx = parseInt(selectTsDateCol?.value, 10) || 1;

    const staffCounts = new Map();
    const dateCounts = new Map();

    const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0));
    for (let r = 2; r <= maxRows; r++) {
      const row = ws.getRow(r);
      if (!row) continue;
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;

      // Ressource
      const cellKeyRes = `${r}_${resColIdx}`;
      let rawRes = (state.appliedCorrections && state.appliedCorrections[cellKeyRes] !== undefined)
        ? String(state.appliedCorrections[cellKeyRes]).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
      rawRes = rawRes.replace(/[,.]0+$/, "");

      if (rawRes) {
        const key = (/^\d+$/.test(rawRes) && rawRes.length < 4) ? WebExcelEngine.padNumber(rawRes, 4) : rawRes;
        staffCounts.set(key, (staffCounts.get(key) || 0) + 1);
      }

      // Datum
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        const dObj = parseRowDate(cell.value, cell);
        if (dObj && dObj.isoKey) {
          if (!dateCounts.has(dObj.isoKey)) {
            dateCounts.set(dObj.isoKey, { label: dObj.label || dObj.displayDate, count: 0 });
          }
          dateCounts.get(dObj.isoKey).count += 1;
        }
      }
    }

    const curStaffVal = filterFullStaff.value;
    filterFullStaff.innerHTML = "";

    const hasStaffList = state.staffList && state.staffList.length > 0;
    const workerKeys = [];
    const nonWorkerKeys = [];

    Array.from(staffCounts.keys()).sort().forEach(res => {
      const isKnownWorker = Boolean(getStaffName(res));
      if (hasStaffList) {
        if (isKnownWorker) workerKeys.push(res);
        else nonWorkerKeys.push(res);
      } else {
        workerKeys.push(res);
      }
    });

    const optAll = document.createElement("option");
    optAll.value = "all";
    optAll.textContent = `Alle Ressourcen (${staffCounts.size})`;
    filterFullStaff.appendChild(optAll);

    if (hasStaffList && nonWorkerKeys.length > 0) {
      const optAllWorkers = document.createElement("option");
      optAllWorkers.value = "all_staff";
      optAllWorkers.textContent = `👷 Nur echte Arbeiter (${workerKeys.length})`;
      filterFullStaff.appendChild(optAllWorkers);

      const optAllNonWorkers = document.createElement("option");
      optAllNonWorkers.value = "all_non_staff";
      optAllNonWorkers.textContent = `⚙️ Nur Fremd-/Maschinen-Ressourcen (${nonWorkerKeys.length})`;
      filterFullStaff.appendChild(optAllNonWorkers);

      const groupWorkers = document.createElement("optgroup");
      groupWorkers.label = "👷 Arbeiter (aus Stammdaten)";
      workerKeys.forEach(res => {
        const name = getStaffName(res) || `Arbeiter ${res}`;
        const count = staffCounts.get(res);
        const opt = document.createElement("option");
        opt.value = res;
        opt.textContent = `${name} (Nr. ${res}) [${count} Zeilen]`;
        if (res === curStaffVal) opt.selected = true;
        groupWorkers.appendChild(opt);
      });
      filterFullStaff.appendChild(groupWorkers);

      const groupNonWorkers = document.createElement("optgroup");
      groupNonWorkers.label = "⚙️ Sonstige Ressourcen (Maschinen / Fremd)";
      nonWorkerKeys.forEach(res => {
        const count = staffCounts.get(res);
        const opt = document.createElement("option");
        opt.value = res;
        opt.textContent = `Ressource ${res} [${count} Zeilen] (Kein Mitarbeiter)`;
        if (res === curStaffVal) opt.selected = true;
        groupNonWorkers.appendChild(opt);
      });
      filterFullStaff.appendChild(groupNonWorkers);
    } else {
      workerKeys.forEach(res => {
        const name = getStaffName(res) || `Ressource ${res}`;
        const count = staffCounts.get(res);
        const opt = document.createElement("option");
        opt.value = res;
        opt.textContent = `${name} (Nr. ${res}) [${count} Zeilen]`;
        if (res === curStaffVal) opt.selected = true;
        filterFullStaff.appendChild(opt);
      });
    }

    if (curStaffVal && (curStaffVal === "all" || curStaffVal === "all_staff" || curStaffVal === "all_non_staff" || staffCounts.has(curStaffVal))) {
      filterFullStaff.value = curStaffVal;
    }

    const curDateVal = filterFullDate.value;
    filterFullDate.innerHTML = `<option value="all">Alle Tage (${dateCounts.size})</option>`;
    Array.from(dateCounts.keys()).sort().forEach(isoKey => {
      const dInfo = dateCounts.get(isoKey);
      const opt = document.createElement("option");
      opt.value = isoKey;
      opt.textContent = `${dInfo.label} [${dInfo.count} Zeilen]`;
      if (isoKey === curDateVal) opt.selected = true;
      filterFullDate.appendChild(opt);
    });
  }

  function renderFullExcelTable() {
    if (!fulltableThead || !fulltableTbody || !state.targetWorkbook || !state.currentTargetSheet) return;

    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) return;

    const headerRow = ws.getRow(1);
    let colCount = ws.columnCount;
    headerRow.eachCell((cell, colNum) => {
      if (colNum > colCount) colCount = colNum;
    });

    // Kopfzeile erstellen
    let theadHtml = `<tr><th class="col-sticky-row"># Zeile</th>`;
    for (let c = 1; c <= colCount; c++) {
      const headerCell = headerRow.getCell(c);
      const colName = WebExcelEngine.extractCellValue(headerCell.value, headerCell) || `Spalte ${getColLetter(c)}`;
      theadHtml += `<th><span style="color:var(--text-muted); font-size:0.75rem;">${getColLetter(c)}:</span> ${escapeHtml(colName)}</th>`;
    }
    theadHtml += `</tr>`;
    fulltableThead.innerHTML = theadHtml;

    const resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;
    const dateColIdx = parseInt(selectTsDateCol?.value, 10) || 1;

    const staffFilter = filterFullStaff ? filterFullStaff.value : "all";
    const dateFilter = filterFullDate ? filterFullDate.value : "all";
    const statusFilter = filterFullStatus ? filterFullStatus.value : "all";
    const searchQuery = filterFullSearch ? filterFullSearch.value.trim().toLowerCase() : "";

    let totalDataRows = 0;
    let matchingRowsCount = 0;
    let tbodyHtml = "";

    const resultMap = new Map();
    if (state.allResults) {
      state.allResults.forEach(r => {
        resultMap.set(`${r.row}_${r.col_idx}`, r);
      });
    }

    const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0));
    for (let r = 2; r <= maxRows; r++) {
      const row = ws.getRow(r);
      if (!row) continue;
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;
      totalDataRows++;

      // Mitarbeiter-Filter prüfen
      const cellKeyRes = `${r}_${resColIdx}`;
      let rawRes = (state.appliedCorrections && state.appliedCorrections[cellKeyRes] !== undefined)
        ? String(state.appliedCorrections[cellKeyRes]).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/\u00a0/g, " ").replace(/^['"`\s]+|['"`\s]+$/g, "").trim();
      rawRes = rawRes.replace(/[,.]0+$/, "");
      const paddedRes = (/^\d+$/.test(rawRes) && rawRes.length < 4) ? WebExcelEngine.padNumber(rawRes, 4) : rawRes;
      const unpaddedRes = rawRes.replace(/^0+/, "");
      const empName = getStaffName(rawRes);
      const isKnownWorker = Boolean(empName);

      if (staffFilter === "all_staff") {
        if (!isKnownWorker) continue;
      } else if (staffFilter === "all_non_staff") {
        if (isKnownWorker) continue;
      } else if (staffFilter !== "all") {
        const filterName = getStaffName(staffFilter);
        const matchesStaff = Boolean(empName && filterName && empName === filterName);
        const matchesKey = (rawRes === staffFilter || paddedRes === staffFilter || unpaddedRes === staffFilter);
        if (!matchesStaff && !matchesKey) {
          continue;
        }
      }

      // Datums-Filter prüfen
      let rowDateIso = "";
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        const dObj = parseRowDate(cell.value, cell);
        rowDateIso = dObj.isoKey;
      }
      if (dateFilter !== "all" && rowDateIso !== dateFilter) {
        continue;
      }

      // Status- & Such-Prüfung
      let rowHasError = false;
      let rowHasCorrection = false;
      let rowTextAcc = "";

      for (let c = 1; c <= colCount; c++) {
        const cellKey = `${r}_${c}`;
        if (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined) {
          rowHasCorrection = true;
        }
        const errObj = resultMap.get(cellKey);
        if (errObj && errObj.status !== "OK" && errObj.status !== "OK_LEER") {
          rowHasError = true;
        }
        const val = (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined)
          ? String(state.appliedCorrections[cellKey])
          : WebExcelEngine.extractCellValue(row.getCell(c).value, row.getCell(c));
        rowTextAcc += " " + String(val).toLowerCase();
      }

      const empNameForSearch = empName || "";
      if (empNameForSearch) rowTextAcc += " " + empNameForSearch.toLowerCase();

      if (statusFilter === "errors" && !rowHasError) continue;
      if (statusFilter === "corrected" && !rowHasCorrection) continue;
      if (searchQuery && !rowTextAcc.includes(searchQuery)) continue;

      matchingRowsCount++;

      let rowIndicator = "";
      if (rowHasError) rowIndicator = `<span title="Enthält Prüffehler" style="color:#ef4444; font-size:0.8rem;">⚠️</span>`;
      else if (rowHasCorrection) rowIndicator = `<span title="Manuell korrigiert" style="color:#10b981; font-size:0.8rem;">✏️</span>`;

      const isJustInserted = (state.lastInsertedRow === r);
      const rowClass = isJustInserted ? "row-just-inserted" : "";
      tbodyHtml += `<tr class="${rowClass}">`;
      tbodyHtml += `<td class="col-sticky-row">
        <span>${r}</span> ${rowIndicator}
        <button type="button" class="btn-row-add-below" data-row="${r}" title="Zeile darunter einfügen (Werte als Vorlage übernehmen)">➕</button>
      </td>`;

      for (let c = 1; c <= colCount; c++) {
        const cellKey = `${r}_${c}`;
        const cell = row.getCell(c);
        const isCorrected = (state.appliedCorrections && state.appliedCorrections[cellKey] !== undefined);
        const currentVal = isCorrected ? state.appliedCorrections[cellKey] : WebExcelEngine.extractCellValue(cell.value, cell);

        const errObj = resultMap.get(cellKey);
        const isError = errObj && errObj.status !== "OK" && errObj.status !== "OK_LEER" && !isCorrected;

        let cellClass = "full-cell-editable";
        let badgeHtml = "";

        if (isCorrected) {
          cellClass += " cell-corrected";
          badgeHtml = `<span class="cell-diag-badge badge-corr" title="Manuell korrigiert">✓</span>`;
        } else if (isError) {
          cellClass += " cell-error";
          badgeHtml = `<span class="cell-diag-badge badge-err" title="${escapeHtml(errObj.status_label)}: ${escapeHtml(errObj.detail)}">⚠️</span>`;
        }

        let displayStr = String(currentVal ?? "");
        if (typeof currentVal === "number" && !isNaN(currentVal) && !Number.isInteger(currentVal)) {
          displayStr = currentVal.toLocaleString("de-DE", { minimumFractionDigits: 1, maximumFractionDigits: 2 });
        }

        tbodyHtml += `
          <td class="${cellClass}" data-row="${r}" data-col="${c}" data-val="${escapeHtml(currentVal ?? '')}" title="Klick zum Ändern (Zelle ${getColLetter(c)}${r})">
            <div style="display: flex; justify-content: space-between; align-items: center; gap: 4px;">
              <span>${escapeHtml(displayStr)}</span>
              ${badgeHtml}
            </div>
          </td>
        `;
      }
      tbodyHtml += `</tr>`;
    }

    if (matchingRowsCount === 0) {
      tbodyHtml = `<tr><td colspan="${colCount + 1}" style="text-align: center; padding: 2rem; color: var(--text-muted);">Keine Zeilen für die aktuellen Filterkriterien gefunden. Klicken Sie auf "Filter zurücksetzen".</td></tr>`;
    }

    fulltableTbody.innerHTML = tbodyHtml;

    if (fulltableCountsText) {
      fulltableCountsText.innerHTML = `Zeige <strong>${matchingRowsCount}</strong> von <strong>${totalDataRows}</strong> Zeilen`;
    }
    if (tabBadgeTotalRows) {
      tabBadgeTotalRows.textContent = `${totalDataRows} Zeilen`;
    }

    // Inline-Editor für alle Tabellenzellen aktivieren
    fulltableTbody.querySelectorAll(".full-cell-editable").forEach(td => {
      td.addEventListener("click", () => {
        makeCellEditable(td);
      });
    });

    // Event-Listener für Zeilen-Plus-Buttons
    fulltableTbody.querySelectorAll(".btn-row-add-below").forEach(btn => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const rIdx = parseInt(btn.dataset.row, 10);
        state.modalCaller = "fulltable";
        openInsertRowModal(rIdx);
      });
    });

    if (state.lastInsertedRow) {
      const justInsertedTr = fulltableTbody.querySelector(".row-just-inserted");
      if (justInsertedTr) {
        setTimeout(() => {
          justInsertedTr.scrollIntoView({ behavior: "smooth", block: "center" });
        }, 100);
      }
      state.lastInsertedRow = null;
    }
  }

  function makeCellEditable(td) {
    if (td.classList.contains("full-cell-editing")) return;
    const rowIdx = parseInt(td.dataset.row, 10);
    const colIdx = parseInt(td.dataset.col, 10);
    const origVal = td.dataset.val ?? "";

    td.classList.add("full-cell-editing");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "cell-inline-input";
    input.value = origVal;
    td.innerHTML = "";
    td.appendChild(input);
    input.focus();
    input.select();

    let isSaved = false;
    function finishEdit() {
      if (isSaved) return;
      isSaved = true;
      const newVal = input.value.trim();
      if (newVal !== origVal) {
        saveCellCorrection(rowIdx, colIdx, newVal);
      } else {
        renderFullExcelTable();
      }
    }

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        finishEdit();
      } else if (e.key === "Escape") {
        isSaved = true;
        renderFullExcelTable();
      }
    });

    input.addEventListener("blur", () => {
      finishEdit();
    });
  }

  function saveCellCorrection(rowIdx, colIdx, newVal) {
    const cellKey = `${rowIdx}_${colIdx}`;
    state.appliedCorrections[cellKey] = newVal;

    const ws = state.targetWorkbook ? state.targetWorkbook.getWorksheet(state.currentTargetSheet) : null;
    if (ws) {
      const cell = ws.getRow(rowIdx).getCell(colIdx);
      const valStr = String(newVal).trim();
      if (colIdx === 1) {
        cell.value = valStr;
        cell.numFmt = "DD.MM.YYYY";
      } else if (/^\d+$/.test(valStr)) {
        cell.value = parseInt(valStr, 10);
      } else if (/^-?\d+\.\d+$/.test(valStr.replace(",", "."))) {
        cell.value = parseFloat(valStr.replace(",", "."));
      } else {
        cell.value = valStr;
      }
    }

    const resItem = state.allResults.find(item => item.id === cellKey);
    if (resItem) {
      resItem.current_value = newVal;
      resItem.is_corrected = (newVal !== resItem.original_value);
    }

    updateFilterCounts(state.allResults);
    renderResultsTable();

    // Arbeitszeiten-Matrix sofort neu berechnen
    renderTimesheetMatrix();

    // Gesamttabelle und Filter neu aktualisieren
    populateFullTableFilters();
    renderFullExcelTable();

    showToast(`✅ Zelle ${getColLetter(colIdx)}${rowIdx} aktualisiert auf "${newVal}". Arbeitszeiten & Excel synchronisiert!`);
  }

  function drillDownToFullTable(resource, dateIso) {
    activateTab("fulltable");
    if (filterFullStaff && resource) {
      filterFullStaff.value = resource;
    }
    if (filterFullDate && dateIso) {
      filterFullDate.value = dateIso;
    } else if (filterFullDate) {
      filterFullDate.value = "all";
    }
    if (filterFullStatus) filterFullStatus.value = "all";
    if (filterFullSearch) filterFullSearch.value = "";

    renderFullExcelTable();

    const fullTableEl = document.getElementById("view-fulltable");
    if (fullTableEl) {
      fullTableEl.scrollIntoView({ behavior: "smooth" });
    }

    const empName = getStaffName(resource) || `Ressource ${resource}`;
    showToast(`🔍 Zeige alle Zeilen für ${empName} ${dateIso ? 'am ' + dateIso : ''}. Klicken Sie auf eine Zelle zum Ändern.`);
  }

  // Filter Event Listener für Gesamttabelle
  [filterFullStaff, filterFullDate, filterFullStatus].forEach(sel => {
    if (sel) {
      sel.addEventListener("change", () => {
        renderFullExcelTable();
      });
    }
  });

  if (filterFullSearch) {
    filterFullSearch.addEventListener("input", () => {
      renderFullExcelTable();
    });
  }

  if (btnResetFullFilters) {
    btnResetFullFilters.addEventListener("click", () => {
      if (filterFullStaff) filterFullStaff.value = "all";
      if (filterFullDate) filterFullDate.value = "all";
      if (filterFullStatus) filterFullStatus.value = "all";
      if (filterFullSearch) filterFullSearch.value = "";
      renderFullExcelTable();
      showToast("Filter zurückgesetzt: Alle Zeilen werden angezeigt.");
    });
  }

  // ==============================================================================
  // Zeile nachtragen (Neu anlegen) ODER Zeile bearbeiten (Live in Prüfdatei)
  // ==============================================================================

  function formatDateIsoDe(isoStr) {
    if (!isoStr) return "";
    const parts = String(isoStr).split("-");
    if (parts.length === 3) return `${parts[2]}.${parts[1]}.${parts[0]}`;
    return String(isoStr);
  }

  function closeInsertRowModal() {
    if (modalInsertRow) modalInsertRow.classList.add("hidden");
  }

  function closeChooseBookingModal() {
    if (modalChooseBooking) modalChooseBooking.classList.add("hidden");
  }

  // Hilfsfunktion: Ressourcennummer im Dropdown oder als manuelle Nummer sicher einstellen
  function setResourceFieldValue(resColIdx, resVal) {
    if (resVal === null || resVal === undefined) return;
    const clean = String(resVal).replace(/[,.]0+$/, "").trim();
    const padded = (/^\d+$/.test(clean) && clean.length < 4) ? WebExcelEngine.padNumber(clean, 4) : clean;

    const selectEl = document.getElementById(`modal-field-${resColIdx}`);
    const customEl = document.getElementById(`modal-field-${resColIdx}-custom`);
    if (!selectEl) return;

    if (selectEl.tagName === "SELECT") {
      let matched = false;
      if (selectEl.options) {
        for (let opt of selectEl.options) {
          if (opt.value === clean || opt.value === padded) {
            selectEl.value = opt.value;
            matched = true;
            break;
          }
        }
      }
      if (matched) {
        if (customEl) customEl.style.display = "none";
      } else {
        selectEl.value = "__custom__";
        if (customEl) {
          customEl.style.display = "block";
          customEl.value = clean;
        }
      }
    } else {
      selectEl.value = clean;
    }
  }

  // Findet alle Zeilen in der Excel-Tabelle für einen bestimmten Mitarbeiter und ein Datum
  function findRowsForResourceAndDate(resource, dateIso) {
    if (!state.targetWorkbook || !state.currentTargetSheet) return [];
    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) return [];

    let dateColIdx = parseInt(selectTsDateCol?.value, 10);
    if (isNaN(dateColIdx) || dateColIdx < 1) dateColIdx = 1;

    let resColIdx = parseInt(selectTsResourceCol?.value, 10);
    if (isNaN(resColIdx) || resColIdx < 1) resColIdx = 3;

    let hoursColIdx = parseInt(selectTsHoursCol?.value, 10);
    if (isNaN(hoursColIdx) || hoursColIdx < 1) hoursColIdx = null;

    let colCount = 0;
    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell, colNum) => {
      if (colNum > colCount) colCount = colNum;
    });
    if (colCount === 0 && ws.columnCount) colCount = ws.columnCount;
    if (colCount === 0 && ws.actualColumnCount) colCount = ws.actualColumnCount;
    if (colCount === 0) colCount = 10;

    const cleanRes = String(resource || "").replace(/[,.]0+$/, "").trim();
    const cleanNum = /^\d+$/.test(cleanRes) ? parseInt(cleanRes, 10) : null;

    const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0));
    const matches = [];

    for (let r = 2; r <= maxRows; r++) {
      const row = ws.getRow(r);
      if (!row) continue;
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;

      // 1. Ressource prüfen
      const cellKeyRes = `${r}_${resColIdx}`;
      let rRes = (state.appliedCorrections && state.appliedCorrections[cellKeyRes] !== undefined)
        ? String(state.appliedCorrections[cellKeyRes]).trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).trim();
      rRes = rRes.replace(/[,.]0+$/, "").trim();

      let resMatches = (rRes === cleanRes);
      if (!resMatches && cleanNum !== null && /^\d+$/.test(rRes)) {
        resMatches = (parseInt(rRes, 10) === cleanNum);
      }
      if (!resMatches) continue;

      // 2. Datum prüfen
      let rDateIso = "";
      let rDateDisp = "";
      if (dateColIdx) {
        const cell = row.getCell(dateColIdx);
        const cellKeyDate = `${r}_${dateColIdx}`;
        const corrVal = state.appliedCorrections ? state.appliedCorrections[cellKeyDate] : undefined;
        if (corrVal !== undefined) {
          const pd = WebExcelEngine.parseDateValue(corrVal);
          if (pd) {
            rDateIso = pd.toISOString().split("T")[0];
            rDateDisp = WebExcelEngine.formatDate(pd);
          } else {
            rDateIso = String(corrVal);
            rDateDisp = String(corrVal);
          }
        } else {
          const dObj = parseRowDate(cell.value, cell);
          if (dObj) {
            rDateIso = dObj.isoKey;
            rDateDisp = dObj.displayDate;
          }
        }
      }

      if (dateIso && rDateIso !== dateIso) continue;

      // 3. Stunden
      let rHours = 0;
      if (hoursColIdx) {
        const cellKeyHours = `${r}_${hoursColIdx}`;
        const hVal = (state.appliedCorrections && state.appliedCorrections[cellKeyHours] !== undefined)
          ? state.appliedCorrections[cellKeyHours]
          : row.getCell(hoursColIdx).value;
        rHours = parseHours(hVal);
      } else {
        rHours = 1;
      }

      // 4. Weitere Infos (z. B. Auftrag, Tätigkeit)
      let infoParts = [];
      for (let c = 1; c <= colCount; c++) {
        if (c !== dateColIdx && c !== resColIdx && c !== hoursColIdx) {
          const cell = row.getCell(c);
          const cellKeyOther = `${r}_${c}`;
          const val = (state.appliedCorrections && state.appliedCorrections[cellKeyOther] !== undefined)
            ? String(state.appliedCorrections[cellKeyOther]).trim()
            : WebExcelEngine.extractCellValue(cell.value, cell).trim();
          if (val) {
            const hCell = headerRow.getCell(c);
            const hName = WebExcelEngine.extractCellValue(hCell.value, hCell) || `Spalte ${getColLetter(c)}`;
            infoParts.push(`${hName}: ${val}`);
            if (infoParts.length >= 2) break;
          }
        }
      }

      matches.push({
        rowIdx: r,
        resource: rRes,
        dateIso: rDateIso,
        dateDisp: rDateDisp || dateIso,
        hours: rHours,
        info: infoParts.join(" • ")
      });
    }

    return matches;
  }

  // Öffnet den Auswahldialog, wenn für einen Tag mehrere Buchungen existieren
  function openChooseBookingModal(resource, dateIso, matches) {
    if (!modalChooseBooking || !chooseBookingList) return;

    const empName = getStaffName(resource) || `Ressource ${resource}`;
    const dDisp = (matches[0] && matches[0].dateDisp) ? matches[0].dateDisp : (dateIso ? formatDateIsoDe(dateIso) : "");

    if (chooseBookingTitle) {
      chooseBookingTitle.textContent = `Buchungen: ${empName}`;
    }
    if (chooseBookingSubtitle) {
      chooseBookingSubtitle.textContent = `Am ${dDisp} liegen ${matches.length} Buchungen vor:`;
    }

    let html = "";
    matches.forEach(m => {
      html += `
        <div class="choose-booking-item" style="display: flex; justify-content: space-between; align-items: center; background: #f8fafc; border: 1px solid var(--border); border-radius: 6px; padding: 0.65rem 0.85rem; gap: 0.75rem;">
          <div>
            <div style="font-weight: 600; color: #0f172a; font-size: 0.92rem;">
              Zeile ${m.rowIdx}: <span style="color: #0369a1;">${formatHoursDe(m.hours)} Stunden</span>
            </div>
            ${m.info ? `<div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 3px;">${escapeHtml(m.info)}</div>` : ''}
          </div>
          <button type="button" class="btn btn-primary btn-sm btn-edit-specific-row" data-row="${m.rowIdx}" style="padding: 4px 12px; font-size: 0.82rem; white-space: nowrap;">
            ✏️ Bearbeiten
          </button>
        </div>
      `;
    });
    chooseBookingList.innerHTML = html;

    chooseBookingList.querySelectorAll(".btn-edit-specific-row").forEach(btn => {
      btn.addEventListener("click", () => {
        const rIdx = parseInt(btn.dataset.row, 10);
        state.modalCaller = "timesheet";
        closeChooseBookingModal();
        openEditRowModal(rIdx);
      });
    });

    if (btnChooseAddNew) {
      btnChooseAddNew.onclick = () => {
        state.modalCaller = "timesheet";
        closeChooseBookingModal();
        openInsertRowModal(null, { res: resource, dateIso: dateIso });
      };
    }

    modalChooseBooking.classList.remove("hidden");
  }

  // Generiert die dynamischen Input-Felder im Modal für jede Spalte
  function renderModalFields(tgtHeaders, dateColIdx, resColIdx, hoursColIdx) {
    if (!insertFieldsGrid) return;
    let fieldsHtml = "";

    tgtHeaders.forEach(h => {
      const c = h.colNum;
      const isDate = (c === dateColIdx);
      const isRes = (c === resColIdx && c !== dateColIdx);
      const isHours = (c === hoursColIdx && c !== dateColIdx && c !== resColIdx);

      fieldsHtml += `<div class="field-item">`;
      fieldsHtml += `<label for="modal-field-${c}"><span>${escapeHtml(h.name)}</span> <span class="field-col-tag">Spalte ${h.letter}</span></label>`;

      if (isDate) {
        fieldsHtml += `
          <div style="display: flex; gap: 0.35rem; align-items: center;">
            <input type="date" class="form-input form-input-sm insert-field-val" data-col="${c}" id="modal-field-${c}" style="flex: 1;">
            <button type="button" class="btn btn-secondary btn-sm btn-quick-today" style="padding: 2px 7px; font-size: 0.75rem;" title="Heutiges Datum setzen">Heute</button>
          </div>
        `;
      } else if (isRes) {
        let staffOpts = `<option value="">-- Bitte wählen --</option>`;
        if (state.staffList && state.staffList.length > 0) {
          state.staffList.forEach(s => {
            staffOpts += `<option value="${escapeHtml(s.resource)}">${escapeHtml(s.resource)} - ${escapeHtml(s.name)}</option>`;
          });
        }
        staffOpts += `<option value="__custom__">✏️ Andere Nummer manuell eingeben...</option>`;

        fieldsHtml += `
          <div>
            <select class="form-select form-select-sm insert-field-val" data-col="${c}" id="modal-field-${c}">
              ${staffOpts}
            </select>
            <input type="text" class="form-input form-input-sm insert-custom-res" id="modal-field-${c}-custom" placeholder="Nummer eingeben..." style="display: none; margin-top: 4px;">
          </div>
        `;
      } else if (isHours) {
        fieldsHtml += `
          <input type="number" step="0.25" min="0" max="24" class="form-input form-input-sm insert-field-val" data-col="${c}" id="modal-field-${c}" placeholder="z. B. 8.0">
        `;
      } else {
        fieldsHtml += `
          <input type="text" class="form-input form-input-sm insert-field-val" data-col="${c}" id="modal-field-${c}">
        `;
      }
      fieldsHtml += `</div>`;
    });
    insertFieldsGrid.innerHTML = fieldsHtml;

    // Event listener für Quick "Heute" Button
    insertFieldsGrid.querySelectorAll(".btn-quick-today").forEach(btn => {
      btn.addEventListener("click", () => {
        const inp = btn.previousElementSibling;
        if (inp && inp.type === "date") {
          const now = new Date();
          inp.value = now.toISOString().split("T")[0];
        }
      });
    });

    // Event listener für Manuelle Nummer bei Mitarbeiter
    const resSelect = document.getElementById(`modal-field-${resColIdx}`);
    const resCustom = document.getElementById(`modal-field-${resColIdx}-custom`);
    if (resSelect && resCustom) {
      resSelect.addEventListener("change", () => {
        if (resSelect.value === "__custom__") {
          resCustom.style.display = "block";
          resCustom.focus();
        } else {
          resCustom.style.display = "none";
        }
      });
    }
  }

  // Füllt die Formularfelder mit den Werten einer bestimmten Zeile
  function applyRowValuesToFields(rowIdx, tgtHeaders, dateColIdx, resColIdx, hoursColIdx) {
    if (!rowIdx || !state.targetWorkbook || !state.currentTargetSheet) return;
    const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
    if (!ws) return;
    const tRow = ws.getRow(rowIdx);
    if (!tRow) return;

    tgtHeaders.forEach(h => {
      const c = h.colNum;
      const inputEl = document.getElementById(`modal-field-${c}`);
      if (!inputEl) return;

      const cell = tRow.getCell(c);
      const cellKey = `${rowIdx}_${c}`;
      const corrVal = state.appliedCorrections ? state.appliedCorrections[cellKey] : undefined;

      if (c === dateColIdx) {
        if (corrVal !== undefined) {
          const pd = WebExcelEngine.parseDateValue(corrVal);
          if (pd) inputEl.value = pd.toISOString().split("T")[0];
          else inputEl.value = String(corrVal);
        } else {
          const dObj = parseRowDate(cell.value, cell);
          if (dObj && dObj.isoKey && dObj.isoKey !== "ohne_datum") {
            inputEl.value = dObj.isoKey;
          } else {
            const raw = WebExcelEngine.extractCellValue(cell.value, cell);
            const pd = WebExcelEngine.parseDateValue(raw);
            if (pd) inputEl.value = pd.toISOString().split("T")[0];
            else inputEl.value = "";
          }
        }
      } else if (c === resColIdx) {
        const raw = (corrVal !== undefined)
          ? String(corrVal).trim()
          : WebExcelEngine.extractCellValue(cell.value, cell).replace(/[,.]0+$/, "").trim();
        setResourceFieldValue(c, raw);
      } else {
        const val = (corrVal !== undefined)
          ? corrVal
          : WebExcelEngine.extractCellValue(cell.value, cell);
        inputEl.value = (val !== null && val !== undefined) ? String(val) : "";
      }
    });
  }

  // Öffnet das Modal im BEARBEITEN-Modus für eine bestehende Zeile
  function openEditRowModal(rowIdx) {
    try {
      if (!state.targetWorkbook || !state.currentTargetSheet) {
        alert("Bitte laden Sie zuerst eine zu prüfende Excel-Datei.");
        return;
      }
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      if (!ws) {
        alert("Das ausgewählte Arbeitsblatt konnte nicht gefunden werden.");
        return;
      }

      const row = ws.getRow(rowIdx);
      if (!row) {
        alert(`Die Zeile ${rowIdx} existiert nicht.`);
        return;
      }

      state.modalRowMode = "edit";
      state.modalEditRowIdx = rowIdx;

      let colCount = 0;
      const headerRow = ws.getRow(1);
      headerRow.eachCell((cell, colNum) => {
        if (colNum > colCount) colCount = colNum;
      });
      if (colCount === 0 && ws.columnCount) colCount = ws.columnCount;
      if (colCount === 0 && ws.actualColumnCount) colCount = ws.actualColumnCount;
      if (colCount === 0) colCount = 10;

      const tgtHeaders = [];
      for (let c = 1; c <= colCount; c++) {
        const headerCell = headerRow.getCell(c);
        const colName = WebExcelEngine.extractCellValue(headerCell.value, headerCell) || `Spalte ${getColLetter(c)}`;
        tgtHeaders.push({
          colNum: c,
          letter: getColLetter(c),
          name: colName
        });
      }

      let dateColIdx = parseInt(selectTsDateCol?.value, 10);
      if (isNaN(dateColIdx) || dateColIdx < 1) {
        const foundDate = tgtHeaders.find(h => /datum|date|tag|zeitpunkt|buchung/i.test(h.name));
        dateColIdx = foundDate ? foundDate.colNum : 1;
      }

      let resColIdx = parseInt(selectTsResourceCol?.value, 10);
      if (isNaN(resColIdx) || resColIdx < 1) {
        const foundRes = tgtHeaders.find(h => /ressource|mitarbeiter|pers|ma|worker/i.test(h.name));
        resColIdx = foundRes ? foundRes.colNum : (colCount >= 3 ? 3 : 1);
      }

      let hoursColIdx = parseInt(selectTsHoursCol?.value, 10);
      if (isNaN(hoursColIdx) || hoursColIdx < 1) {
        const foundHours = tgtHeaders.find(h => /stunde|dauer|zeit|menge|std|^h$/i.test(h.name));
        hoursColIdx = foundHours ? foundHours.colNum : null;
      }

      // UI auf Edit-Modus anpassen
      if (modalIcon) modalIcon.textContent = "✏️";
      if (modalInsertOnlySections) modalInsertOnlySections.classList.add("hidden");
      if (modalEditBanner) modalEditBanner.classList.remove("hidden");

      const cellKeyRes = `${rowIdx}_${resColIdx}`;
      const curRes = (state.appliedCorrections && state.appliedCorrections[cellKeyRes] !== undefined)
        ? String(state.appliedCorrections[cellKeyRes]).trim()
        : WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).trim();
      const empName = getStaffName(curRes) || `Ressource ${curRes}`;

      const cellKeyDate = `${rowIdx}_${dateColIdx}`;
      let curDateIso = "";
      let curDateDisp = "";
      const corrDate = state.appliedCorrections ? state.appliedCorrections[cellKeyDate] : undefined;
      if (corrDate !== undefined) {
        const pd = WebExcelEngine.parseDateValue(corrDate);
        if (pd) {
          curDateIso = pd.toISOString().split("T")[0];
          curDateDisp = WebExcelEngine.formatDate(pd);
        } else {
          curDateIso = String(corrDate);
          curDateDisp = String(corrDate);
        }
      } else {
        const dObj = parseRowDate(row.getCell(dateColIdx).value, row.getCell(dateColIdx));
        if (dObj) {
          curDateIso = dObj.isoKey;
          curDateDisp = dObj.displayDate;
        }
      }

      if (modalTitle) modalTitle.textContent = `Zeile ${rowIdx} bearbeiten: ${empName}`;
      if (modalDesc) modalDesc.textContent = `Ändern Sie die Werte dieser Zeile (${curDateDisp || 'Datum unbestimmt'}). Die Daten werden direkt in Excel übernommen.`;
      if (modalEditRowNum) modalEditRowNum.textContent = `Zeile ${rowIdx} (${empName})`;
      if (modalFieldsTitle) modalFieldsTitle.textContent = `Werte für Zeile ${rowIdx} anpassen`;
      if (btnConfirmIcon) btnConfirmIcon.textContent = "💾";
      if (btnConfirmInsertText) btnConfirmInsertText.textContent = `Änderungen für Zeile ${rowIdx} speichern`;

      if (btnSwitchToInsert) {
        btnSwitchToInsert.onclick = () => {
          state.modalCaller = "timesheet";
          openInsertRowModal(null, { res: curRes, dateIso: curDateIso });
        };
      }

      // Felder rendern & Werte belegen
      renderModalFields(tgtHeaders, dateColIdx, resColIdx, hoursColIdx);
      applyRowValuesToFields(rowIdx, tgtHeaders, dateColIdx, resColIdx, hoursColIdx);

      if (modalInsertRow) modalInsertRow.classList.remove("hidden");

      setTimeout(() => {
        if (hoursColIdx) {
          const hInput = document.getElementById(`modal-field-${hoursColIdx}`);
          if (hInput) {
            hInput.focus();
            hInput.select();
            return;
          }
        }
        const firstInput = insertFieldsGrid ? insertFieldsGrid.querySelector("input, select") : null;
        if (firstInput) firstInput.focus();
      }, 50);

    } catch (err) {
      console.error("Fehler beim Öffnen des Bearbeiten-Dialogs:", err);
      alert("Fehler beim Öffnen des Bearbeiten-Dialogs: " + err.message);
    }
  }

  // Öffnet das Modal im NEU-ANLEGEN-Modus (Zeile nachtragen)
  function openInsertRowModal(sourceRowIdx = null, prefill = null) {
    try {
      if (!state.targetWorkbook || !state.currentTargetSheet) {
        alert("Bitte laden Sie zuerst eine zu prüfende Excel-Datei.");
        return;
      }
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      if (!ws) {
        alert("Das ausgewählte Arbeitsblatt konnte nicht gefunden werden.");
        return;
      }

      state.modalRowMode = "insert";
      state.modalEditRowIdx = null;

      let colCount = 0;
      const headerRow = ws.getRow(1);
      headerRow.eachCell((cell, colNum) => {
        if (colNum > colCount) colCount = colNum;
      });
      if (colCount === 0 && ws.columnCount) colCount = ws.columnCount;
      if (colCount === 0 && ws.actualColumnCount) colCount = ws.actualColumnCount;
      if (colCount === 0) colCount = 10;

      const tgtHeaders = [];
      for (let c = 1; c <= colCount; c++) {
        const headerCell = headerRow.getCell(c);
        const colName = WebExcelEngine.extractCellValue(headerCell.value, headerCell) || `Spalte ${getColLetter(c)}`;
        tgtHeaders.push({
          colNum: c,
          letter: getColLetter(c),
          name: colName
        });
      }

      let dateColIdx = parseInt(selectTsDateCol?.value, 10);
      if (isNaN(dateColIdx) || dateColIdx < 1) {
        const foundDate = tgtHeaders.find(h => /datum|date|tag|zeitpunkt|buchung/i.test(h.name));
        dateColIdx = foundDate ? foundDate.colNum : 1;
      }

      let resColIdx = parseInt(selectTsResourceCol?.value, 10);
      if (isNaN(resColIdx) || resColIdx < 1) {
        const foundRes = tgtHeaders.find(h => /ressource|mitarbeiter|pers|ma|worker/i.test(h.name));
        resColIdx = foundRes ? foundRes.colNum : (colCount >= 3 ? 3 : 1);
      }

      let hoursColIdx = parseInt(selectTsHoursCol?.value, 10);
      if (isNaN(hoursColIdx) || hoursColIdx < 1) {
        const foundHours = tgtHeaders.find(h => /stunde|dauer|zeit|menge|std|^h$/i.test(h.name));
        hoursColIdx = foundHours ? foundHours.colNum : null;
      }

      // UI auf Insert-Modus setzen
      if (modalIcon) modalIcon.textContent = "➕";
      if (modalInsertOnlySections) modalInsertOnlySections.classList.remove("hidden");
      if (modalEditBanner) modalEditBanner.classList.add("hidden");
      if (modalFieldsTitle) modalFieldsTitle.textContent = "Werte der neuen Zeile eingeben";
      if (btnConfirmIcon) btnConfirmIcon.textContent = "➕";
      if (btnConfirmInsertText) btnConfirmInsertText.textContent = "Zeile jetzt einfügen";

      // Titel dynamisch anpassen
      if (prefill && prefill.res) {
        const empName = getStaffName(prefill.res) || `Ressource ${prefill.res}`;
        const dDisp = prefill.dateIso ? formatDateIsoDe(prefill.dateIso) : "";
        if (modalTitle) modalTitle.textContent = `Zeile nachtragen für ${empName}${dDisp ? ' am ' + dDisp : ''}`;
        if (modalDesc) modalDesc.textContent = `Fügt eine neue Zeile für ${empName} mit korrekter Formatierung an die richtige Stelle der Datei ein.`;
      } else {
        if (modalTitle) modalTitle.textContent = "Zeile in Prüfdatei nachtragen";
        if (modalDesc) modalDesc.textContent = "Fügt eine fehlende Zeile mit korrekter Formatierung an die richtige Stelle der Datei ein.";
      }

      // 1. Dropdowns für Zeilenauswahl & Vorlage aufbauen
      const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0));
      const rowOptions = [];

      for (let r = 2; r <= maxRows; r++) {
        const row = ws.getRow(r);
        if (!row) continue;
        let hasData = false;
        row.eachCell(() => { hasData = true; });
        if (!hasData) continue;

        let dateDisp = "";
        if (dateColIdx) {
          const cell = row.getCell(dateColIdx);
          const dObj = parseRowDate(cell.value, cell);
          dateDisp = dObj ? dObj.displayDate : WebExcelEngine.extractCellValue(cell.value, cell);
        }
        let resVal = WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/[,.]0+$/, "").trim();
        const staffName = getStaffName(resVal);
        const resDisp = staffName ? `${resVal} (${staffName})` : (resVal ? `Nr. ${resVal}` : "—");

        let otherColVal = "";
        for (let c = 1; c <= colCount; c++) {
          if (c !== dateColIdx && c !== resColIdx && c !== hoursColIdx) {
            const v = WebExcelEngine.extractCellValue(row.getCell(c).value, row.getCell(c)).trim();
            if (v) {
              otherColVal = v;
              break;
            }
          }
        }

        const label = `Zeile ${r}: ${dateDisp ? '[' + dateDisp + '] ' : ''}MA: ${resDisp}${otherColVal ? ' | ' + otherColVal : ''}`;
        rowOptions.push({ rowIdx: r, label: label });
      }

      if (insertTargetRowSelect) {
        insertTargetRowSelect.innerHTML = rowOptions.map(opt => `<option value="${opt.rowIdx}">${escapeHtml(opt.label)}</option>`).join("");
        if (sourceRowIdx && rowOptions.some(o => o.rowIdx === sourceRowIdx)) {
          insertTargetRowSelect.value = sourceRowIdx;
        } else if (rowOptions.length > 0) {
          insertTargetRowSelect.value = rowOptions[rowOptions.length - 1].rowIdx;
        }
      }

      if (insertTemplateRowSelect) {
        insertTemplateRowSelect.innerHTML = `<option value="">-- Keine Vorlage (neu ausfüllen) --</option>` +
          rowOptions.map(opt => `<option value="${opt.rowIdx}">${escapeHtml(opt.label)}</option>`).join("");
        if (sourceRowIdx && rowOptions.some(o => o.rowIdx === sourceRowIdx)) {
          insertTemplateRowSelect.value = sourceRowIdx;
        } else {
          insertTemplateRowSelect.value = "";
        }
      }

      // 2. Eingabefelder dynamisch rendern
      renderModalFields(tgtHeaders, dateColIdx, resColIdx, hoursColIdx);

      if (insertTemplateRowSelect) {
        insertTemplateRowSelect.onchange = (e) => {
          const selVal = parseInt(e.target.value, 10);
          if (selVal) applyRowValuesToFields(selVal, tgtHeaders, dateColIdx, resColIdx, hoursColIdx);
        };
      }

      // 3. Modus & Initialwerte setzen
      const radioAuto = document.querySelector('input[name="insert-pos-mode"][value="auto"]');
      const radioAfter = document.querySelector('input[name="insert-pos-mode"][value="after_row"]');

      if (sourceRowIdx) {
        if (radioAfter) radioAfter.checked = true;
        if (insertAfterRowContainer) insertAfterRowContainer.classList.remove("hidden");
        if (insertAutoHint) insertAutoHint.classList.add("hidden");
        applyRowValuesToFields(sourceRowIdx, tgtHeaders, dateColIdx, resColIdx, hoursColIdx);
      } else {
        if (radioAuto) radioAuto.checked = true;
        if (insertAfterRowContainer) insertAfterRowContainer.classList.add("hidden");
        if (insertAutoHint) insertAutoHint.classList.remove("hidden");

        if (prefill) {
          if (prefill.dateIso) {
            const dateInput = document.getElementById(`modal-field-${dateColIdx}`);
            if (dateInput) dateInput.value = prefill.dateIso;
          }
          if (prefill.res) {
            setResourceFieldValue(resColIdx, prefill.res);
          }
          if (prefill.hours !== undefined && hoursColIdx) {
            const hInput = document.getElementById(`modal-field-${hoursColIdx}`);
            if (hInput) hInput.value = prefill.hours;
          }
        } else {
          if (filterFullDate && filterFullDate.value && filterFullDate.value !== "all") {
            const dateInput = document.getElementById(`modal-field-${dateColIdx}`);
            if (dateInput) dateInput.value = filterFullDate.value;
          }
          if (filterFullStaff && filterFullStaff.value && filterFullStaff.value !== "all" && filterFullStaff.value !== "all_staff" && filterFullStaff.value !== "all_non_staff") {
            setResourceFieldValue(resColIdx, filterFullStaff.value);
          }
        }
      }

      if (modalInsertRow) modalInsertRow.classList.remove("hidden");

      setTimeout(() => {
        if (prefill && prefill.res && prefill.dateIso && hoursColIdx) {
          const hInput = document.getElementById(`modal-field-${hoursColIdx}`);
          if (hInput) {
            hInput.focus();
            return;
          }
        }
        const firstInp = insertFieldsGrid ? insertFieldsGrid.querySelector("input, select") : null;
        if (firstInp) firstInp.focus();
      }, 50);

    } catch (err) {
      console.error("Fehler beim Öffnen des Zeile-nachtragen-Dialogs:", err);
      alert("Fehler beim Öffnen des Zeile-nachtragen-Dialogs: " + err.message);
    }
  }

  // Speichert die Änderungen einer bearbeiteten Zeile direkt in die Prüfdatei
  function executeSaveEditedRow(rowIdx) {
    try {
      if (!state.targetWorkbook || !state.currentTargetSheet) return;
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      if (!ws) return;

      const row = ws.getRow(rowIdx);
      if (!row) return;

      let dateColIdx = parseInt(selectTsDateCol?.value, 10);
      if (isNaN(dateColIdx) || dateColIdx < 1) dateColIdx = 1;

      let resColIdx = parseInt(selectTsResourceCol?.value, 10);
      if (isNaN(resColIdx) || resColIdx < 1) resColIdx = 3;

      let hoursColIdx = parseInt(selectTsHoursCol?.value, 10);
      if (isNaN(hoursColIdx) || hoursColIdx < 1) hoursColIdx = null;

      let colCount = 0;
      const headerRow = ws.getRow(1);
      headerRow.eachCell((cell, colNum) => {
        if (colNum > colCount) colCount = colNum;
      });
      if (colCount === 0 && ws.columnCount) colCount = ws.columnCount;
      if (colCount === 0 && ws.actualColumnCount) colCount = ws.actualColumnCount;
      if (colCount === 0) colCount = 10;

      let savedRes = "";

      for (let c = 1; c <= colCount; c++) {
        const inputEl = document.getElementById(`modal-field-${c}`);
        let val = inputEl ? inputEl.value.trim() : "";

        if (c === resColIdx && val === "__custom__") {
          const customEl = document.getElementById(`modal-field-${c}-custom`);
          val = customEl ? customEl.value.trim() : "";
        }

        const cell = row.getCell(c);
        const cellKey = `${rowIdx}_${c}`;

        if (c === dateColIdx) {
          if (val) {
            const parts = val.split("-");
            let dispDate = val;
            if (parts.length === 3) {
              dispDate = `${parts[2]}.${parts[1]}.${parts[0]}`;
            }
            cell.value = dispDate;
            cell.numFmt = "DD.MM.YYYY";
            cell.alignment = { horizontal: "center", vertical: "middle" };
            state.appliedCorrections[cellKey] = dispDate;
          } else {
            cell.value = "";
            state.appliedCorrections[cellKey] = "";
          }
        } else if (c === resColIdx) {
          let cleanRes = val.replace(/[,.]0+$/, "").trim();
          if (/^\d{1,4}$/.test(cleanRes)) {
            cleanRes = WebExcelEngine.padNumber(cleanRes, 4);
            cell.numFmt = "0000";
          }
          cell.value = cleanRes;
          cell.alignment = { horizontal: "center", vertical: "middle" };
          state.appliedCorrections[cellKey] = cleanRes;
          savedRes = cleanRes;
        } else if (c === hoursColIdx) {
          if (val !== "") {
            const num = parseFloat(val.replace(",", "."));
            if (!isNaN(num)) {
              cell.value = num;
              cell.numFmt = "#,##0.0";
              cell.alignment = { horizontal: "right", vertical: "middle" };
              state.appliedCorrections[cellKey] = num;
            } else {
              cell.value = val;
              state.appliedCorrections[cellKey] = val;
            }
          } else {
            cell.value = null;
            state.appliedCorrections[cellKey] = "";
          }
        } else {
          if (/^\d+$/.test(val) && val.length < 10) {
            cell.value = parseInt(val, 10);
          } else {
            cell.value = val;
          }
          state.appliedCorrections[cellKey] = val;
        }
      }

      closeInsertRowModal();

      // Mitarbeiter für sanftes Hervorheben in der Wochentabelle merken
      if (savedRes) state.highlightEmployeeRes = savedRes;

      // UI & Analyse aktualisieren (OHNE Scrollen nach oben und ohne Toast-Überschreibung!)
      runInspection({ skipScroll: true, skipToast: true });

      showToast(`✅ Änderungen für Zeile ${rowIdx} erfolgreich in Prüfdatei gespeichert!`);

      // Wenn aus der Wochentabelle aufgerufen: Wochentabelle direkt wieder im Blick behalten
      if (state.modalCaller === "timesheet") {
        setTimeout(() => {
          const tsSec = document.getElementById("timesheet-section");
          if (tsSec) {
            tsSec.scrollIntoView({ behavior: "smooth", block: "start" });
          }
        }, 50);
      }
    } catch (err) {
      console.error("Fehler beim Speichern der bearbeiteten Zeile:", err);
      alert("Fehler beim Speichern der Zeile: " + err.message);
    }
  }

  // Fügt eine neue Zeile an die gewünschte Position ein
  function executeInsertRow() {
    try {
      if (!state.targetWorkbook || !state.currentTargetSheet) return;
      const ws = state.targetWorkbook.getWorksheet(state.currentTargetSheet);
      if (!ws) return;

      let colCount = 0;
      const headerRow = ws.getRow(1);
      headerRow.eachCell((cell, colNum) => {
        if (colNum > colCount) colCount = colNum;
      });
      if (colCount === 0 && ws.columnCount) colCount = ws.columnCount;
      if (colCount === 0 && ws.actualColumnCount) colCount = ws.actualColumnCount;
      if (colCount === 0) colCount = 10;

      let dateColIdx = parseInt(selectTsDateCol?.value, 10) || 1;
      let resColIdx = parseInt(selectTsResourceCol?.value, 10) || 3;
      let hoursColIdx = parseInt(selectTsHoursCol?.value, 10) || null;

      // 1. Spaltenwerte einsammeln
      const rowValues = [];
      let enteredDateIso = "";
      let enteredRes = "";

      for (let c = 1; c <= colCount; c++) {
        const inputEl = document.getElementById(`modal-field-${c}`);
        let val = inputEl ? inputEl.value.trim() : "";

        if (c === resColIdx && val === "__custom__") {
          const customEl = document.getElementById(`modal-field-${c}-custom`);
          val = customEl ? customEl.value.trim() : "";
        }

        if (c === dateColIdx) {
          if (val) {
            enteredDateIso = val;
            const parts = val.split("-");
            if (parts.length === 3) {
              val = `${parts[2]}.${parts[1]}.${parts[0]}`;
            }
          }
        } else if (c === resColIdx) {
          enteredRes = val.replace(/[,.]0+$/, "").trim();
          if (/^\d{1,4}$/.test(enteredRes)) {
            val = WebExcelEngine.padNumber(enteredRes, 4);
          }
        } else if (c === hoursColIdx && val) {
          const parsedH = parseFloat(val.replace(",", "."));
          if (!isNaN(parsedH)) val = parsedH;
        }

        rowValues.push(val !== "" ? val : null);
      }

      // 2. Zielposition ermitteln
      const mode = document.querySelector('input[name="insert-pos-mode"]:checked')?.value || "auto";
      const maxRows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, (ws._rows ? ws._rows.length - 1 : 0));
      let insertAtRowIndex = maxRows + 1;

      if (mode === "end") {
        insertAtRowIndex = maxRows + 1;
      } else if (mode === "after_row") {
        const selectedRow = parseInt(insertTargetRowSelect?.value, 10);
        if (selectedRow && selectedRow >= 1 && selectedRow <= maxRows) {
          insertAtRowIndex = selectedRow + 1;
        } else {
          insertAtRowIndex = maxRows + 1;
        }
      } else {
        // Auto-Modus: Chronologisch nach Datum & Mitarbeiter
        let foundDateMatchLast = null;
        let foundStaffAndDateMatchLast = null;
        let foundFirstLaterDate = null;

        for (let r = 2; r <= maxRows; r++) {
          const row = ws.getRow(r);
          if (!row) continue;

          let rowDateIso = "";
          if (dateColIdx) {
            const dObj = parseRowDate(row.getCell(dateColIdx).value, row.getCell(dateColIdx));
            if (dObj && dObj.isoKey) rowDateIso = dObj.isoKey;
          }

          let rowRes = "";
          if (resColIdx) {
            rowRes = WebExcelEngine.extractCellValue(row.getCell(resColIdx).value, row.getCell(resColIdx)).replace(/[,.]0+$/, "").trim();
          }

          if (enteredDateIso && rowDateIso) {
            if (rowDateIso === enteredDateIso) {
              foundDateMatchLast = r;
              if (enteredRes && (rowRes === enteredRes || (/^\d+$/.test(rowRes) && parseInt(rowRes, 10) === parseInt(enteredRes, 10)))) {
                foundStaffAndDateMatchLast = r;
              }
            } else if (rowDateIso > enteredDateIso && !foundFirstLaterDate) {
              foundFirstLaterDate = r;
            }
          }
        }

        if (foundStaffAndDateMatchLast) {
          insertAtRowIndex = foundStaffAndDateMatchLast + 1;
        } else if (foundDateMatchLast) {
          insertAtRowIndex = foundDateMatchLast + 1;
        } else if (foundFirstLaterDate) {
          insertAtRowIndex = foundFirstLaterDate;
        } else {
          insertAtRowIndex = maxRows + 1;
        }
      }

      // 3. Zeile in Worksheet einfügen
      const insertedRow = ws.insertRow(insertAtRowIndex, rowValues);

      // Formate und Ausrichtung setzen
      for (let c = 1; c <= colCount; c++) {
        const cell = insertedRow.getCell(c);
        if (c === dateColIdx) {
          cell.numFmt = "DD.MM.YYYY";
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (c === resColIdx) {
          if (/^\d{1,4}$/.test(String(cell.value || "").trim())) {
            cell.numFmt = "0000";
          }
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (c === hoursColIdx && typeof cell.value === "number") {
          cell.numFmt = "#,##0.0";
          cell.alignment = { horizontal: "right", vertical: "middle" };
        }
      }

      // 4. Manuelle Korrekturen anpassen (Zeilen >= insertAtRowIndex verschieben sich um +1)
      const updatedCorrections = {};
      for (const [key, val] of Object.entries(state.appliedCorrections || {})) {
        const parts = key.split("_");
        const r = parseInt(parts[0], 10);
        const c = parseInt(parts[1], 10);
        if (r >= insertAtRowIndex) {
          updatedCorrections[`${r + 1}_${c}`] = val;
        } else {
          updatedCorrections[key] = val;
        }
      }
      state.appliedCorrections = updatedCorrections;

      // 5. Modal schließen
      closeInsertRowModal();

      // 6. Neu analysieren & UI aktualisieren
      state.lastInsertedRow = insertAtRowIndex;
      if (enteredRes) state.highlightEmployeeRes = enteredRes;

      // UI & Analyse aktualisieren (OHNE Scrollen nach oben und ohne Toast-Überschreibung!)
      runInspection({ skipScroll: true, skipToast: true });

      showToast(`✅ Zeile erfolgreich an Position ${insertAtRowIndex} in Prüfdatei eingefügt!`);

      // Wenn aus der Wochentabelle aufgerufen: Wochentabelle direkt wieder im Blick behalten
      if (state.modalCaller === "timesheet") {
        setTimeout(() => {
          const tsSec = document.getElementById("timesheet-section");
          if (tsSec) {
            tsSec.scrollIntoView({ behavior: "smooth", block: "start" });
          }
        }, 50);
      } else if (state.modalCaller === "fulltable") {
        activateTab("fulltable");
      }
    } catch (err) {
      console.error("Fehler beim Einfügen der Zeile:", err);
      alert("Fehler beim Einfügen der Zeile: " + err.message);
    }
  }

  // Event Listeners für Zeile nachtragen / bearbeiten Modal
  if (btnCloseInsertModal) btnCloseInsertModal.addEventListener("click", closeInsertRowModal);
  if (btnCancelInsertRow) btnCancelInsertRow.addEventListener("click", closeInsertRowModal);
  if (modalInsertRow) {
    modalInsertRow.addEventListener("click", (e) => {
      if (e.target === modalInsertRow) closeInsertRowModal();
    });
  }
  if (btnConfirmInsertRow) {
    btnConfirmInsertRow.addEventListener("click", () => {
      if (state.modalRowMode === "edit" && state.modalEditRowIdx) {
        executeSaveEditedRow(state.modalEditRowIdx);
      } else {
        executeInsertRow();
      }
    });
  }
  if (btnOpenInsertRowActionbar) {
    btnOpenInsertRowActionbar.addEventListener("click", () => {
      state.modalCaller = "timesheet";
      openInsertRowModal(null);
    });
  }
  if (btnOpenInsertRowFulltable) {
    btnOpenInsertRowFulltable.addEventListener("click", () => {
      state.modalCaller = "fulltable";
      openInsertRowModal(null);
    });
  }
  if (btnOpenInsertRowTimesheet) {
    btnOpenInsertRowTimesheet.addEventListener("click", () => {
      state.modalCaller = "timesheet";
      openInsertRowModal(null);
    });
  }

  // Event Listeners für Buchung-Auswählen Modal
  if (btnCloseChooseModal) btnCloseChooseModal.addEventListener("click", closeChooseBookingModal);
  if (btnCloseChoose) btnCloseChoose.addEventListener("click", closeChooseBookingModal);
  if (modalChooseBooking) {
    modalChooseBooking.addEventListener("click", (e) => {
      if (e.target === modalChooseBooking) closeChooseBookingModal();
    });
  }

  document.querySelectorAll('input[name="insert-pos-mode"]').forEach(radio => {
    radio.addEventListener("change", (e) => {
      const mode = e.target.value;
      if (mode === "after_row") {
        if (insertAfterRowContainer) insertAfterRowContainer.classList.remove("hidden");
        if (insertAutoHint) insertAutoHint.classList.add("hidden");
      } else {
        if (insertAfterRowContainer) insertAfterRowContainer.classList.add("hidden");
        if (mode === "auto") {
          if (insertAutoHint) insertAutoHint.classList.remove("hidden");
        } else {
          if (insertAutoHint) insertAutoHint.classList.add("hidden");
        }
      }
    });
  });

  if (insertRowSearch && insertTargetRowSelect) {
    insertRowSearch.addEventListener("input", () => {
      const q = insertRowSearch.value.trim().toLowerCase();
      let matchCount = 0;
      Array.from(insertTargetRowSelect.options).forEach(opt => {
        const matches = !q || opt.textContent.toLowerCase().includes(q);
        opt.style.display = matches ? "" : "none";
        if (matches) matchCount++;
      });
      const countEl = document.getElementById("insert-row-match-count");
      if (countEl) countEl.textContent = q ? `${matchCount} Treffer` : "";
    });
  }

  // --- Demo-Dateien direkt im Speicher laden ---
  if (btnLoadDemo) {
    btnLoadDemo.addEventListener("click", async () => {
      btnLoadDemo.disabled = true;
      btnLoadDemo.innerHTML = `<span class="icon">⏳</span> Erstelle Demo...`;

    // 0. Demo-Mitarbeiter im Browser speichern (falls noch keine gespeichert)
    if (state.staffList.length === 0) {
      const demoStaff = [
        { resource: "0045", name: "Max Mustermann", dept: "Montage" },
        { resource: "0120", name: "Anna Schmidt", dept: "Kundendienst" },
        { resource: "0300", name: "Michael Weber", dept: "Logistik" },
        { resource: "0400", name: "Sarah Fischer", dept: "Projektleitung" },
        { resource: "0500", name: "Thomas Becker", dept: "Qualitätssicherung" },
        { resource: "0600", name: "Julia Wagner", dept: "Service" },
        { resource: "0700", name: "Stefan Hoffmann", dept: "Instandhaltung" },
        { resource: "1773", name: "Uwe Althön", dept: "Bauleitung" },
        { resource: "1373", name: "Uwe Althön", dept: "Bauleitung" }
      ];
      saveStaffToStorage(demoStaff);
    }

    // 1. Referenz-Arbeitsmappe
    const refWb = new ExcelJS.Workbook();
    const refWs = refWb.addWorksheet("Stammdaten");
    refWs.columns = [
      { header: "Artikelnummer", key: "art", width: 16 },
      { header: "EAN_Code", key: "ean", width: 18 },
      { header: "Kundennummer", key: "knd", width: 16 },
      { header: "PLZ", key: "plz", width: 10 },
      { header: "Leistungsnummer", key: "leist", width: 18 },
      { header: "Resourcen Nummer", key: "res", width: 18 },
      { header: "Bezeichnung", key: "bez", width: 30 }
    ];

    const refRows = [
      ["10021", "4012345000101", "KND-80410", "10115", "12",   "0045", "Laptop Pro 15 Zoll"],
      ["10022", "4012345000102", "KND-80420", "20095", "45",   "0120", "Kabellose Maus Optical"],
      ["10023", "4012345000103", "KND-80430", "30159", "78",   "0300", "Mechanische Tastatur RGB"],
      ["10024", "4012345000104", "KND-80440", "40213", "88",   "0400", "USB-C Dockingstation"],
      ["10025", "4012345000105", "KND-80450", "50667", "1100", "0500", "Ultra-HD Monitor 27 Zoll"],
      ["10026", "4012345000106", "KND-80460", "60311", "1200", "0600", "Noise-Cancelling Headset"],
      ["10027", "4012345000107", "KND-80470", "70173", "1300", "0700", "Externe NVMe SSD 1TB"],
      ["10028", "4012345000108", "KND-80480", "80331", "1400", "1773", "Baustellen-Server Mobile"]
    ];

    refRows.forEach(r => refWs.addRow(r));

    // Arbeitsblatt "Personalstamm" in der Referenzdatei (Spalte A: Personalnummer, Spalte B: Name)
    const staffWs = refWb.addWorksheet("Personalstamm");
    staffWs.columns = [
      { header: "Personalnummer", key: "persnr", width: 18 },
      { header: "Name", key: "name", width: 25 },
      { header: "Abteilung", key: "dept", width: 20 }
    ];
    const staffRows = [
      ["0045", "Max Mustermann", "Montage"],
      ["0120", "Anna Schmidt", "Kundendienst"],
      ["0300", "Michael Weber", "Logistik"],
      ["0400", "Sarah Fischer", "Projektleitung"],
      ["0500", "Thomas Becker", "Qualitätssicherung"],
      ["0600", "Julia Wagner", "Service"],
      ["0700", "Stefan Hoffmann", "Instandhaltung"],
      ["1773", "Uwe Althön", "Bauleitung"],
      ["1373", "Uwe Althön", "Bauleitung"]
    ];
    staffRows.forEach(r => staffWs.addRow(r));

    // Mitarbeiter direkt aus der Referenz-Arbeitsmappe extrahieren & registrieren
    extractStaffFromReferenceWorkbook(refWb);

    const refBuf = await refWb.xlsx.writeBuffer();
    state.refWorkbook = refWb;
    state.refFileName = "Beispiel_Referenz_Stammdaten.xlsx";
    refFilename.textContent = state.refFileName;
    populateSheetSelect(selectRefSheet, refWb.worksheets);
    state.currentRefSheet = selectRefSheet.value;
    refFileInfo.classList.remove("hidden");
    dropRef.querySelector(".drop-zone-content").classList.add("hidden");

    // Falls noch keine Referenzdatei im Browser gespeichert ist, Demo-Referenz speichern
    const storedRef = await loadRefFileFromStorage();
    if (!storedRef) {
      await saveRefFileToStorage(state.refFileName, refBuf, state.currentRefSheet);
      if (refStorageBadge) {
        refStorageBadge.textContent = "Im Browser gespeichert";
        refStorageBadge.className = "badge badge-ok";
      }
      if (refInfoDetail) {
        refInfoDetail.textContent = "💾 Im Browser gespeichert (dauerhaft erhalten)";
      }
    } else {
      if (refStorageBadge) {
        refStorageBadge.textContent = "Demo aktiv";
        refStorageBadge.className = "badge badge-warning";
      }
      if (refInfoDetail) {
        refInfoDetail.textContent = "Demo-Daten aktiv (gespeicherte Datei bleibt erhalten)";
      }
    }

    // 2. Prüfdatei-Arbeitsmappe (mit Datum, Menge/Stunden und Ressourcen)
    const tgtWb = new ExcelJS.Workbook();
    const tgtWs = tgtWb.addWorksheet("Report");
    tgtWs.columns = [
      { header: "Datum", key: "dat", width: 14 },
      { header: "Artikelnummer", key: "art", width: 16 },
      { header: "Resourcen Nummer", key: "res", width: 18 },
      { header: "Menge", key: "mng", width: 10 },
      { header: "Leistung", key: "leist", width: 16 },
      { header: "EAN_Code", key: "ean", width: 18 }
    ];

    const tgtRows = [
      ["15.09.2026", "10021", "0045", 8.0, "",     "4012345000101"], // Leistung leer -> OK_LEER; 8 Std für Max Mustermann
      ["15.09.2026", "10022", "45",   4.5, "45",   "4012345000102"], // Res '45' -> wird '0045'; Leistung '45' OK; 4.5 Std für Max
      ["16.09.2026", "10023", "120",  7.5, "78",   "4012345000103"], // Res '120' -> wird '0120'; Leistung '78' OK; 7.5 Std für Anna Schmidt
      ["16.09.2026", "10042", "0400", 8.0, "54",   "4012345000104"], // Artikel Zahlendreher (10042); Leistung Zahlendreher ('54'); 8 Std Sarah Fischer
      ["17.09.2026", "10025", "0500", 6.0, "1100", "4012345000150"], // EAN Zahlendreher; 6 Std Thomas Becker
      ["17.09.2026", "99999", "0700", 8.5, "",     "4012345000107"], // 99999 Nicht existent; 8.5 Std Stefan Hoffmann
      ["17.09.2026", "10028", "1773", 8.0, "1400", "4012345000108"], // 8 Std Uwe Althön (Personalnummer 1773)
      ["17.09.2026", "10027", "1373", 7.5, "1300", "4012345000107"], // 7.5 Std Uwe Althön (Personalnummer 1373)
      ["17.09.2026", "10026", "9900", 4.0, "1200", "4012345000106"]  // Fremd-Ressource 9900 (Maschine/Bagger - kein Mitarbeiter)
    ];

    tgtRows.forEach(r => tgtWs.addRow(r));
    const tgtBuf = await tgtWb.xlsx.writeBuffer();
    state.targetWorkbook = tgtWb;
    state.lastTargetBuffer = tgtBuf;
    state.targetFileName = "Beispiel_Zu_Pruefen.xlsx";
    tgtFilename.textContent = state.targetFileName;
    populateSheetSelect(selectTgtSheet, tgtWb.worksheets);
    state.currentTargetSheet = selectTgtSheet.value;
    tgtFileInfo.classList.remove("hidden");
    dropTgt.querySelector(".drop-zone-content").classList.add("hidden");

    if (tgtStatusBadge) {
      tgtStatusBadge.textContent = "Bereit zur Prüfung";
      tgtStatusBadge.className = "badge badge-ok";
      tgtStatusBadge.style.display = "inline-block";
    }

    btnLoadDemo.disabled = false;
    btnLoadDemo.innerHTML = `<span class="icon">✨</span> Demo-Dateien laden`;

    checkReadyForConfig();
    showToast("Demo-Dateien & Mitarbeiter geladen! Klicken Sie auf 'Prüfung starten'.");
  });
  }

  function getColLetter(colIdx) {
    let temp = "";
    let letter = "";
    while (colIdx > 0) {
      temp = (colIdx - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIdx = (colIdx - temp - 1) / 26;
    }
    return letter;
  }

  let toastTimer = null;
  function showToast(msg) {
    if (toastTimer) clearTimeout(toastTimer);
    toast.textContent = msg;
    toast.classList.remove("hidden");
    toastTimer = setTimeout(() => {
      toast.classList.add("hidden");
    }, 3500);
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
});

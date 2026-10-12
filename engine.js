// ==============================================================================
// Excel-Prüfer Web Engine (100% Client-Side für Netlify & Browser)
// Damerau-Levenshtein, O(1) Anagramm-Zahlendreher-Erkennung, 0000-Formatierung
// ==============================================================================

function extractCellValue(val, cell) {
  if (val === null || val === undefined) return "";

  // 1. Primitive Typen (Zahl, Boolean)
  if (typeof val === "number" || typeof val === "boolean") {
    return String(val);
  }

  // 2. String
  if (typeof val === "string") {
    const trimmed = val.trim();
    if (trimmed === "[object Object]" || trimmed === "object Objekt" || trimmed === "(object Objekt)" || trimmed.toLowerCase().includes("object object")) {
      if (cell && typeof cell.text === "string" && cell.text && !cell.text.toLowerCase().includes("object")) {
        return cell.text;
      }
      return "";
    }
    return val;
  }

  // 3. Date-Objekt
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      const day = String(val.getDate()).padStart(2, "0");
      const month = String(val.getMonth() + 1).padStart(2, "0");
      const year = val.getFullYear();
      return `${day}.${month}.${year}`;
    }
    return "";
  }

  // 4. Array (z.B. richText oder Listen)
  if (Array.isArray(val)) {
    return val.map(part => extractCellValue(part, cell)).filter(Boolean).join("");
  }

  // 5. ExcelJS-Objekte (Formeln, Hyperlinks, RichText etc.)
  if (typeof val === "object") {
    // a) Formel mit berechnetem Ergebnis: { formula: '...', result: '...' }
    if ("result" in val && val.result !== undefined && val.result !== null) {
      return extractCellValue(val.result, cell);
    }
    // b) Shared Formula mit Ergebnis
    if ("sharedFormula" in val && "result" in val && val.result !== undefined && val.result !== null) {
      return extractCellValue(val.result, cell);
    }
    // c) Hyperlink: { text: '...', hyperlink: '...' }
    if ("text" in val && val.text !== undefined && val.text !== null) {
      return extractCellValue(val.text, cell);
    }
    // d) Rich Text: { richText: [{ text: '...' }, ...] }
    if (Array.isArray(val.richText)) {
      return val.richText.map(part => (part && part.text) ? String(part.text) : "").join("");
    }
    // e) Excel-Fehler: { error: '#N/A' }
    if ("error" in val) {
      return String(val.error);
    }
    // f) Falls cell.text existiert (ExcelJS formatierter Text)
    if (cell && typeof cell.text === "string" && cell.text && !cell.text.toLowerCase().includes("object")) {
      return cell.text;
    }
    // g) Verschachtelter Wert
    if ("value" in val && val.value !== undefined) {
      return extractCellValue(val.value, cell);
    }
    // h) Formel ohne Ergebnis -> Formel selbst nehmen
    if ("formula" in val && val.formula) {
      return String(val.formula);
    }
  }

  // 6. cell.text als direkter Fallback
  if (cell && typeof cell.text === "string" && cell.text && !cell.text.toLowerCase().includes("object")) {
    return cell.text;
  }

  const s = String(val).trim();
  return (s.toLowerCase().includes("object")) ? "" : s;
}

function normalizeCellValue(val, cell) {
  const extracted = extractCellValue(val, cell);
  let s = String(extracted).trim();
  // Wandle floats wie 12345.0 in 12345 um
  if (/^-?\d+\.0$/.test(s)) {
    s = s.slice(0, -2);
  }
  return s;
}

function padNumber(val, minLength = 4) {
  let s = String(val).trim();
  if (/^\d+$/.test(s) && s.length < minLength) {
    return s.padStart(minLength, "0");
  }
  return s;
}

function parseDateValue(val, cell) {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === "object" && val.result instanceof Date) {
    return isNaN(val.result.getTime()) ? null : val.result;
  }
  if (typeof val === "number" && val >= 20000 && val <= 90000) {
    const ms = Math.round((val - 25569) * 86400000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  let s = "";
  if (cell && typeof cell.text === "string" && cell.text && !cell.text.toLowerCase().includes("object")) {
    s = cell.text.trim();
  } else {
    s = String(val).trim();
  }
  if (!s || s.toLowerCase().includes("object")) return null;

  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const f = parseFloat(s);
    if (f >= 20000 && f <= 90000) {
      const ms = Math.round((f - 25569) * 86400000);
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d;
    }
  }

  const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (dmy) {
    let day = parseInt(dmy[1], 10);
    let month = parseInt(dmy[2], 10) - 1;
    let year = parseInt(dmy[3], 10);
    if (year < 100) year += (year < 50 ? 2000 : 1900);
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime())) return d;
  }

  const ymd = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (ymd) {
    let year = parseInt(ymd[1], 10);
    let month = parseInt(ymd[2], 10) - 1;
    let day = parseInt(ymd[3], 10);
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime())) return d;
  }

  const fallback = new Date(s);
  if (!isNaN(fallback.getTime())) return fallback;
  return null;
}

function formatDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}.${month}.${year}`;
}


function damerauLevenshtein(s1, s2) {
  const len1 = s1.length;
  const len2 = s2.length;
  if (len1 === 0) return len2;
  if (len2 === 0) return len1;

  // DP Matrix
  const d = [];
  for (let i = 0; i <= len1 + 1; i++) {
    d[i] = new Array(len2 + 2).fill(0);
  }

  const maxdist = len1 + len2;
  d[0][0] = maxdist;
  for (let i = 0; i <= len1; i++) {
    d[i + 1][0] = maxdist;
    d[i + 1][1] = i;
  }
  for (let j = 0; j <= len2; j++) {
    d[0][j + 1] = maxdist;
    d[1][j + 1] = j;
  }

  const da = {};
  for (let i = 1; i <= len1; i++) {
    let db = 0;
    for (let j = 1; j <= len2; j++) {
      const k = da[s2[j - 1]] || 0;
      const l = db;
      let cost = 0;
      if (s1[i - 1] === s2[j - 1]) {
        db = j;
      } else {
        cost = 1;
      }

      d[i + 1][j + 1] = Math.min(
        d[i][j] + cost,             // Ersetzung
        d[i + 1][j] + 1,            // Einfügung
        d[i][j + 1] + 1,            // Löschung
        d[k][l] + (i - k - 1) + 1 + (j - l - 1) // Transposition (Zahlendreher)
      );
    }
    da[s1[i - 1]] = i;
  }

  return d[len1 + 1][len2 + 1];
}

class ColumnIndex {
  constructor(values) {
    this.exactSet = new Set();
    this.anagramMap = new Map();
    this.lengthMap = new Map();
    this.allValues = [];

    for (const v of values) {
      const norm = normalizeCellValue(v);
      if (!norm) continue;

      if (!this.exactSet.has(norm)) {
        this.exactSet.add(norm);
        this.allValues.push(norm);

        // Anagramm-Schlüssel: sortierte Zeichenfolge
        const key = norm.split("").sort().join("");
        if (!this.anagramMap.has(key)) {
          this.anagramMap.set(key, []);
        }
        this.anagramMap.get(key).push(norm);

        // Längen-Map
        const len = norm.length;
        if (!this.lengthMap.has(len)) {
          this.lengthMap.set(len, []);
        }
        this.lengthMap.get(len).push(norm);
      }
    }
  }

  contains(val) {
    return this.exactSet.has(normalizeCellValue(val));
  }

  findCandidates(target, maxCandidates = 4) {
    const normTarget = normalizeCellValue(target);
    if (!normTarget) return [];

    const candidates = [];
    const alreadySeen = new Set();

    // 1. Zahlendreher & Zifferntausch per O(1) Anagramm-Hashing
    const sortedKey = normTarget.split("").sort().join("");
    if (this.anagramMap.has(sortedKey)) {
      for (const refVal of this.anagramMap.get(sortedKey)) {
        if (refVal === normTarget) continue;
        alreadySeen.add(refVal);

        const dist = damerauLevenshtein(normTarget, refVal);
        if (dist === 1) {
          candidates.push({
            value: refVal,
            score: 0.98,
            type: "ZAHLENDREHER",
            detail: `Echter Zahlendreher (benachbarte Ziffern vertauscht)`
          });
        } else {
          candidates.push({
            value: refVal,
            score: 0.92,
            type: "ZIFFERENTAUSCH",
            detail: `Zifferntausch (Distanz: ${dist})`
          });
        }
      }
    }

    // 2. Tippfehler (gleiche Länge, 1 Ziffer abweichend)
    const sameLen = this.lengthMap.get(normTarget.length) || [];
    for (const refVal of sameLen) {
      if (alreadySeen.has(refVal) || refVal === normTarget) continue;
      const dist = damerauLevenshtein(normTarget, refVal);
      if (dist === 1) {
        candidates.push({
          value: refVal,
          score: 0.88,
          type: "TIPPFEHLER",
          detail: `Tippfehler (genau 1 Ziffer weicht ab)`
        });
        alreadySeen.add(refVal);
      }
    }

    // 3. Ziffer zu viel / zu wenig (+-1 Länge)
    for (const l of [normTarget.length - 1, normTarget.length + 1]) {
      const nearLen = this.lengthMap.get(l) || [];
      for (const refVal of nearLen) {
        if (alreadySeen.has(refVal)) continue;
        const dist = damerauLevenshtein(normTarget, refVal);
        if (dist === 1) {
          const type = (refVal.length > normTarget.length) ? "ZIFFER_FEHLT" : "ZIFFER_ZUVIEL";
          const label = (refVal.length > normTarget.length) ? "1 Ziffer fehlt" : "1 Ziffer zu viel";
          candidates.push({
            value: refVal,
            score: 0.82,
            type: type,
            detail: `${label} gegenüber Referenz`
          });
          alreadySeen.add(refVal);
        }
      }
    }

    // 4. Fallback: Nächste Alternative in Stammdaten
    if (candidates.length === 0 && this.allValues.length > 0) {
      const scored = [];
      const tIsNum = /^\d+$/.test(normTarget);
      const tNum = tIsNum ? parseInt(normTarget, 10) : 0;

      for (const rVal of this.allValues) {
        if (alreadySeen.has(rVal)) continue;
        const dist = damerauLevenshtein(normTarget, rVal);
        const rIsNum = /^\d+$/.test(rVal);
        const numDiff = (tIsNum && rIsNum) ? Math.abs(tNum - parseInt(rVal, 10)) : 999999;
        const maxLen = Math.max(normTarget.length, rVal.length);
        const sim = Math.max(0.1, Number((1.0 - (dist / Math.max(1, maxLen))).toFixed(2)));
        scored.push({ dist, numDiff, rVal, sim });
      }

      scored.sort((a, b) => {
        if (a.dist !== b.dist) return a.dist - b.dist;
        return a.numDiff - b.numDiff;
      });

      for (let i = 0; i < Math.min(maxCandidates, scored.length); i++) {
        const item = scored[i];
        candidates.push({
          value: item.rVal,
          score: item.sim,
          type: "NAECHSTE_ALTERNATIVE",
          detail: `Nächstgelegene Alternative in Stammdaten (Distanz: ${item.dist})`
        });
      }
    }

    candidates.sort((a, b) => b.score - a.score);
    return candidates.slice(0, maxCandidates);
  }
}

// Global für Web-App
window.WebExcelEngine = {
  extractCellValue,
  normalizeCellValue,
  padNumber,
  parseDateValue,
  formatDate,
  damerauLevenshtein,
  ColumnIndex
};


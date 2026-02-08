const COLUMN_LETTERS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));

function colIndexToLetter(index) {
  return COLUMN_LETTERS[index] || '';
}

function letterToColIndex(letter) {
  return COLUMN_LETTERS.indexOf(letter.toUpperCase());
}

function readCell(sheet, cellRef) {
  const cell = sheet[cellRef];
  return cell ? cell.v : undefined;
}

function normalizeCellRef(value) {
  return String(value || '').trim().toUpperCase();
}

function normalizeAnchorRef(value) {
  return String(value || '').trim();
}

function readCellValue(sheet, cellRef) {
  if (!cellRef) return undefined;
  return readCell(sheet, normalizeCellRef(cellRef));
}

function readCellNumber(sheet, cellRef) {
  const value = readCellValue(sheet, cellRef);
  const numeric = parseNumericValue(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function parseNumericValue(value) {
  if (value === undefined || value === null) return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const normalized = String(value).replace(/\s+/g, '').replace(/,/g, '.');
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
}

function round2(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function normalizeLabelText(value) {
  // normalize for matching labels like "Трудоемкость, человеко-часы=" -> "трудоемкостьчеловекочасы"
  return String(value ?? '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/=/g, '')
    .replace(/\s+/g, '')
    .replace(/[^0-9a-zа-я]/g, '');
}

function parseRef(rawRef, defaultSheet) {
  if (!rawRef) return null;
  const cleaned = String(rawRef).trim();

  const fallbackSheet = defaultSheet || state.activeSheet || null;

  // Plain cell reference (e.g. D92)
  const plain = cleaned.match(/^(\$?[A-Za-z]{1,3}\$?\d+)$/);
  if (plain) {
    const cellRef = plain[1].replace(/\$/g, '').toUpperCase();
    return { sheetName: fallbackSheet, cellRef };
  }

  // Sheet!Cell (supports quoted sheet names)
  const match = cleaned.match(/^(?:'([^']+)'|([^'!]+))!\s*(\$?[A-Za-z]{1,3}\$?\d+)$/);
  if (!match) return null;
  const sheetName = match[1] || match[2] || fallbackSheet;
  const cellRef = match[3].replace(/\$/g, '').toUpperCase();
  return { sheetName, cellRef };
}

function readRefCell(workbook, rawRef, fallbackSheet) {
  const parsed = parseRef(rawRef, fallbackSheet);
  if (!parsed || !parsed.sheetName) return null;
  const sheet = workbook.Sheets[parsed.sheetName];
  if (!sheet) return null;
  const cell = sheet[parsed.cellRef];
  return { sheetName: parsed.sheetName, cellRef: parsed.cellRef, cell, sheet };
}

function isNumericCell(cell) {
  if (!cell) return false;
  if (typeof cell.v === 'number') return true;
  return Number.isFinite(parseNumericValue(cell.v));
}

function parseRefsFromFormula(formulaString) {
  if (!formulaString) return [];
  const formula = String(formulaString).replace(/^=/, '');
  const regex = /(?:'[^']+'|[A-Za-z0-9_]+)?!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g;
  return formula.match(regex) || [];
}

function expandRangeRefs(rangeRef, sheetName) {
  const [start, end] = rangeRef.split(':');
  if (!end) return [`${sheetName}!${start}`];
  const range = XLSX.utils.decode_range(`${start}:${end}`);
  const refs = [];
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      refs.push(`${sheetName}!${XLSX.utils.encode_cell({ r, c })}`);
    }
  }
  return refs;
}

function expandRanges(refs, defaultSheet) {
  if (!Array.isArray(refs)) return [];
  const out = [];

  refs.forEach((raw) => {
    if (!raw) return;
    const cleaned = String(raw).trim();

    // Extract optional sheet name: 'Sheet'!A1 or Sheet!A1 (also supports ranges A1:B2)
    const sheetMatch = cleaned.match(/^(?:'([^']+)'|([^'!]+))!([A-Za-z0-9$]+(?::[A-Za-z0-9$]+)?)$/);

    let sheetName = defaultSheet || state.activeSheet || null;
    let refPart = cleaned;

    if (sheetMatch) {
      sheetName = sheetMatch[1] || sheetMatch[2] || sheetName;
      refPart = sheetMatch[3];
    }

    refPart = refPart.replace(/\$/g, '').toUpperCase();

    if (refPart.includes(':')) {
      // Range: if we know sheet - return full refs, else return plain cell refs (caller will add sheet)
      if (sheetMatch && sheetName) {
        expandRangeRefs(refPart, sheetName).forEach((full) => out.push(full));
      } else {
        const range = XLSX.utils.decode_range(refPart);
        for (let r = range.s.r; r <= range.e.r; r += 1) {
          for (let c = range.s.c; c <= range.e.c; c += 1) {
            out.push(XLSX.utils.encode_cell({ r, c }));
          }
        }
      }
    } else if (sheetMatch && sheetName) {
      out.push(`${sheetName}!${refPart}`);
    } else {
      out.push(refPart);
    }
  });

  return out;
}

function getCellFormulaString(cell) {
  if (!cell) return '';
  if (typeof cell.f === 'string' && cell.f.trim()) return cell.f.trim();
  if (typeof cell.v === 'string' && cell.v.trim().startsWith('=')) return cell.v.trim().slice(1);
  return '';
}

function toFullRef(refLike, defaultSheet) {
  const parsed = parseRef(refLike, defaultSheet);
  if (!parsed) return null;
  return `${parsed.sheetName}!${parsed.cellRef}`;
}

function decodeRefParts(fullRef, defaultSheet) {
  const parsed = parseRef(fullRef, defaultSheet);
  if (!parsed) return null;
  const addr = XLSX.utils.decode_cell(parsed.cellRef);
  return { sheetName: parsed.sheetName, cellRef: parsed.cellRef, r: addr.r, c: addr.c, col: XLSX.utils.encode_col(addr.c) };
}

function extractA1Refs(formulaString) {
  if (!formulaString) return [];
  const formula = String(formulaString).replace(/\$/g, '').toUpperCase();
  return formula.match(/[A-Z]{1,3}\d+/g) || [];
}

function expandSumRange(rangeRef) {
  if (!rangeRef) return [];
  const cleaned = rangeRef.replace(/\$/g, '').toUpperCase();
  if (!cleaned.includes(':')) return [cleaned];
  const range = XLSX.utils.decode_range(cleaned);
  const refs = [];
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      refs.push(XLSX.utils.encode_cell({ r, c }));
    }
  }
  return refs;
}

function parseSumFormulaRefs(formulaString) {
  if (!formulaString) return [];
  const formula = String(formulaString).replace(/\s+/g, '');
  const match = formula.match(/sum\(([^)]+)\)/i);
  if (!match) return [];
  const inner = match[1];
  const parts = inner.split(/[;,]/).map((part) => part.trim()).filter(Boolean);
  return parts.flatMap((part) => expandSumRange(part));
}

const state = {
  workbook: null,
  mapping: {},
  originalSpec: null,
  newSpec: null,
  worker: null,
  activeSheet: null,
  previewHoverBound: false,
  activeCellInput: null,
  previewClickBound: false,
  activeResultsTab: 'corpus',
  calcSummary: null,
};

const COLUMN_LETTERS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
const MATERIAL_DENSITY = 730;
const MAX_SECTION_WIDTH = 1200;
const PARTITION_THRESHOLD = 800;

function setUploadError(message) {
  const box = document.getElementById('upload-error');
  if (!box) return;
  if (message) {
    box.textContent = message;
    box.classList.remove('hidden');
  } else {
    box.textContent = '';
    box.classList.add('hidden');
  }
}

function colIndexToLetter(index) {
  return COLUMN_LETTERS[index] || '';
}

function letterToColIndex(letter) {
  return COLUMN_LETTERS.indexOf(letter.toUpperCase());
}

function getWorker() {
  if (state.worker) return state.worker;
  const workerSource = document.getElementById('worker-src').textContent;
  const blob = new Blob([workerSource], { type: 'text/javascript' });
  const url = URL.createObjectURL(blob);
  state.worker = new Worker(url);
  return state.worker;
}

function showScreen(id) {
  document.querySelectorAll('.screen').forEach((screen) => {
    screen.classList.add('hidden');
    screen.classList.remove('active');
  });
  const target = document.getElementById(id);
  target.classList.remove('hidden');
  target.classList.add('active');
  if (id === 'results-screen' && state.originalSpec) {
    renderBaseSummary(state.originalSpec);
  }
}

function populateColumnSelects() {
  document.querySelectorAll('select[data-highlight]').forEach((select) => {
    select.innerHTML = COLUMN_LETTERS.map((letter) => `<option value="${letter}">${letter}</option>`).join('');
  });
}

function renderPreview(sheet) {
  const table = document.getElementById('preview-table');
  table.innerHTML = '';
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const rowCount = range.e.r + 1;
  const maxCols = range.e.c + 1;

  const headerRow = document.createElement('tr');
  headerRow.appendChild(document.createElement('th'));
  for (let i = 0; i < maxCols; i += 1) {
    const th = document.createElement('th');
    th.textContent = colIndexToLetter(i);
    headerRow.appendChild(th);
  }
  table.appendChild(headerRow);

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const row = json[rowIndex] || [];
    const tr = document.createElement('tr');
    const rowHeader = document.createElement('th');
    rowHeader.textContent = rowIndex + 1;
    tr.appendChild(rowHeader);
    for (let col = 0; col < maxCols; col += 1) {
      const td = document.createElement('td');
      td.textContent = row[col] ?? '';
      td.dataset.col = col;
      td.dataset.row = rowIndex;
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  attachPreviewHover();
  attachCellSelection();
}

function highlightColumn(colIndex) {
  const table = document.getElementById('preview-table');
  if (!table) return;
  table.querySelectorAll('td, th').forEach((cell) => {
    cell.classList.remove('highlighted');
  });
  table.querySelectorAll(`td:nth-child(${colIndex + 2}), th:nth-child(${colIndex + 2})`).forEach((cell) => {
    cell.classList.add('highlighted');
  });
}

function attachPreviewHover() {
  if (state.previewHoverBound) return;
  const table = document.getElementById('preview-table');
  const indicator = document.getElementById('cursor-indicator');

  const clearHover = () => {
    table.querySelectorAll('.row-hover, .col-hover, .cell-hover').forEach((cell) => {
      cell.classList.remove('row-hover', 'col-hover', 'cell-hover');
    });
  };

  table.addEventListener('mouseover', (event) => {
    const cell = event.target.closest('td');
    if (!cell || !table.contains(cell)) return;
    const rowIndex = Number(cell.dataset.row);
    const colIndex = Number(cell.dataset.col);
    if (Number.isNaN(rowIndex) || Number.isNaN(colIndex)) return;
    clearHover();

    cell.classList.add('cell-hover');

    const row = table.querySelectorAll('tr')[rowIndex + 1];
    if (row) {
      row.querySelectorAll('td, th').forEach((item) => item.classList.add('row-hover'));
    }

    table.querySelectorAll(`tr > :nth-child(${colIndex + 2})`).forEach((item) => {
      item.classList.add('col-hover');
    });

    if (indicator) {
      const cellRef = `${colIndexToLetter(colIndex)}${rowIndex + 1}`;
      indicator.textContent = `${cellRef}: ${cell.textContent || '—'}`;
    }
  });

  table.addEventListener('mouseleave', () => {
    clearHover();
    if (indicator) {
      indicator.textContent = 'Наведите курсор на ячейку';
    }
  });

  state.previewHoverBound = true;
}

function attachCellSelection() {
  if (state.previewClickBound) return;
  const table = document.getElementById('preview-table');
  const indicator = document.getElementById('cursor-indicator');
  table.addEventListener('click', (event) => {
    const cell = event.target.closest('td');
    if (!cell || !state.activeCellInput) return;
    const rowIndex = Number(cell.dataset.row);
    const colIndex = Number(cell.dataset.col);
    if (Number.isNaN(rowIndex) || Number.isNaN(colIndex)) return;
    const cellRef = `${colIndexToLetter(colIndex)}${rowIndex + 1}`;
    state.activeCellInput.value = cellRef;
    if (indicator) {
      indicator.textContent = `Выбрана ячейка: ${cellRef}`;
    }
  });

  state.previewClickBound = true;
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

function findCalcSummaryAnchors(workbook) {
  // Keys must be normalized with normalizeLabelText()
  const labels = {
    вескг: 'weightRef',
    трудоемкость: 'laborHoursRef',
    стоимостьдсп: 'dspRef',
    стоимостькромки: 'edgeRef',
    стоимостьпластика: 'plasticRef',
    стоимостьткани: 'fabricRef',
    стоимостьфурнитурыимп: 'hwImpRef',
    стоимостьфурнитурыотч: 'hwRepRef',
    стоимостьупаковки: 'packRef',
    трудрабочих: 'laborRef',
  };

  const labelEntries = Object.entries(labels).sort((a, b) => b[0].length - a[0].length);

  let totalCostRef = null;
  const anchors = {};

  workbook.SheetNames.forEach((sheetName) => {
    const ws = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell || cell.v == null) continue;

        const normalized = normalizeLabelText(cell.v);
        if (!normalized) continue;


        // total cost label (different templates call it either "Прямые затраты" or "Стоимость расчета суммарно")
        if (!totalCostRef) {
          const isTotalCostLabel = (normalized === 'стоимостьрасчетасуммарно')
            || normalized.startsWith('прямыезатрат')
            || normalized.includes('прямыезатрат');
          if (isTotalCostLabel) {
            const right = XLSX.utils.encode_cell({ r, c: c + 1 });
            const rightCell = ws[right];
            if (rightCell && isNumericCell(rightCell)) {
              totalCostRef = `${sheetName}!${right}`;
            } else if (rightCell) {
              // still accept non-empty (e.g. formula / string number), numeric will be parsed later
              totalCostRef = `${sheetName}!${right}`;
            }
          }
        }

        // other anchors (match by contains/prefix to tolerate extra words like ", человеко-часы")
        for (const [labelKey, anchorKey] of labelEntries) {
          if (normalized === labelKey || normalized.startsWith(labelKey) || normalized.includes(labelKey)) {
            if (anchors[anchorKey]) break;
            const right = XLSX.utils.encode_cell({ r, c: c + 1 });
            if (ws[right]) anchors[anchorKey] = `${sheetName}!${right}`;
            break;
          }
        }
      }
    }
  });

  return { totalCostRef, ...anchors };
}

function extractRefsFromFormula(formulaString) {
  if (!formulaString) return [];
  const formula = String(formulaString).replace(/^=/, '');
  const regex = /(?:'[^']+'|[A-Za-z0-9_]+)?!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g;
  const matches = formula.match(regex) || [];
  return matches;
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


function traceCellLeaves(workbook, rootRef, opts = {}) {
  const leaves = [];
  const tree = [];
  const maxDepth = opts.maxDepth || 20;
  const visited = opts.visited || new Set();

  const trace = (ref, depth, path) => {
    if (!ref || depth > maxDepth) return;
    if (visited.has(ref)) return;
    visited.add(ref);
    const parsed = parseRef(ref);
    if (!parsed) return;
    const sheet = workbook.Sheets[parsed.sheetName];
    if (!sheet) return;
    const cell = sheet[parsed.cellRef];
    const node = { ref, children: [] };
    if (cell?.f) {
      const refs = extractRefsFromFormula(cell.f);
      if (refs.length) {
        refs.forEach((raw) => {
          const sheetMatch = raw.match(/^(?:'([^']+)'|([^'!]+))!([A-Za-z0-9$]+(?::[A-Za-z0-9$]+)?)$/);
          let targetSheet = parsed.sheetName;
          let targetRef = raw;
          if (sheetMatch) {
            targetSheet = sheetMatch[1] || sheetMatch[2] || parsed.sheetName;
            targetRef = sheetMatch[3];
          }
          const normalized = targetRef.replace(/\$/g, '');
          if (normalized.includes(':')) {
            const expanded = expandRangeRefs(normalized, targetSheet);
            expanded.forEach((expandedRef) => {
              node.children.push(expandedRef);
              trace(expandedRef, depth + 1, [...path, ref]);
            });
          } else {
            const fullRef = `${targetSheet}!${normalized}`;
            node.children.push(fullRef);
            trace(fullRef, depth + 1, [...path, ref]);
          }
        });
      } else {
        const value = parseNumericValue(cell.v);
        leaves.push({
          ref,
          value: Number.isFinite(value) ? value : 0,
          sheet: parsed.sheetName,
          row: XLSX.utils.decode_cell(parsed.cellRef).r + 1,
          col: XLSX.utils.decode_cell(parsed.cellRef).c + 1,
        });
      }
    } else {
      const value = parseNumericValue(cell?.v);
      leaves.push({
        ref,
        value: Number.isFinite(value) ? value : 0,
        sheet: parsed.sheetName,
        row: XLSX.utils.decode_cell(parsed.cellRef).r + 1,
        col: XLSX.utils.decode_cell(parsed.cellRef).c + 1,
      });
    }
    tree.push(node);
  };

  trace(rootRef, 0, []);
  return { leaves, tree };
}

function findHeaderRow(sheetMatrix, rowIndex) {
  const start = Math.max(rowIndex - 10, 1);
  for (let row = rowIndex - 1; row >= start; row -= 1) {
    const rowData = sheetMatrix[row - 1] || [];
    const normalizedRow = rowData.map((cell) => normalizeLabelText(cell));
    if (normalizedRow.some((text) => text.includes('наимен')) && normalizedRow.some((text) => text.includes('кол'))) {
      return row;
    }
  }
  return null;
}

function inferDetailRowContext(sheetMatrix, rowIndex) {
  const headerRow = findHeaderRow(sheetMatrix, rowIndex);
  if (!headerRow) {
    return { rowIndex };
  }

  const headerData = sheetMatrix[headerRow - 1] || [];
  const mapping = {};
  const qtyCandidates = [];
  const lengthCandidates = [];
  const widthCandidates = [];
  const areaCandidates = [];
  const costCandidates = [];

  headerData.forEach((cell, idx) => {
    const text = normalizeLabelText(cell);
    if (text.includes('наимен')) mapping.name = idx;
    if (text.includes('кол') || text.includes('qty') || text.includes('шт')) qtyCandidates.push(idx);
    if (text.includes('площад')) areaCandidates.push(idx);
    if (text.includes('стоим') || text.includes('цена') || text.includes('cost')) costCandidates.push(idx);
    if (text.includes('длин')) lengthCandidates.push(idx);
    if (text.includes('ширин')) widthCandidates.push(idx);
    if (text.includes('толщ')) mapping.thickness = idx;
  });

  // Best-effort defaults (the "real" qty is usually the last qty-like column).
  if (qtyCandidates.length) mapping.qty = qtyCandidates[qtyCandidates.length - 1];
  if (areaCandidates.length) mapping.area = areaCandidates[areaCandidates.length - 1];
  if (costCandidates.length) mapping.cost = costCandidates[costCandidates.length - 1];
  if (lengthCandidates.length) mapping.length = lengthCandidates[lengthCandidates.length - 1];
  if (widthCandidates.length) mapping.width = widthCandidates[widthCandidates.length - 1];

  // Refine qty column: pick the one that best matches area ≈ (L*W)*Q across a sample of rows.
  if (qtyCandidates.length >= 2 && mapping.length != null && mapping.width != null && mapping.area != null) {
    const sampleStart = Math.max(headerRow, rowIndex - 20);
    const sampleEnd = Math.min(sheetMatrix.length, rowIndex + 20);
    const scores = new Map();

    const pieceArea = (len, wid) => {
      if (!isFinite(len) || !isFinite(wid) || len <= 0 || wid <= 0) return null;
      return (len * wid) / 1e6;
    };

    for (const cand of qtyCandidates) {
      let score = 0;
      let used = 0;
      for (let r = sampleStart; r <= sampleEnd; r++) {
        const row = sheetMatrix[r - 1] || [];
        const len = parseNumericValue(row[mapping.length]);
        const wid = parseNumericValue(row[mapping.width]);
        const area = parseNumericValue(row[mapping.area]);
        const qty = parseNumericValue(row[cand]);

        const pa = pieceArea(len, wid);
        if (!isFinite(area) || area <= 0 || !pa || pa <= 0) continue;
        if (!isFinite(qty) || qty <= 0) continue;

        const expected = area / pa;
        if (!isFinite(expected) || expected <= 0) continue;

        const rel = Math.abs(qty - expected) / Math.max(1, expected);
        const intPenalty = Math.min(1, Math.abs(Math.round(expected) - expected));
        score += rel + 0.25 * intPenalty;
        used += 1;
      }
      if (used >= 3) scores.set(cand, score / used);
    }

    if (scores.size) {
      let best = mapping.qty;
      let bestScore = Infinity;
      for (const [cand, s] of scores.entries()) {
        if (s < bestScore) { bestScore = s; best = cand; }
      }
      mapping.qty = best;
    }
  }

  const rowData = sheetMatrix[rowIndex - 1] || [];
  return {
    rowIndex,
    name: rowData[mapping.name],
    qty: parseNumericValue(rowData[mapping.qty]),
    area_m2: parseNumericValue(rowData[mapping.area]),
    thickness: parseNumericValue(rowData[mapping.thickness]),
    length_mm: parseNumericValue(rowData[mapping.length]),
    width_mm: parseNumericValue(rowData[mapping.width]),
    __mapping: mapping,
    __headerRow: headerRow,
  };
}

function buildDetailBreakdown(workbook, leaves) {
  const details = [];
  const grouped = new Map();
  leaves.forEach((leaf) => {
    const key = `${leaf.sheet}!${leaf.row}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(leaf);
  });

  grouped.forEach((leafItems, key) => {
    const [sheetName, rowStr] = key.split('!');
    const rowIndex = Number(rowStr);
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    const context = inferDetailRowContext(matrix, rowIndex);
    const cost = leafItems.reduce((sum, item) => sum + (item.value || 0), 0);
    let area = context.area_m2;
    if (!area && context.length_mm && context.width_mm && context.qty) {
      area = (context.length_mm / 1000) * (context.width_mm / 1000) * context.qty;
    }
    details.push({
      name: context.name || `Строка ${rowIndex}`,
      qty: context.qty,
      area_m2: area,
      thickness: context.thickness,
      cost: round2(cost),
      rowIndex,
    });
  });
  return details;
}

function computeCoverage(leafSum, totalValue) {
  if (!totalValue) return null;
  return leafSum / totalValue;
}

function buildDspRates(details) {
  const byThickness = {};
  const totals = {};
  let totalArea = 0;
  let totalCost = 0;
  details.forEach((item) => {
    if (!item.area_m2 || !item.cost) return;
    const thicknessKey = item.thickness ? String(item.thickness) : null;
    if (thicknessKey) {
      totals[thicknessKey] = totals[thicknessKey] || { area: 0, cost: 0 };
      totals[thicknessKey].area += item.area_m2;
      totals[thicknessKey].cost += item.cost;
    }
    totalArea += item.area_m2;
    totalCost += item.cost;
  });
  Object.keys(totals).forEach((key) => {
    const area = totals[key].area;
    byThickness[key] = area > 0 ? totals[key].cost / area : null;
  });
  return {
    avgRate: totalArea > 0 ? totalCost / totalArea : null,
    byThickness,
  };
}

function readAnchorValue(workbook, ref) {
  if (!ref) return null;
  const parsed = readRefCell(workbook, ref);
  if (!parsed || !parsed.cell) return null;
  const value = parseNumericValue(parsed.cell.v);
  return Number.isFinite(value) ? value : null;
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


function inferCostTableFallback(workbook, sheetName, anchorValue, materials) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;

  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
  const norm = (v) => String(v || '').toLowerCase();

  let headerRow = -1;
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const rowText = row.map(norm).join(' | ');
    if (rowText.includes('длин') && rowText.includes('ширин') && rowText.includes('кол') && (rowText.includes('цен') || rowText.includes('стоим') || rowText.includes('руб'))) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) return null;

  const header = matrix[headerRow] || [];
  const findCol = (pred) => {
    for (let c = 0; c < header.length; c++) if (pred(norm(header[c]))) return c;
    return -1;
  };

  const nameCol = findCol((t) => t.includes('плита') || t.includes('наимен') || t.includes('детал'));
  const lenCol = findCol((t) => t.includes('длин'));
  const widCol = findCol((t) => t.includes('ширин'));
  const qtyCol = findCol((t) => t.includes('кол'));
  const areaCol = findCol((t) => t.includes('площ'));
  const costCol = findCol((t) => t.includes('цен') || t.includes('стоим') || t.includes('руб'));

  if (lenCol < 0 || widCol < 0 || qtyCol < 0 || costCol < 0) return null;

  const leaves = [];
  const items = [];
  let areaSum = 0;
  let leafSum = 0;

  for (let r = headerRow + 1; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const name = row[nameCol] || row[0] || '';
    const len = parseNumericValue(row[lenCol]);
    const wid = parseNumericValue(row[widCol]);
    const qty = parseNumericValue(row[qtyCol]);
    const cost = parseNumericValue(row[costCol]);
    const area = areaCol >= 0 ? parseNumericValue(row[areaCol]) : (len && wid && qty ? (len * wid * qty) / 1e6 : null);

    const emptyish = !String(name || '').trim() && !(len || wid || qty || cost);
    if (emptyish) {
      // stop after a few empties to avoid scanning the whole sheet
      break;
    }

    if (!(len && wid && qty)) continue;
    if (!cost || cost <= 0) continue;

    const excelRow = r + 1;
    const costRef = `${sheetName}!${XLSX.utils.encode_col(costCol)}${excelRow}`;
    leaves.push(costRef);

    const thicknessGuess = parseNumericValue(row[nameCol + 1]);
    const thicknessMm = thicknessGuess || 16;

    items.push({
      name: String(name || '').trim() || 'Деталь',
      thicknessMm,
      areaM2: area || 0,
      costRub: cost || 0,
      costRef,
    });

    areaSum += area || 0;
    leafSum += cost || 0;
  }

  if (!items.length || areaSum <= 0 || leafSum <= 0) return null;

  const avgRate = leafSum / areaSum;
  const coverage = anchorValue ? leafSum / anchorValue : null;

  return {
    usable: true,
    leafCount: leaves.length,
    leafRefs: leaves,
    leafSum,
    areaSum,
    coverage,
    items,
    costColsLetters: [XLSX.utils.encode_col(costCol)],
    notes: 'fallback: table scan (no formulas)',
    rates: {
      avgRate,
      thicknessRates: [{ thicknessMm: 16, rateRubPerM2: avgRate, areaM2: areaSum, costRub: leafSum }],
    },
  };
}


function buildCostTableBreakdown(workbook, anchorRef, materials, options = {}) {
  const anchorParts = decodeRefParts(anchorRef);
  if (!anchorParts) {
    return { anchorValue: 0, leaves: [], items: [], rates: { usable: false }, debug: { reason: 'bad anchor' } };
  }

  const ws = workbook.Sheets[anchorParts.sheetName];
  if (!ws) {
    return { anchorValue: 0, leaves: [], items: [], rates: { usable: false }, debug: { reason: 'sheet missing' } };
  }

  const anchorCell = ws[anchorParts.cellRef];
  const anchorValue = parseNumericValue(anchorCell?.v) || 0;
  const anchorFormula = getCellFormulaString(anchorCell);

  // anchorFormula (e.g. D92): "M66+O66+...+AU66"
  const totalsRefsRaw = anchorFormula ? extractRefsFromFormula(anchorFormula) : [];

  if (!totalsRefsRaw.length) {
    const fallback = inferCostTableFallback(workbook, anchorParts.sheetName, anchorValue, materials);
    if (fallback) {
      return {
        anchorValue,
        leaves: fallback.leafRefs,
        items: fallback.items,
        rates: fallback.rates,
        debug: {
          method: 'table-scan',
          ...fallback,
        },
      };
    }
  }
  const totalsFullRefs = expandRanges(totalsRefsRaw)
    .map((r) => toFullRef(r, anchorParts.sheetName))
    .filter(Boolean);

  // Determine which cost columns are actually used (non-zero totals). If none, keep all.
  const totalsWithValues = totalsFullRefs
    .map((fullRef) => {
      const p = decodeRefParts(fullRef, anchorParts.sheetName);
      const sheet = workbook.Sheets[p.sheetName];
      const cell = sheet?.[p.cellRef];
      return {
        fullRef,
        p,
        value: parseNumericValue(cell?.v),
        formula: getCellFormulaString(cell),
      };
    })
    .filter((x) => x.p);

  const nonZeroTotals = totalsWithValues.filter((x) => typeof x.value === 'number' && x.value > 0.000001);
  const activeTotals = (nonZeroTotals.length ? nonZeroTotals : totalsWithValues);

  const leafSet = new Set();
  const activeCostCols = new Set();

  activeTotals.forEach((t) => {
    if (!t.p) return;
    activeCostCols.add(t.p.col);

    const refs = extractRefsFromFormula(t.formula);
    const expanded = expandRanges(refs);
    // Keep leaves in the same column as the total cell (usually SUM(AA51:AA65))
    expanded.forEach((r) => {
      const leafFull = toFullRef(r, t.p.sheetName);
      if (!leafFull) return;
      const lp = decodeRefParts(leafFull, t.p.sheetName);
      if (!lp) return;
      if (lp.col !== t.p.col) return;
      if (lp.r === t.p.r) return; // don't include the total cell itself
      leafSet.add(leafFull);
    });
  });

  const leaves = Array.from(leafSet).sort((a, b) => {
    const pa = decodeRefParts(a, anchorParts.sheetName);
    const pb = decodeRefParts(b, anchorParts.sheetName);
    if (pa.sheetName !== pb.sheetName) return pa.sheetName.localeCompare(pb.sheetName);
    if (pa.c !== pb.c) return pa.c - pb.c;
    return pa.r - pb.r;
  });

  // Build "items" from the standard "Плитный материал" layout (name A, thickness B, len C, width D, qty I)
  const items = [];
  let leafSum = 0;
  let areaSum = 0;

  leaves.forEach((fullRef) => {
    const p = decodeRefParts(fullRef, anchorParts.sheetName);
    const sheet = workbook.Sheets[p.sheetName];
    const costCell = sheet?.[p.cellRef];
    const cost = parseNumericValue(costCell?.v);
    if (!(cost > 0)) return;

    const nameVal = sheet?.[XLSX.utils.encode_cell({ r: p.r, c: 0 })]?.v;
    const matIdVal = sheet?.[XLSX.utils.encode_cell({ r: p.r, c: 1 })]?.v;
    const lenVal = sheet?.[XLSX.utils.encode_cell({ r: p.r, c: 2 })]?.v;
    const widVal = sheet?.[XLSX.utils.encode_cell({ r: p.r, c: 3 })]?.v;
    const qtyVal = sheet?.[XLSX.utils.encode_cell({ r: p.r, c: 8 })]?.v;

    const areaCellRef = XLSX.utils.encode_cell({ r: p.r, c: Math.max(0, p.c - 1) });
    const areaVal = parseNumericValue(sheet?.[areaCellRef]?.v);

    const thicknessMm = extractThickness(matIdVal);
    const materialName = materials?.[String(matIdVal)]?.name || '';

    const qty = parseNumericValue(qtyVal) || 0;

    items.push({
      name: String(nameVal ?? `Строка ${p.r + 1}`),
      material: materialName,
      thickness_mm: thicknessMm || null,
      material_id: matIdVal != null ? String(matIdVal) : '',
      length_mm: parseNumericValue(lenVal) || null,
      width_mm: parseNumericValue(widVal) || null,
      qty: qty || null,
      area_m2: areaVal || null,
      cost_rub: cost,
      source: {
        costRef: `${p.sheetName}!${p.cellRef}`,
        areaRef: `${p.sheetName}!${areaCellRef}`,
        row: p.r + 1,
      },
    });

    leafSum += cost;
    if (areaVal > 0) areaSum += areaVal;
  });

  const byThickness = {};
  items.forEach((it) => {
    const key = it.thickness_mm ? `${Math.round(it.thickness_mm)}mm` : 'unknown';
    if (!byThickness[key]) byThickness[key] = { area: 0, cost: 0 };
    byThickness[key].area += it.area_m2 || 0;
    byThickness[key].cost += it.cost_rub || 0;
  });

  const rateByThickness = {};
  Object.keys(byThickness).forEach((k) => {
    const a = byThickness[k].area;
    const c = byThickness[k].cost;
    if (a > 0 && c > 0) rateByThickness[k] = c / a;
  });

  const avgRate = areaSum > 0 ? (leafSum / areaSum) : null;
  const coverage = anchorValue > 0 ? (leafSum / anchorValue) : 0;

  const usable = Boolean(avgRate && avgRate > 0 && coverage > 0.85 && coverage < 1.15);

  const debug = {
    anchorValue,
    leafCount: leaves.length,
    leafSum,
    areaSum,
    coverage,
    leafRefs: leaves.slice(0, 200).join(', '), // keep it short in UI
    costColsLetters: Array.from(activeCostCols).join(', ') || '—',
  };

  return {
    anchorValue,
    leaves,
    items,
    rates: { usable, avgRate, byThickness: rateByThickness },
    debug,
  };
}
function buildCalcSummary(workbook, anchors, materials) {
  const baseValues = {
    totalCost: anchors.totalCostRef ? readAnchorValue(workbook, anchors.totalCostRef) : 0,
    dsp: anchors.dspRef ? readAnchorValue(workbook, anchors.dspRef) : 0,
    edge: anchors.edgeRef ? readAnchorValue(workbook, anchors.edgeRef) : 0,
    plastic: anchors.plasticRef ? readAnchorValue(workbook, anchors.plasticRef) : 0,
    fabric: anchors.fabricRef ? readAnchorValue(workbook, anchors.fabricRef) : 0,
    hwImported: anchors.hwImpRef ? readAnchorValue(workbook, anchors.hwImpRef) : 0,
    hwReplaced: anchors.hwRepRef ? readAnchorValue(workbook, anchors.hwRepRef) : 0,
    pack: anchors.packRef ? readAnchorValue(workbook, anchors.packRef) : 0,
    labor: anchors.laborRef ? readAnchorValue(workbook, anchors.laborRef) : 0,
    laborHours: anchors.laborHoursRef ? readAnchorValue(workbook, anchors.laborHoursRef) : null,
  };

  const summary = {
    anchors,
    baseValues,
    breakdown: {},
    rates: {},
    debug: {},
  };

  if (anchors.dspRef) {
    const dsp = buildCostTableBreakdown(workbook, anchors.dspRef, materials);
    summary.breakdown.dsp = {
      usable: dsp.rates.usable,
      // store leaf row items (so we can inspect what was counted)
      items: dsp.items,
      leaves: dsp.leaves,
    };
    summary.rates.dsp = dsp.rates;
    summary.debug.dsp = dsp.debug;
  }

  if (anchors.edgeRef) {
    const edge = buildCostTableBreakdown(workbook, anchors.edgeRef, materials);
    summary.breakdown.edge = {
      usable: edge.rates.usable,
      items: edge.items,
      leaves: edge.leaves,
    };
    summary.rates.edge = edge.rates;
    summary.debug.edge = edge.debug;
  }

  return summary;
}

function parseDimensions(sheet, mapping) {
  let width;
  let depth;
  let height;

  if (mapping.dimensionsCell) {
    const value = String(readCellValue(sheet, mapping.dimensionsCell) || '').trim();
    const match = value.match(/(\d{3,4})\s*[*xх×]\s*(\d{3,4})\s*[*xх×]\s*(\d{3,4})/i);
    if (match) {
      width = Number(match[1]);
      depth = Number(match[2]);
      height = Number(match[3]);
    }
  }

  if (!width && mapping.widthCell) {
    width = readCellNumber(sheet, mapping.widthCell);
  }
  if (!depth && mapping.depthCell) {
    depth = readCellNumber(sheet, mapping.depthCell);
  }
  if (!height && mapping.heightCell) {
    height = readCellNumber(sheet, mapping.heightCell);
  }

  return { width: width || null, depth: depth || null, height: height || null };
}

function detectBaseCostCell(sheet) {
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const keywords = ['прямые затраты', 'итого', 'стоимость'];
  for (let rowIndex = 0; rowIndex < json.length; rowIndex += 1) {
    const row = json[rowIndex] || [];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cellText = String(row[colIndex] || '').toLowerCase();
      if (!keywords.some((keyword) => cellText.includes(keyword))) continue;
      const candidate = parseNumericValue(row[colIndex + 1]);
      if (Number.isFinite(candidate)) {
        return `${colIndexToLetter(colIndex + 1)}${rowIndex + 1}`;
      }
      for (let next = colIndex + 1; next < row.length; next += 1) {
        const alt = parseNumericValue(row[next]);
        if (Number.isFinite(alt)) {
          return `${colIndexToLetter(next)}${rowIndex + 1}`;
        }
      }
    }
  }
  return '';
}

function parseMaterialDictionary(sheet, mapping) {
  const materials = {};
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  for (let row = mapping.materialDictStart; row <= mapping.materialDictEnd; row += 1) {
    const rowData = json[row - 1];
    if (!rowData) continue;
    const name = rowData[mapping.materialNameCol];
    const price = Number(rowData[mapping.materialPriceCol]);
    const waste = Number(rowData[mapping.materialWasteCol] || 0);
    const id = String(rowData[mapping.materialIdCol] || '').trim();
    if (!id || !name) continue;
    materials[id] = { name, price: Number.isFinite(price) ? price : 0, waste };
  }
  return materials;
}

function parseCorpusDetails(sheet, mapping, materials) {
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const parts = [];
  for (let row = mapping.detailsStartRow; row <= mapping.detailsEndRow; row += 1) {
    const rowData = json[row - 1];
    if (!rowData) continue;
    const name = rowData[mapping.detailsNameCol];
    if (!name) continue;
    const materialId = String(rowData[mapping.detailsThicknessCol] || '').trim();
    const length = Number(rowData[mapping.detailsLengthCol]);
    const width = Number(rowData[mapping.detailsWidthCol]);
    const qty = Number(rowData[mapping.detailsQtyCol] || 1);
    const material = materials[materialId]?.name || '';

    parts.push({
      name: String(name),
      material_id: materialId,
      material,
      length_mm: length,
      width_mm: width,
      thickness: extractThickness(materialId) || 16,
      qty: Number.isFinite(qty) ? qty : 1,
    });
  }
  return parts;
}

function parseFurniture(sheet, mapping) {
  if (!sheet) return [];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const items = [];
  for (let row = mapping.furnitureHeaderRow + 1; row < mapping.furnitureHeaderRow + 40; row += 1) {
    const rowData = json[row - 1];
    if (!rowData) continue;
    const code = rowData[mapping.furnitureCodeCol];
    const name = rowData[mapping.furnitureNameCol];
    if (!code && !name) continue;
    items.push({
      code,
      name,
      qty: Number(rowData[mapping.furnitureQtyCol] || 0),
      unit: rowData[mapping.furnitureUnitCol] || 'шт',
      price: Number(rowData[mapping.furniturePriceCol] || 0),
    });
  }
  return items;
}

function extractThickness(value) {
  const str = String(value || '');
  const match = str.match(/(\d{2})/);
  return match ? Number(match[1]) : null;
}

function autoDetectMapping(sheet) {
  const mapping = {};
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  for (let row = 0; row < 50; row += 1) {
    const rowData = json[row];
    if (!rowData) continue;
    const nameA = String(rowData[0] || '').toLowerCase();
    const valueF = rowData[5];
    if (nameA.includes('лдсп') || nameA.includes('мдф') || nameA.includes('дсп')) {
      if (!Number.isNaN(Number(valueF))) {
        mapping.materialDictStart = row + 1;
        mapping.materialDictEnd = row + 10;
        mapping.materialNameCol = 0;
        mapping.materialPriceCol = 1;
        mapping.materialWasteCol = 2;
        mapping.materialIdCol = 5;
        break;
      }
    }
  }

  for (let row = 0; row < 100; row += 1) {
    const cellA = String(json[row]?.[0] || '');
    if (/\d{3,4}\s*[*хx×]\s*\d{3,4}\s*[*хx×]\s*\d{3,4}/.test(cellA)) {
      mapping.dimensionsCell = `A${row + 1}`;
      break;
    }
  }

  for (let row = 0; row < 100; row += 1) {
    const rowData = json[row];
    const rowText = rowData?.map((c) => String(c || '').toLowerCase()).join(' ');
    if (rowText?.includes('наимен') && rowText?.includes('длина') && rowText?.includes('ширина')) {
      mapping.detailsHeaderRow = row + 1;
      mapping.detailsStartRow = row + 2;
      rowData.forEach((cell, idx) => {
        const cellLow = String(cell || '').toLowerCase();
        if (cellLow.includes('наимен')) mapping.detailsNameCol = idx;
        if (cellLow.includes('тлщн')) mapping.detailsThicknessCol = idx;
        if (cellLow.includes('длин')) mapping.detailsLengthCol = idx;
        if (cellLow.includes('ширин')) mapping.detailsWidthCol = idx;
        if (cellLow.includes('кол')) mapping.detailsQtyCol = idx;
      });
      mapping.detailsEndRow = mapping.detailsStartRow + 30;
      break;
    }
  }

  return mapping;
}

function autoDetectFurnitureMapping(sheet) {
  const mapping = {};
  if (!sheet) return mapping;
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  for (let row = 0; row < 50; row += 1) {
    const rowData = json[row];
    const rowText = rowData?.map((cell) => String(cell || '').toLowerCase()).join(' ') || '';
    if (rowText.includes('код') && rowText.includes('наимен') && rowText.includes('кол')) {
      mapping.furnitureHeaderRow = row + 1;
      rowData.forEach((cell, idx) => {
        const cellLow = String(cell || '').toLowerCase();
        if (cellLow.includes('код')) mapping.furnitureCodeCol = idx;
        if (cellLow.includes('наимен')) mapping.furnitureNameCol = idx;
        if (cellLow.includes('кол')) mapping.furnitureQtyCol = idx;
        if (cellLow.includes('ед')) mapping.furnitureUnitCol = idx;
        if (cellLow.includes('цен')) mapping.furniturePriceCol = idx;
      });
      break;
    }
  }

  return mapping;
}

function applyMappingToUI(mapping) {
  if (!mapping) return;
  if (mapping.materialDictStart) document.getElementById('mat-start').value = mapping.materialDictStart;
  if (mapping.materialDictEnd) document.getElementById('mat-end').value = mapping.materialDictEnd;
  if (Number.isInteger(mapping.materialNameCol)) document.getElementById('mat-name-col').value = colIndexToLetter(mapping.materialNameCol);
  if (Number.isInteger(mapping.materialPriceCol)) document.getElementById('mat-price-col').value = colIndexToLetter(mapping.materialPriceCol);
  if (Number.isInteger(mapping.materialWasteCol)) document.getElementById('mat-waste-col').value = colIndexToLetter(mapping.materialWasteCol);
  if (Number.isInteger(mapping.materialIdCol)) document.getElementById('mat-id-col').value = colIndexToLetter(mapping.materialIdCol);
  if (mapping.dimensionsCell) document.getElementById('dims-cell').value = normalizeCellRef(mapping.dimensionsCell);
  if (mapping.widthCell) document.getElementById('dims-width').value = normalizeCellRef(mapping.widthCell);
  if (mapping.depthCell) document.getElementById('dims-depth').value = normalizeCellRef(mapping.depthCell);
  if (mapping.heightCell) document.getElementById('dims-height').value = normalizeCellRef(mapping.heightCell);
  if (mapping.detailsHeaderRow) document.getElementById('details-header').value = mapping.detailsHeaderRow;
  if (mapping.detailsStartRow) document.getElementById('details-start').value = mapping.detailsStartRow;
  if (mapping.detailsEndRow) document.getElementById('details-end').value = mapping.detailsEndRow;
  if (Number.isInteger(mapping.detailsNameCol)) document.getElementById('details-name-col').value = colIndexToLetter(mapping.detailsNameCol);
  if (Number.isInteger(mapping.detailsThicknessCol)) document.getElementById('details-thickness-col').value = colIndexToLetter(mapping.detailsThicknessCol);
  if (Number.isInteger(mapping.detailsLengthCol)) document.getElementById('details-length-col').value = colIndexToLetter(mapping.detailsLengthCol);
  if (Number.isInteger(mapping.detailsWidthCol)) document.getElementById('details-width-col').value = colIndexToLetter(mapping.detailsWidthCol);
  if (Number.isInteger(mapping.detailsQtyCol)) document.getElementById('details-qty-col').value = colIndexToLetter(mapping.detailsQtyCol);
  if (mapping.furnitureSheet !== undefined) document.getElementById('furniture-sheet').value = mapping.furnitureSheet;
  if (mapping.furnitureHeaderRow) document.getElementById('furniture-header').value = mapping.furnitureHeaderRow;
  if (Number.isInteger(mapping.furnitureCodeCol)) document.getElementById('furniture-code-col').value = colIndexToLetter(mapping.furnitureCodeCol);
  if (Number.isInteger(mapping.furnitureQtyCol)) document.getElementById('furniture-qty-col').value = colIndexToLetter(mapping.furnitureQtyCol);
  if (Number.isInteger(mapping.furnitureNameCol)) document.getElementById('furniture-name-col').value = colIndexToLetter(mapping.furnitureNameCol);
  if (Number.isInteger(mapping.furnitureUnitCol)) document.getElementById('furniture-unit-col').value = colIndexToLetter(mapping.furnitureUnitCol);
  if (Number.isInteger(mapping.furniturePriceCol)) document.getElementById('furniture-price-col').value = colIndexToLetter(mapping.furniturePriceCol);
  if (mapping.baseCostCell) document.getElementById('base-cost-cell').value = normalizeCellRef(mapping.baseCostCell);
  if (mapping.anchorOverrides) {
    const overrides = mapping.anchorOverrides;
    if (overrides.weightRef) document.getElementById('anchor-weight').value = overrides.weightRef;
    if (overrides.laborHoursRef) document.getElementById('anchor-labor-hours').value = overrides.laborHoursRef;
    if (overrides.dspRef) document.getElementById('anchor-dsp').value = overrides.dspRef;
    if (overrides.edgeRef) document.getElementById('anchor-edge').value = overrides.edgeRef;
    if (overrides.plasticRef) document.getElementById('anchor-plastic').value = overrides.plasticRef;
    if (overrides.fabricRef) document.getElementById('anchor-fabric').value = overrides.fabricRef;
    if (overrides.hwImpRef) document.getElementById('anchor-hw-imp').value = overrides.hwImpRef;
    if (overrides.hwRepRef) document.getElementById('anchor-hw-rep').value = overrides.hwRepRef;
    if (overrides.packRef) document.getElementById('anchor-pack').value = overrides.packRef;
    if (overrides.laborRef) document.getElementById('anchor-labor').value = overrides.laborRef;
    if (overrides.totalCostRef) document.getElementById('anchor-total').value = overrides.totalCostRef;
  }
}

function saveTemplate(name, mapping) {
  const templates = JSON.parse(localStorage.getItem('mapping-templates') || '{}');
  templates[name] = mapping;
  localStorage.setItem('mapping-templates', JSON.stringify(templates));
  renderTemplateOptions();
}

function loadTemplate(name) {
  const templates = JSON.parse(localStorage.getItem('mapping-templates') || '{}');
  return templates[name];
}

function renderTemplateOptions() {
  const select = document.getElementById('template-select');
  const templates = JSON.parse(localStorage.getItem('mapping-templates') || '{}');
  select.innerHTML = '<option value="">Выбрать шаблон...</option>';
  Object.keys(templates).forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });
}

function renderCalcSummaryAnchors(anchors) {
  const list = document.getElementById('calc-summary-list');
  const warning = document.getElementById('calc-summary-warning');
  if (!list || !warning) return;
  list.innerHTML = '';
  warning.innerHTML = '';

  const labels = [
    ['weightRef', 'Вес (кг)', 'anchor-weight'],
    ['laborHoursRef', 'Трудоемкость', 'anchor-labor-hours'],
    ['dspRef', 'Стоимость ДСП', 'anchor-dsp'],
    ['edgeRef', 'Стоимость кромки', 'anchor-edge'],
    ['plasticRef', 'Стоимость пластика', 'anchor-plastic'],
    ['fabricRef', 'Стоимость ткани', 'anchor-fabric'],
    ['hwImpRef', 'Фурнитура имп.', 'anchor-hw-imp'],
    ['hwRepRef', 'Фурнитура отч.', 'anchor-hw-rep'],
    ['packRef', 'Стоимость упаковки', 'anchor-pack'],
    ['laborRef', 'Труд рабочих', 'anchor-labor'],
    ['totalCostRef', 'Прямые затраты', 'anchor-total'],
  ];

  const missing = [];
  labels.forEach(([key, label, inputId]) => {
    const li = document.createElement('li');
    const ref = anchors?.[key];
    li.textContent = `${label}: ${ref || 'не найдено'}`;
    list.appendChild(li);
    const input = document.getElementById(inputId);
    if (input) {
      if (ref) {
        input.value = ref;
        input.setAttribute('disabled', 'disabled');
      } else {
        input.value = '';
        input.removeAttribute('disabled');
        missing.push(label);
      }
    }
  });

  const baseCostInput = document.getElementById('base-cost-cell');
  if (baseCostInput) {
    if (anchors?.totalCostRef) {
      baseCostInput.value = anchors.totalCostRef;
      baseCostInput.setAttribute('disabled', 'disabled');
    } else {
      baseCostInput.removeAttribute('disabled');
    }
  }

  if (missing.length) {
    const warningItem = document.createElement('div');
    warningItem.textContent = `⚠️ Не найдено автоматически: ${missing.join(', ')}. Можно указать вручную.`;
    warning.appendChild(warningItem);
  } else {
    const okItem = document.createElement('div');
    okItem.textContent = '✅ Итоги калькуляции найдены автоматически.';
    warning.appendChild(okItem);
  }
}

function collectMapping() {
  return {
    materialDictStart: Number(document.getElementById('mat-start').value),
    materialDictEnd: Number(document.getElementById('mat-end').value),
    materialNameCol: letterToColIndex(document.getElementById('mat-name-col').value),
    materialPriceCol: letterToColIndex(document.getElementById('mat-price-col').value),
    materialWasteCol: letterToColIndex(document.getElementById('mat-waste-col').value),
    materialIdCol: letterToColIndex(document.getElementById('mat-id-col').value),
    dimensionsCell: normalizeCellRef(document.getElementById('dims-cell').value),
    widthCell: normalizeCellRef(document.getElementById('dims-width').value),
    depthCell: normalizeCellRef(document.getElementById('dims-depth').value),
    heightCell: normalizeCellRef(document.getElementById('dims-height').value),
    detailsHeaderRow: Number(document.getElementById('details-header').value),
    detailsStartRow: Number(document.getElementById('details-start').value),
    detailsEndRow: Number(document.getElementById('details-end').value),
    detailsNameCol: letterToColIndex(document.getElementById('details-name-col').value),
    detailsThicknessCol: letterToColIndex(document.getElementById('details-thickness-col').value),
    detailsLengthCol: letterToColIndex(document.getElementById('details-length-col').value),
    detailsWidthCol: letterToColIndex(document.getElementById('details-width-col').value),
    detailsQtyCol: letterToColIndex(document.getElementById('details-qty-col').value),
    furnitureSheet: document.getElementById('furniture-sheet').value,
    furnitureHeaderRow: Number(document.getElementById('furniture-header').value),
    furnitureCodeCol: letterToColIndex(document.getElementById('furniture-code-col').value),
    furnitureQtyCol: letterToColIndex(document.getElementById('furniture-qty-col').value),
    furnitureNameCol: letterToColIndex(document.getElementById('furniture-name-col').value),
    furnitureUnitCol: letterToColIndex(document.getElementById('furniture-unit-col').value),
    furniturePriceCol: letterToColIndex(document.getElementById('furniture-price-col').value),
    baseCostCell: normalizeCellRef(document.getElementById('base-cost-cell').value),
    anchorOverrides: {
      weightRef: normalizeAnchorRef(document.getElementById('anchor-weight').value),
      laborHoursRef: normalizeAnchorRef(document.getElementById('anchor-labor-hours').value),
      dspRef: normalizeAnchorRef(document.getElementById('anchor-dsp').value),
      edgeRef: normalizeAnchorRef(document.getElementById('anchor-edge').value),
      plasticRef: normalizeAnchorRef(document.getElementById('anchor-plastic').value),
      fabricRef: normalizeAnchorRef(document.getElementById('anchor-fabric').value),
      hwImpRef: normalizeAnchorRef(document.getElementById('anchor-hw-imp').value),
      hwRepRef: normalizeAnchorRef(document.getElementById('anchor-hw-rep').value),
      packRef: normalizeAnchorRef(document.getElementById('anchor-pack').value),
      laborRef: normalizeAnchorRef(document.getElementById('anchor-labor').value),
      totalCostRef: normalizeAnchorRef(document.getElementById('anchor-total').value),
    },
  };
}

function resolveAnchors(autoAnchors, overrides, defaultSheet = null) {
  const resolved = { ...(autoAnchors || {}) };
  if (!overrides) return resolved;

  Object.keys(overrides).forEach((key) => {
    const raw = String(overrides[key] ?? '').trim();
    if (!raw) return;

    // Only fill missing anchors (manual fields are disabled when auto is present)
    if (resolved[key]) return;

    const full = toFullRef(raw, defaultSheet || parseRef(resolved.totalCostRef || resolved.dspRef || resolved.edgeRef || '')?.sheetName);
    resolved[key] = full || raw;
  });

  return resolved;
}

function parseExcelWithMapping(workbook, mapping) {
  const sheet = workbook.Sheets[state.activeSheet];
  const materials = parseMaterialDictionary(sheet, mapping);
  const corpus = parseCorpusDetails(sheet, mapping, materials);
  const furnitureSheet = workbook.Sheets[mapping.furnitureSheet];
  const furniture = parseFurniture(furnitureSheet, mapping);
  const dims = parseDimensions(sheet, mapping);
  const anchors = resolveAnchors(state.calcSummary?.anchors, mapping.anchorOverrides);
  const calcSummary = buildCalcSummary(workbook, anchors, materials);
  const baseCostFromAnchors = calcSummary.baseValues.totalCost;
  const baseCost = Number.isFinite(baseCostFromAnchors)
    ? baseCostFromAnchors
    : (mapping.baseCostCell ? readCellNumber(sheet, mapping.baseCostCell) : null);
  const baseMaterialCost = calcSummary.baseValues.dsp || calculatePrice(corpus, materials);

  return {
    dims,
    corpus,
    furniture,
    materials,
    baseCost: Number.isFinite(baseCost) ? baseCost : null,
    baseMaterialCost,
    calcSummary,
  };
}

function formatNumber(value, unit = '') {
  if (value === undefined || value === null || Number.isNaN(value)) return '—';
  return `${Number(value).toLocaleString('ru-RU')} ${unit}`.trim();
}

function formatDimensions(dims) {
  if (!dims || !dims.width || !dims.depth || !dims.height) return '—';
  return `${dims.width}×${dims.depth}×${dims.height}`;
}

function getMaterialDensity(materialName) {
  const mat = String(materialName || '').toLowerCase();
  if (mat.includes('лдсп') || mat.includes('дсп')) return 730;
  if (mat.includes('мдф')) return 750;
  if (mat.includes('фанер')) return 650;
  if (mat.includes('двп') || mat.includes('оргалит')) return 850;
  if (mat.includes('стекл')) return 2500;
  return MATERIAL_DENSITY;
}

function calculateWeight(parts) {
  let totalKg = 0;
  parts.forEach((part) => {
    if (!part.thickness || !part.length_mm || !part.width_mm || !part.qty) return;
    const density = getMaterialDensity(part.material);
    const areaM2 = (part.length_mm / 1000) * (part.width_mm / 1000);
    const thicknessM = part.thickness / 1000;
    totalKg += density * areaM2 * thicknessM * part.qty;
  });
  return Math.round(totalKg * 100) / 100;
}


function calculatePrice(parts, materials) {
  let total = 0;
  parts.forEach((part) => {
    if (!part.length_mm || !part.width_mm || !part.qty) return;
    const areaM2 = (part.length_mm / 1000) * (part.width_mm / 1000) * part.qty;
    const material = materials[part.material_id];
    if (material) {
      const wasteFactor = 1 + (material.waste || 0) / 100;
      total += areaM2 * material.price * wasteFactor;
    }
  });
  return Math.round(total);
}


function calculateFurnitureCost(furniture) {
  return (furniture || []).reduce((sum, item) => {
    const price = Number(item.price || 0);
    if (!price || item.unit === '%') return sum;
    return sum + Number(item.qty || 0) * price;
  }, 0);
}

function inferPartType(name) {
  const n = (name || '').toLowerCase();
  if (n.includes('бок')) return 'side';
  if (n.includes('дно') || n.includes('крыш')) return 'base';
  if (n.includes('зад') || n.includes('двп')) return 'back';
  if (n.includes('перегород')) return 'partition';
  if (n.includes('фасад') || n.includes('двер')) return 'facade';
  if (n.includes('полк')) return 'shelf';
  if (n.includes('ящик')) return 'drawer';
  if (n.includes('штанг')) return 'rod';
  return 'other';
}

function splitSections(totalWidth) {
  const minSections = totalWidth >= PARTITION_THRESHOLD ? 2 : 1;
  if (totalWidth <= MAX_SECTION_WIDTH && minSections === 1) {
    return [totalWidth];
  }
  const requiredSections = Math.ceil(totalWidth / MAX_SECTION_WIDTH);
  const numSections = Math.max(requiredSections, minSections);
  const baseWidth = Math.floor(totalWidth / numSections);
  const remainder = totalWidth - baseWidth * numSections;
  const sections = Array(numSections).fill(baseWidth);
  for (let i = 0; i < remainder; i += 1) {
    sections[i] += 1;
  }
  return sections;
}

function inferSectionCount(spec) {
  const backs = spec.corpus.filter((p) => inferPartType(p.name) === 'back' && (p.qty || 0) > 0);
  if (backs.length) {
    // Back panels in these templates are typically listed per module and split by height into 2+ rows
    // with the same qty. Take the most common qty (mode).
    const qtys = backs.map((p) => Math.round(p.qty || 0)).filter((n) => n > 0);
    const freq = new Map();
    qtys.forEach((n) => freq.set(n, (freq.get(n) || 0) + 1));
    let best = 0;
    let bestCnt = -1;
    for (const [n, c] of freq.entries()) {
      if (c > bestCnt) {
        best = n;
        bestCnt = c;
      }
    }
    if (best > 0) return best;
    return Math.max(1, Math.round(Math.max(...qtys)));
  }

  const partitionsQty = spec.corpus.reduce((sum, part) => {
    return inferPartType(part.name) === 'partition' ? sum + (part.qty || 0) : sum;
  }, 0);
  if (partitionsQty > 0) {
    return Math.max(1, Math.round(partitionsQty) + 1);
  }

  const fallback = splitSections(spec.dims.width || 0).length;
  return fallback || 1;
}

function getBaseStructure(spec) {
  const sections = inferSectionCount(spec);
  const shelves = spec.corpus.reduce((sum, part) => {
    return inferPartType(part.name) === 'shelf' ? sum + (part.qty || 0) : sum;
  }, 0);
  return {
    sections,
    partitions: Math.max(sections - 1, 0),
    shelves,
  };
}
function getBasePriceFromSpec(spec) {
  if (spec.baseCost !== null && spec.baseCost !== undefined) {
    return spec.baseCost;
  }
  return calculatePrice(spec.corpus, spec.materials || {});
}

function renderBaseSummary(spec) {
  const baseWeight = calculateWeight(spec.corpus);
  const basePrice = getBasePriceFromSpec(spec);
  document.getElementById('current-dims').textContent = formatDimensions(spec.dims);
  document.getElementById('current-weight').textContent = formatNumber(baseWeight, 'кг');
  document.getElementById('current-price').textContent = formatNumber(basePrice, '₽');
  const widthInput = document.getElementById('new-width');
  const depthInput = document.getElementById('new-depth');
  const heightInput = document.getElementById('new-height');
  if (!widthInput.value) widthInput.value = spec.dims.width || '';
  if (!depthInput.value) depthInput.value = spec.dims.depth || '';
  if (!heightInput.value) heightInput.value = spec.dims.height || '';
}

function renderValidationSummary(spec) {
  const baseMaterialsCost = spec.baseMaterialCost || calculatePrice(spec.corpus, spec.materials || {});
  const baseHardwareCostFallback = calculateFurnitureCost(spec.furniture || []);
  const baseValues = spec.calcSummary?.baseValues || {};
  const baseCost = spec.baseCost;
  const baseHardwareCost = Number.isFinite(baseValues.hwImp) || Number.isFinite(baseValues.hwRep)
    ? (Number(baseValues.hwImp || 0) + Number(baseValues.hwRep || 0))
    : baseHardwareCostFallback;
  const baseOther = baseCost !== null && baseCost !== undefined
    ? baseCost - (baseMaterialsCost + baseHardwareCost)
    : null;
  document.getElementById('validation-base-cost').textContent = formatNumber(baseCost, '₽');
  document.getElementById('validation-base-materials').textContent = formatNumber(baseMaterialsCost, '₽');
  document.getElementById('validation-base-hardware').textContent = formatNumber(baseHardwareCost, '₽');
  document.getElementById('validation-base-other').textContent = formatNumber(baseOther, '₽');

  const warningBox = document.getElementById('validation-warning');
  warningBox.innerHTML = '';
  if (baseCost !== null && baseCost !== undefined) {
    const other = Number(baseOther);
    if (other < 0 || other > baseCost * 0.7) {
      const warning = document.createElement('div');
      warning.textContent = '⚠️ проверь маппинг или в КП есть услуги/наценки';
      warningBox.appendChild(warning);
    }
  }
  const dspCoverage = spec.calcSummary?.breakdown?.dsp?.coverage;
  if (dspCoverage !== null && dspCoverage !== undefined && dspCoverage < 0.995) {
    const warning = document.createElement('div');
    warning.textContent = `⚠️ Разбор ДСП покрывает ${(dspCoverage * 100).toFixed(1)}% — используется fallback.`;
    warningBox.appendChild(warning);
  }
}

function renderResultsTable(type, spec) {
  const table = document.getElementById('results-table');
  const calcWrap = document.getElementById('calc-breakdown');
  if (calcWrap) {
    calcWrap.classList.toggle('hidden', type !== 'calc');
  }
  if (!table) return;
  table.classList.toggle('hidden', type === 'calc');
  table.innerHTML = '';
  if (type === 'calc') {
    renderCalcBreakdown(spec);
    return;
  }
  if (type === 'furniture') {
    const headers = ['Код', 'Наименование', 'Кол-во', 'Ед.', 'Цена ₽'];
    const headerRow = document.createElement('tr');
    headers.forEach((text) => {
      const th = document.createElement('th');
      th.textContent = text;
      headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    if (!spec.furniture || spec.furniture.length === 0) {
      const emptyRow = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = headers.length;
      td.textContent = 'Фурнитура не найдена.';
      emptyRow.appendChild(td);
      table.appendChild(emptyRow);
      return;
    }

    spec.furniture.forEach((item) => {
      const tr = document.createElement('tr');
      [
        item.code,
        item.name,
        item.qty,
        item.unit,
        item.price,
      ].forEach((value) => {
        const td = document.createElement('td');
        td.textContent = value ?? '';
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    return;
  }

  const headers = ['Наименование', 'Материал', 'Длина', 'Ширина', 'Толщина', 'Кол-во'];
  const headerRow = document.createElement('tr');
  headers.forEach((text) => {
    const th = document.createElement('th');
    th.textContent = text;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  spec.corpus.forEach((part) => {
    const tr = document.createElement('tr');
    [
      part.name,
      part.material,
      part.length_mm,
      part.width_mm,
      part.thickness,
      part.qty,
    ].forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
}

function setActiveResultsTab(type) {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('active', tab.dataset.tab === type);
  });
  state.activeResultsTab = type;
}

function renderCalcBreakdown(spec) {
  const table = document.getElementById('calc-breakdown-table');
  const leafCount = document.getElementById('calc-leaf-count');
  const coverage = document.getElementById('calc-coverage');
  const leafSum = document.getElementById('calc-leaf-sum');
  const totalValue = document.getElementById('calc-total');
  if (!table) return;
  table.innerHTML = '';
  const breakdown = spec.calcSummary?.breakdown?.dsp;
  if (!breakdown || !breakdown.details || breakdown.details.length === 0) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 4;
    cell.textContent = 'Разбор ДСП не найден.';
    row.appendChild(cell);
    table.appendChild(row);
    if (leafCount) leafCount.textContent = '—';
    if (coverage) coverage.textContent = '—';
    if (leafSum) leafSum.textContent = '—';
    if (totalValue) totalValue.textContent = '—';
    return;
  }
  const headers = ['Деталь', 'Qty', 'Area (м²)', 'Cost (₽)'];
  const headerRow = document.createElement('tr');
  headers.forEach((text) => {
    const th = document.createElement('th');
    th.textContent = text;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  breakdown.details.forEach((detail) => {
    const tr = document.createElement('tr');
    [detail.name, detail.qty ?? '', detail.area_m2 ? round2(detail.area_m2) : '', detail.cost ?? '']
      .forEach((value) => {
        const td = document.createElement('td');
        td.textContent = value;
        tr.appendChild(td);
      });
    table.appendChild(tr);
  });

  if (leafCount) leafCount.textContent = formatNumber(breakdown.leafCount);
  if (coverage) coverage.textContent = breakdown.coverage ? `${round2(breakdown.coverage * 100)}%` : '—';
  if (leafSum) leafSum.textContent = formatNumber(breakdown.leafSum, '₽');
  if (totalValue) totalValue.textContent = formatNumber(breakdown.totalValue, '₽');
}

function renderMaterialSpecOptions(spec) {
  const select = document.getElementById('material-spec-select');
  if (!select) return;
  const materials = Object.values(spec.materials || {})
    .map((mat) => mat.name)
    .filter(Boolean);
  const unique = [...new Set(materials)];
  select.innerHTML = '<option value="">Выберите материал</option>';
  unique.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });
}

function renderResults(spec, weight, price, warnings, breakdown) {
  if (state.originalSpec) {
    renderBaseSummary(state.originalSpec);
  }

  document.getElementById('new-dims').textContent = formatDimensions(spec.dims);
  document.getElementById('new-weight').textContent = formatNumber(weight, 'кг');
  document.getElementById('new-price').textContent = formatNumber(price, '₽');
  document.getElementById('price-materials').textContent = formatNumber(breakdown?.materials, '₽');
  document.getElementById('price-hardware').textContent = formatNumber(breakdown?.hardware, '₽');
  document.getElementById('price-other').textContent = formatNumber(breakdown?.other, '₽');
  document.getElementById('price-total').textContent = formatNumber(breakdown?.total ?? price, '₽');

  const warningsBox = document.getElementById('warnings');
  warningsBox.innerHTML = '';
  warnings.forEach((warning) => {
    const item = document.createElement('div');
    item.textContent = warning;
    warningsBox.appendChild(item);
  });

  const activeTab = state.activeResultsTab || document.querySelector('.tab.active')?.dataset.tab || 'corpus';
  setActiveResultsTab(activeTab);
  renderResultsTable(activeTab, spec);
  renderMaterialSpecOptions(spec);

  document.getElementById('results-card').classList.remove('hidden');
}

function exportToExcel(spec) {
  const corpusSheet = XLSX.utils.json_to_sheet(spec.corpus);
  const furnitureSheet = XLSX.utils.json_to_sheet(spec.furniture);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, corpusSheet, 'Корпус');
  XLSX.utils.book_append_sheet(workbook, furnitureSheet, 'Фурнитура');
  XLSX.writeFile(workbook, 'wardrobe_result.xlsx');
}

function updateMappingFromAuto(type) {
  if (!state.workbook || !state.activeSheet) {
    alert('Сначала загрузите Excel-файл.');
    return;
  }
  const sheet = state.workbook.Sheets[state.activeSheet];
  const mapping = autoDetectMapping(sheet);
  if (!mapping || Object.keys(mapping).length === 0) {
    if (type !== 'furniture') {
      alert('Не удалось найти структуру. Проверьте выбранный лист и попробуйте снова.');
      return;
    }
  }
  if (type === 'materials') {
    applyMappingToUI({
      materialDictStart: mapping.materialDictStart,
      materialDictEnd: mapping.materialDictEnd,
      materialNameCol: mapping.materialNameCol,
      materialPriceCol: mapping.materialPriceCol,
      materialWasteCol: mapping.materialWasteCol,
      materialIdCol: mapping.materialIdCol,
    });
  }
  if (type === 'dimensions') {
    applyMappingToUI({
      dimensionsCell: mapping.dimensionsCell,
      widthCell: mapping.widthCell,
      depthCell: mapping.depthCell,
      heightCell: mapping.heightCell,
    });
  }
  if (type === 'details') {
    applyMappingToUI({
      detailsHeaderRow: mapping.detailsHeaderRow,
      detailsStartRow: mapping.detailsStartRow,
      detailsEndRow: mapping.detailsEndRow,
      detailsNameCol: mapping.detailsNameCol,
      detailsThicknessCol: mapping.detailsThicknessCol,
      detailsLengthCol: mapping.detailsLengthCol,
      detailsWidthCol: mapping.detailsWidthCol,
      detailsQtyCol: mapping.detailsQtyCol,
    });
  }
  if (type === 'furniture') {
    const selectedSheet = document.getElementById('furniture-sheet').value || state.activeSheet;
    const furnitureSheet = state.workbook.Sheets[selectedSheet];
    const furnitureMapping = autoDetectFurnitureMapping(furnitureSheet);
    if (!furnitureMapping || Object.keys(furnitureMapping).length === 0) {
      alert('Не удалось найти блок фурнитуры. Проверьте лист и попробуйте снова.');
      return;
    }
    applyMappingToUI({
      furnitureSheet: selectedSheet,
      furnitureHeaderRow: furnitureMapping.furnitureHeaderRow,
      furnitureCodeCol: furnitureMapping.furnitureCodeCol,
      furnitureQtyCol: furnitureMapping.furnitureQtyCol,
      furnitureNameCol: furnitureMapping.furnitureNameCol,
      furnitureUnitCol: furnitureMapping.furnitureUnitCol,
      furniturePriceCol: furnitureMapping.furniturePriceCol,
    });
  }
  if (type === 'cost') {
    const costCell = detectBaseCostCell(sheet);
    if (!costCell) {
      alert('Не удалось найти строку с прямыми затратами. Укажите ячейку вручную.');
      return;
    }
    applyMappingToUI({ baseCostCell: costCell });
  }
  const indicator = document.getElementById('cursor-indicator');
  if (indicator) {
    indicator.textContent = 'Автоопределение применено к текущему блоку.';
  }
}

async function handleFileUpload(file) {
  if (!window.XLSX) {
    setUploadError('Не удалось загрузить библиотеку чтения Excel. Проверьте наличие vendor/xlsx.full.min.js.');
    return;
  }
  setUploadError('');
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellFormula: true, cellNF: true, cellDates: true });
  state.workbook = workbook;
  state.activeSheet = workbook.SheetNames[0];
  const anchors = findCalcSummaryAnchors(workbook);
  state.calcSummary = { anchors };
  renderSheetOptions();
  renderPreview(workbook.Sheets[state.activeSheet]);
  const mapping = autoDetectMapping(workbook.Sheets[state.activeSheet]);
  const costCell = detectBaseCostCell(workbook.Sheets[state.activeSheet]);
  applyMappingToUI({ ...mapping, baseCostCell: costCell });
  renderCalcSummaryAnchors(anchors);
  showScreen('mapping-screen');
}

function renderSheetOptions() {
  const sheetSelect = document.getElementById('sheet-select');
  const furnitureSelect = document.getElementById('furniture-sheet');
  sheetSelect.innerHTML = '';
  furnitureSelect.innerHTML = '<option value="">—</option>';
  state.workbook.SheetNames.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    sheetSelect.appendChild(option);
    const furnitureOption = option.cloneNode(true);
    furnitureSelect.appendChild(furnitureOption);
  });
  sheetSelect.value = state.activeSheet;
}

function attachEventHandlers() {
  if (!window.XLSX) {
    setUploadError('Не удалось загрузить библиотеку чтения Excel. Проверьте наличие vendor/xlsx.full.min.js.');
    document.getElementById('file-input').setAttribute('disabled', 'disabled');
  }

  document.getElementById('file-input').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) handleFileUpload(file);
  });

  const dropZone = document.getElementById('drop-zone');
  dropZone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dropZone.classList.add('dragover');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    dropZone.classList.remove('dragover');
    const file = event.dataTransfer.files[0];
    if (file) handleFileUpload(file);
  });

  document.getElementById('sheet-select').addEventListener('change', (event) => {
    state.activeSheet = event.target.value;
    renderPreview(state.workbook.Sheets[state.activeSheet]);
  });

  document.querySelectorAll('select[data-highlight]').forEach((select) => {
    select.addEventListener('change', (event) => {
      highlightColumn(letterToColIndex(event.target.value));
    });
  });

  document.getElementById('proceed-btn').addEventListener('click', () => {
    state.mapping = collectMapping();
    state.originalSpec = parseExcelWithMapping(state.workbook, state.mapping);
    showScreen('results-screen');
    renderBaseSummary(state.originalSpec);
    renderValidationSummary(state.originalSpec);
  });

  document.getElementById('calculate-btn').addEventListener('click', () => {
    const newWidth = Number(document.getElementById('new-width').value) || state.originalSpec.dims.width;
    const newDepth = Number(document.getElementById('new-depth').value) || state.originalSpec.dims.depth;
    const newHeight = Number(document.getElementById('new-height').value) || state.originalSpec.dims.height;
    const newSections = Number(document.getElementById('new-sections').value);
    const newShelves = Number(document.getElementById('new-shelves').value);
    const enforceRules = document.getElementById('enforce-rules').checked;
    const otherDriver = document.getElementById('other-driver').value;

    const roundedNew = {
      width: Math.round(newWidth),
      depth: Math.round(newDepth),
      height: Math.round(newHeight),
    };
    const roundedBase = {
      width: Math.round(state.originalSpec.dims.width || 0),
      depth: Math.round(state.originalSpec.dims.depth || 0),
      height: Math.round(state.originalSpec.dims.height || 0),
    };
    const overrides = {
      sectionCount: Number.isFinite(newSections) && newSections > 0 ? newSections : null,
      shelfCount: Number.isFinite(newShelves) && newShelves > 0 ? newShelves : null,
    };
    const noOverrides = overrides.sectionCount === null && overrides.shelfCount === null;
    const sameDims = roundedNew.width === roundedBase.width
      && roundedNew.depth === roundedBase.depth
      && roundedNew.height === roundedBase.height;

    if (sameDims && noOverrides) {
      const spec = state.originalSpec;
      const weight = calculateWeight(spec.corpus);
      const baseMaterialsCost = spec.baseMaterialCost || calculatePrice(spec.corpus, spec.materials || {});
      const baseHardwareCost = calculateFurnitureCost(spec.furniture || []);
      const baseOther = spec.baseCost
        ? Math.max(0, spec.baseCost - (baseMaterialsCost + baseHardwareCost))
        : 0;
      const price = spec.baseCost ?? calculatePrice(spec.corpus, spec.materials || {});
      state.newSpec = spec;
      renderResults(spec, weight, price, [], {
        materials: Math.round(baseMaterialsCost * 100) / 100,
        hardware: Math.round(baseHardwareCost * 100) / 100,
        other: Math.round(baseOther * 100) / 100,
        total: Math.round(price * 100) / 100,
      });
      const structure = getBaseStructure(spec);
      document.getElementById('structure-sections').textContent = formatNumber(structure.sections);
      document.getElementById('structure-partitions').textContent = formatNumber(structure.partitions);
      document.getElementById('structure-shelves').textContent = formatNumber(structure.shelves);
      return;
    }

    const worker = getWorker();
    worker.onmessage = (event) => {
      if (event.data.type === 'result') {
        const {
          spec,
          warnings,
          weight,
          price,
          structure,
          breakdown,
        } = event.data.payload;
        state.newSpec = spec;
        renderResults(spec, weight, price, warnings, breakdown);
        if (structure) {
          document.getElementById('structure-sections').textContent = formatNumber(structure.sections);
          document.getElementById('structure-partitions').textContent = formatNumber(structure.partitions);
          document.getElementById('structure-shelves').textContent = formatNumber(structure.shelves);
        }
      }
    };
    worker.postMessage({
      type: 'calculate',
      payload: {
        spec: state.originalSpec,
        newWidth,
        newDepth,
        newHeight,
        overrides,
        enforceRules,
        otherDriver,
      },
    });
  });

  document.getElementById('export-btn').addEventListener('click', () => {
    if (state.newSpec) exportToExcel(state.newSpec);
  });

  document.getElementById('reset-btn').addEventListener('click', () => {
    showScreen('upload-screen');
  });

  document.getElementById('back-btn').addEventListener('click', () => {
    showScreen('mapping-screen');
  });

  document.querySelectorAll('.auto-btn').forEach((btn) => {
    btn.addEventListener('click', () => updateMappingFromAuto(btn.dataset.auto));
  });

  document.querySelectorAll('.tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      const type = tab.dataset.tab;
      setActiveResultsTab(type);
      if (state.newSpec) {
        renderResultsTable(type, state.newSpec);
      }
    });
  });

  document.getElementById('save-template-btn').addEventListener('click', () => {
    const name = prompt('Название шаблона');
    if (!name) return;
    saveTemplate(name, collectMapping());
  });

  document.getElementById('export-template-btn').addEventListener('click', () => {
    const mapping = collectMapping();
    const payload = {
      version: 1,
      createdAt: new Date().toISOString(),
      mapping,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'wardrobe-mapping.json';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  });

  document.getElementById('import-template-btn').addEventListener('click', () => {
    document.getElementById('import-template-input').click();
  });

  document.getElementById('import-template-input').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        const mapping = data.mapping || data;
        applyMappingToUI(mapping);
        const name = data.name || null;
        if (name) {
          saveTemplate(name, mapping);
        }
      } catch (error) {
        alert('Не удалось прочитать файл шаблона. Проверьте формат JSON.');
      } finally {
        event.target.value = '';
      }
    };
    reader.readAsText(file);
  });

  document.getElementById('template-select').addEventListener('change', (event) => {
    const mapping = loadTemplate(event.target.value);
    if (mapping) applyMappingToUI(mapping);
  });

  document.querySelectorAll('input[data-cell-input]').forEach((input) => {
    const setActive = () => {
      state.activeCellInput = input;
    };
    input.addEventListener('focus', setActive);
    input.addEventListener('click', setActive);
  });
}

populateColumnSelects();
renderTemplateOptions();
attachEventHandlers();

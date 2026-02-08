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
  calcDebug: {
    visible: false,
    selectedRow: null,
  },
};

const COLUMN_LETTERS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
const MATERIAL_DENSITY = 720;
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

function setMappingWarning(message) {
  const box = document.getElementById('mapping-warning');
  if (!box) {
    if (message) alert(message);
    return;
  }
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
  let normalized = String(value)
    .replace(/[\s\u00a0]+/g, '')
    .replace(/[^\d,.\-]/g, '');
  if (!normalized) return null;
  normalized = normalized.replace(/(?!^)-/g, '');
  const hasComma = normalized.includes(',');
  const hasDot = normalized.includes('.');
  if (hasComma && hasDot) {
    const lastComma = normalized.lastIndexOf(',');
    const lastDot = normalized.lastIndexOf('.');
    if (lastComma > lastDot) {
      normalized = normalized.replace(/\./g, '').replace(',', '.');
    } else {
      normalized = normalized.replace(/,/g, '');
    }
  } else if (hasComma) {
    normalized = normalized.replace(',', '.');
  }
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
}

function round2(value) {
  return Math.round(value * 100) / 100;
}

function normalizeLabelText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[,.():;–—-]/g, ' ')
    .replace(/[^0-9a-zа-я\s]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isLaborHoursLabel(normalizedText) {
  const tokens = normalizedText.split(' ').filter(Boolean);
  const hasLabor = tokens.some((token) => token.startsWith('трудоемк'));
  const hasHours = tokens.some((token) => token.startsWith('час')) || tokens.some((token) => token.startsWith('человеко'));
  return hasLabor && hasHours;
}

function parseRef(rawRef, defaultSheet) {
  if (!rawRef) return null;
  const cleaned = String(rawRef).trim();
  const match = cleaned.match(/^(?:'([^']+)'|([^'!]+))?!?\s*(\$?[A-Za-z]{1,3}\$?\d+)$/);
  if (!match) return null;
  const sheetName = match[1] || match[2] || defaultSheet || null;
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

function isBlankCell(cell) {
  if (!cell) return true;
  if (cell.v === undefined || cell.v === null) return true;
  if (typeof cell.v === 'string' && !cell.v.trim()) return true;
  return false;
}

function detectMaterialTable(sheet) {
  if (!sheet) return null;
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const maxRows = Math.min(json.length, 200);
  let best = null;
  let bestScore = 0;

  const pickBest = (rowIndex, mapping, score) => {
    if (score > bestScore) {
      bestScore = score;
      best = { rowIndex, mapping };
    }
  };

  for (let row = 0; row < maxRows; row += 1) {
    const rowData = json[row] || [];
    const normalized = rowData.map((cell) => normalizeLabelText(cell));
    let nameCol = null;
    let thicknessCol = null;
    let lengthCol = null;
    let widthCol = null;
    let qtyCol = null;

    normalized.forEach((text, idx) => {
      if (!text) return;
      if (text.includes('наимен')) nameCol = idx;
      if (text.includes('толщ') || text.includes('тлщн')) thicknessCol = idx;
      if (text.includes('длин')) lengthCol = idx;
      if (text.includes('ширин')) widthCol = idx;
      if (text.includes('кол')) qtyCol = idx;
    });

    let score = 0;
    if (nameCol !== null) score += 2;
    if (lengthCol !== null) score += 1;
    if (widthCol !== null) score += 1;
    if (qtyCol !== null) score += 1;
    if (thicknessCol !== null) score += 0.5;

    if (score >= 4) {
      pickBest(row + 1, {
        nameCol,
        thicknessCol,
        lengthCol,
        widthCol,
        qtyCol,
      }, score);
    }
  }

  let headerRow = best?.rowIndex || null;
  let colMap = best?.mapping || null;

  if (!headerRow) {
    const fallbackRow = 50;
    headerRow = fallbackRow;
    colMap = {
      nameCol: 0,
      thicknessCol: 1,
      lengthCol: 2,
      widthCol: 3,
      qtyCol: 8,
    };
  }

  let totalsRow = null;
  const scanStart = Math.max(headerRow, 1);
  for (let row = scanStart; row < Math.min(scanStart + 60, json.length); row += 1) {
    const rowData = json[row] || [];
    const rowText = rowData.map((cell) => normalizeLabelText(cell)).join(' ');
    if (rowText.includes('итого')) {
      totalsRow = row + 1;
      break;
    }
  }

  let detailsStartRow = headerRow + 1;
  let detailsEndRow = totalsRow ? totalsRow - 1 : detailsStartRow + 20;

  if (!totalsRow && detailsStartRow === 51) {
    totalsRow = 66;
    detailsEndRow = 65;
  }

  const confidence = bestScore >= 4;

  return {
    detailsStartRow,
    detailsEndRow,
    totalsRow,
    colMap,
    headerRow,
    confidence,
  };
}

function findCalcSummaryAnchors(workbook) {
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
    прямыезатраты: 'totalCostRef',
  };
  const anchors = {};
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    Object.keys(sheet).forEach((cellRef) => {
      if (cellRef.startsWith('!')) return;
      const cell = sheet[cellRef];
      if (!cell || typeof cell.v !== 'string') return;
      const normalized = normalizeLabelText(cell.v);
      const compact = normalized.replace(/\s+/g, '');
      let anchorKey = labels[compact];
      if (!anchorKey && !anchors.laborHoursRef && isLaborHoursLabel(normalized)) {
        anchorKey = 'laborHoursRef';
      }
      if (!anchorKey || anchors[anchorKey]) return;
      const decoded = XLSX.utils.decode_cell(cellRef);
      let foundValueCell = null;
      for (let offset = 1; offset <= 5; offset += 1) {
        const candidateRef = XLSX.utils.encode_cell({ r: decoded.r, c: decoded.c + offset });
        const candidate = sheet[candidateRef];
        if (isNumericCell(candidate)) {
          foundValueCell = candidateRef;
          break;
        }
      }
      if (!foundValueCell) {
        for (let offset = 1; offset <= 3; offset += 1) {
          const candidateRef = XLSX.utils.encode_cell({ r: decoded.r + offset, c: decoded.c });
          const candidate = sheet[candidateRef];
          if (isNumericCell(candidate)) {
            foundValueCell = candidateRef;
            break;
          }
        }
      }
      if (!foundValueCell && anchorKey === 'laborHoursRef') {
        for (let offset = 1; offset <= 5; offset += 1) {
          const candidateRef = XLSX.utils.encode_cell({ r: decoded.r, c: decoded.c + offset });
          const candidate = sheet[candidateRef];
          if (isBlankCell(candidate)) {
            foundValueCell = candidateRef;
            break;
          }
        }
        if (!foundValueCell) {
          for (let offset = 1; offset <= 3; offset += 1) {
            const candidateRef = XLSX.utils.encode_cell({ r: decoded.r + offset, c: decoded.c });
            const candidate = sheet[candidateRef];
            if (isBlankCell(candidate)) {
              foundValueCell = candidateRef;
              break;
            }
          }
        }
      }
      if (foundValueCell) {
        anchors[anchorKey] = `${sheetName}!${foundValueCell}`;
      }
    });
  });
  return anchors;
}

function getNonZeroCostColumns(workbook, dspTotalRef, totalsRow) {
  const resolved = resolveAnchorFormula(workbook, dspTotalRef);
  if (!resolved) {
    return {
      costCols: [],
      allRefs: [],
      refsWithValues: [],
      warning: 'Не удалось получить формулу итога ДСП.',
    };
  }
  const refs = extractRefsFromFormula(resolved.formula);
  const expandedRefs = [];
  refs.forEach((rawRef) => {
    const normalized = normalizeFormulaRef(rawRef, resolved.sheetName);
    if (!normalized.sheetName || !normalized.ref) return;
    if (normalized.ref.includes(':')) {
      expandedRefs.push(...expandRangeRefs(normalized.ref, normalized.sheetName));
    } else {
      expandedRefs.push(`${normalized.sheetName}!${normalized.ref}`);
    }
  });

  const refsWithValues = [];
  const cols = new Set();
  const allCols = new Set();
  expandedRefs.forEach((fullRef) => {
    const parsed = parseRef(fullRef);
    if (!parsed) return;
    const sheet = workbook.Sheets[parsed.sheetName];
    if (!sheet) return;
    const decoded = XLSX.utils.decode_cell(parsed.cellRef);
    if (totalsRow && decoded.r + 1 !== totalsRow) return;
    const cell = sheet[parsed.cellRef];
    const value = parseNumericValue(cell?.v);
    const colIndex = decoded.c;
    allCols.add(colIndex);
    refsWithValues.push({
      ref: `${colIndexToLetter(colIndex)}${decoded.r + 1}`,
      value: Number.isFinite(value) ? value : 0,
      colIndex,
    });
    if (Number.isFinite(value) && Math.abs(value) > 0.0001) {
      cols.add(colIndex);
    }
  });

  const costCols = [...cols];
  const warning = costCols.length === 0
    ? 'Все итоги по ДСП нулевые — использую все колонки из формулы.'
    : null;

  return {
    costCols: costCols.length ? costCols : [...allCols],
    allRefs: expandedRefs,
    refsWithValues,
    warning,
  };
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

function traceCellLeaves(workbook, rootRef, opts = {}) {
  const leaves = [];
  const tree = [];
  const maxDepth = opts.maxDepth || 20;
  const visited = opts.visited || new Set();
  const stopPredicate = opts.stopPredicate;
  if (typeof stopPredicate !== 'function') {
    throw new Error('traceCellLeaves requires stopPredicate option.');
  }

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
    if (stopPredicate(ref, cell, {
      sheet,
      parsed,
      depth,
      path,
    })) {
      const value = parseNumericValue(cell?.v);
      const decoded = XLSX.utils.decode_cell(parsed.cellRef);
      leaves.push({
        ref,
        value: Number.isFinite(value) ? value : 0,
        formula: cell?.f || null,
        sheet: parsed.sheetName,
        row: decoded.r + 1,
        col: decoded.c + 1,
      });
      tree.push(node);
      return;
    }

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
        const decoded = XLSX.utils.decode_cell(parsed.cellRef);
        leaves.push({
          ref,
          value: Number.isFinite(value) ? value : 0,
          formula: cell.f || null,
          sheet: parsed.sheetName,
          row: decoded.r + 1,
          col: decoded.c + 1,
        });
      }
    } else {
      const value = parseNumericValue(cell?.v);
      const decoded = XLSX.utils.decode_cell(parsed.cellRef);
      leaves.push({
        ref,
        value: Number.isFinite(value) ? value : 0,
        formula: cell?.f || null,
        sheet: parsed.sheetName,
        row: decoded.r + 1,
        col: decoded.c + 1,
      });
    }
    tree.push(node);
  };

  trace(rootRef, 0, []);
  return { leaves, tree };
}

function normalizeFormulaRef(rawRef, defaultSheet) {
  const match = rawRef.match(/^(?:'([^']+)'|([^'!]+))!([A-Za-z0-9$]+(?::[A-Za-z0-9$]+)?)$/);
  if (match) {
    return {
      sheetName: match[1] || match[2] || defaultSheet || null,
      ref: match[3].replace(/\$/g, '').toUpperCase(),
    };
  }
  return {
    sheetName: defaultSheet || null,
    ref: rawRef.replace(/\$/g, '').toUpperCase(),
  };
}

function resolveAnchorFormula(workbook, anchorRef) {
  let currentRef = anchorRef;
  for (let depth = 0; depth < 10; depth += 1) {
    const parsed = readRefCell(workbook, currentRef);
    if (!parsed || !parsed.cell) return null;
    const { cell, sheetName } = parsed;
    if (!cell.f) return null;
    const cleaned = String(cell.f).replace(/\s+/g, '').replace(/^=/, '').replace(/^\+/, '');
    const singleRef = parseRef(cleaned, sheetName);
    if (singleRef) {
      currentRef = `${singleRef.sheetName}!${singleRef.cellRef}`;
      continue;
    }
    return { formula: cell.f, sheetName };
  }
  return null;
}

function extractCostColumnsFromAnchor(workbook, anchorRef) {
  const resolved = resolveAnchorFormula(workbook, anchorRef);
  if (!resolved) return [];
  const refs = extractRefsFromFormula(resolved.formula);
  const columns = new Set();
  refs.forEach((rawRef) => {
    const normalized = normalizeFormulaRef(rawRef, resolved.sheetName);
    if (!normalized.sheetName) return;
    if (normalized.ref.includes(':')) {
      const expanded = expandRangeRefs(normalized.ref, normalized.sheetName);
      expanded.forEach((expandedRef) => {
        const parsed = parseRef(expandedRef);
        if (!parsed) return;
        const col = XLSX.utils.decode_cell(parsed.cellRef).c;
        columns.add(colIndexToLetter(col));
      });
    } else {
      const parsed = parseRef(`${normalized.sheetName}!${normalized.ref}`);
      if (!parsed) return;
      const col = XLSX.utils.decode_cell(parsed.cellRef).c;
      columns.add(colIndexToLetter(col));
    }
  });
  return [...columns].filter(Boolean);
}

function readSheetCellValue(sheet, rowIndex, colIndex) {
  if (!sheet || !Number.isInteger(rowIndex) || !Number.isInteger(colIndex) || colIndex < 0) return undefined;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colIndex });
  return sheet[cellRef];
}

function inferAreaColumn(sheet, costColIndex, detailsStartRow, totalsRow) {
  if (!sheet || costColIndex === null || costColIndex === undefined) return null;
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const maxCol = range.e.c;
  const checkColumn = (colIndex) => {
    if (colIndex < 0 || colIndex > maxCol) return false;
    const totalsCell = totalsRow ? readSheetCellValue(sheet, totalsRow, colIndex) : null;
    const totalsValue = parseNumericValue(totalsCell?.v);
    if (!Number.isFinite(totalsValue) || totalsValue <= 0) return false;
    let hits = 0;
    let fractionalHits = 0;
    for (let row = detailsStartRow; row < Math.min(detailsStartRow + 10, totalsRow || detailsStartRow + 10); row += 1) {
      const cell = readSheetCellValue(sheet, row, colIndex);
      const value = parseNumericValue(cell?.v);
      if (Number.isFinite(value) && value > 0) {
        hits += 1;
        if (Math.abs(value % 1) > 0) fractionalHits += 1;
      }
    }
    return hits >= 2 && fractionalHits >= 1;
  };

  const defaultIndex = costColIndex - 1;
  if (checkColumn(defaultIndex)) return defaultIndex;

  for (let offset = 2; offset <= 6; offset += 1) {
    const candidate = costColIndex - offset;
    if (checkColumn(candidate)) return candidate;
  }

  return null;
}

function buildDspBreakdownFromTotals(sheet, tableInfo, costCols) {
  if (!sheet || !tableInfo || !costCols || costCols.length === 0) return null;
  const {
    detailsStartRow,
    detailsEndRow,
    totalsRow,
    colMap,
  } = tableInfo;
  if (!detailsStartRow || !detailsEndRow || !colMap) return null;

  const areaColMap = new Map();
  costCols.forEach((costColIndex) => {
    const areaColIndex = inferAreaColumn(sheet, costColIndex, detailsStartRow, totalsRow);
    if (areaColIndex !== null && areaColIndex !== undefined) {
      areaColMap.set(costColIndex, areaColIndex);
    }
  });

  const details = [];
  for (let rowIndex = detailsStartRow; rowIndex <= detailsEndRow; rowIndex += 1) {
    const nameRaw = readSheetCell(sheet, rowIndex, colMap.nameCol);
    const qty = parseNumericValue(readSheetCell(sheet, rowIndex, colMap.qtyCol));
    const length = parseNumericValue(readSheetCell(sheet, rowIndex, colMap.lengthCol));
    const width = parseNumericValue(readSheetCell(sheet, rowIndex, colMap.widthCol));
    const thickness = parseNumericValue(readSheetCell(sheet, rowIndex, colMap.thicknessCol));
    const sources = [];
    let rowCost = 0;
    let rowArea = 0;
    let hasAreaCell = false;

    costCols.forEach((costColIndex) => {
      const cell = readSheetCellValue(sheet, rowIndex, costColIndex);
      const value = parseNumericValue(cell?.v);
      const costRef = `${colIndexToLetter(costColIndex)}${rowIndex}`;
      if (Number.isFinite(value)) {
        rowCost += value;
      }
      const areaColIndex = areaColMap.get(costColIndex);
      let areaValue = null;
      let areaRef = null;
      if (areaColIndex !== undefined && areaColIndex !== null) {
        const areaCell = readSheetCellValue(sheet, rowIndex, areaColIndex);
        areaValue = parseNumericValue(areaCell?.v);
        areaRef = `${colIndexToLetter(areaColIndex)}${rowIndex}`;
        if (Number.isFinite(areaValue)) {
          rowArea += areaValue;
          hasAreaCell = true;
        }
      }
      sources.push({
        ref: costRef,
        value: Number.isFinite(value) ? value : 0,
        formula: cell?.f || null,
        areaRef,
        areaValue: Number.isFinite(areaValue) ? areaValue : null,
      });
    });

    if (!hasAreaCell) {
      const derived = length && width && qty ? (length * width / 1e6) * qty : null;
      if (Number.isFinite(derived)) rowArea = derived;
    }

    const name = nameRaw ? String(nameRaw).trim() : '';
    if (!name && rowCost === 0 && (!rowArea || rowArea === 0)) {
      continue;
    }

    details.push({
      name: name || `Строка ${rowIndex}`,
      qty: Number.isFinite(qty) ? qty : null,
      area_m2: Number.isFinite(rowArea) ? rowArea : null,
      thickness,
      length_mm: Number.isFinite(length) ? length : null,
      width_mm: Number.isFinite(width) ? width : null,
      cost: round2(rowCost),
      rowIndex,
      sources,
    });
  }

  return {
    details,
    areaColMap,
  };
}

function buildDspBreakdownFromColumns(workbook, mapping, detailsSheetName, costCols) {
  if (!mapping?.detailsStartRow || !mapping?.detailsEndRow) return [];
  if (!costCols || costCols.length === 0) return [];
  const sheetName = detailsSheetName || mapping.detailsSheet;
  if (!sheetName) return [];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  const costColIndexes = costCols.map((letter) => letterToColIndex(letter)).filter((idx) => idx >= 0);
  if (!costColIndexes.length) return [];

  const details = [];
  for (let rowIndex = mapping.detailsStartRow; rowIndex <= mapping.detailsEndRow; rowIndex += 1) {
    const nameRaw = readSheetCell(sheet, rowIndex, mapping.detailsNameCol);
    const qty = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsQtyCol));
    const length = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsLengthCol));
    const width = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsWidthCol));
    const thickness = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsThicknessCol));
    const sources = [];
    let rowCost = 0;
    costColIndexes.forEach((colIndex) => {
      const cell = readSheetCellValue(sheet, rowIndex, colIndex);
      if (!cell) return;
      const value = parseNumericValue(cell.v);
      if (!Number.isFinite(value)) return;
      rowCost += value;
      sources.push({
        ref: `${colIndexToLetter(colIndex)}${rowIndex}`,
        value,
        formula: cell.f || null,
      });
    });
    const area = length && width && qty ? (length * width / 1e6) * qty : null;
    const name = nameRaw ? String(nameRaw).trim() : '';
    if (!name && rowCost === 0 && (!area || area === 0)) {
      continue;
    }
    details.push({
      name: name || `Строка ${rowIndex}`,
      qty: Number.isFinite(qty) ? qty : null,
      area_m2: area,
      thickness,
      cost: round2(rowCost),
      rowIndex,
      sources,
    });
  }
  return details;
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
  headerData.forEach((cell, idx) => {
    const text = normalizeLabelText(cell);
    if (text.includes('наимен')) mapping.name = idx;
    if (text.includes('кол')) mapping.qty = idx;
    if (text.includes('площад')) mapping.area = idx;
    if (text.includes('стоим')) mapping.cost = idx;
    if (text.includes('длин')) mapping.length = idx;
    if (text.includes('ширин')) mapping.width = idx;
    if (text.includes('толщ')) mapping.thickness = idx;
  });
  const rowData = sheetMatrix[rowIndex - 1] || [];
  return {
    rowIndex,
    name: rowData[mapping.name],
    qty: parseNumericValue(rowData[mapping.qty]),
    area_m2: parseNumericValue(rowData[mapping.area]),
    thickness: parseNumericValue(rowData[mapping.thickness]),
    length_mm: parseNumericValue(rowData[mapping.length]),
    width_mm: parseNumericValue(rowData[mapping.width]),
  };
}

function readSheetCell(sheet, rowIndex, colIndex) {
  if (!sheet || !Number.isInteger(rowIndex) || !Number.isInteger(colIndex) || colIndex < 0) return undefined;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colIndex });
  return sheet[cellRef]?.v;
}

function buildDetailBreakdown(workbook, leaves, mapping, detailsSheetName) {
  if (!mapping?.detailsStartRow || !mapping?.detailsEndRow) return [];
  const sheetName = detailsSheetName || mapping.detailsSheet;
  if (!sheetName) return [];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  const details = [];
  const grouped = new Map();
  leaves.forEach((leaf) => {
    if (leaf.sheet !== sheetName) return;
    if (leaf.row < mapping.detailsStartRow || leaf.row > mapping.detailsEndRow) return;
    const key = `${leaf.sheet}!${leaf.row}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(leaf);
  });

  grouped.forEach((leafItems, key) => {
    const [sheetName, rowStr] = key.split('!');
    const rowIndex = Number(rowStr);
    const nameRaw = readSheetCell(sheet, rowIndex, mapping.detailsNameCol);
    const qty = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsQtyCol));
    const length = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsLengthCol));
    const width = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsWidthCol));
    const thickness = parseNumericValue(readSheetCell(sheet, rowIndex, mapping.detailsThicknessCol));
    const cost = leafItems.reduce((sum, item) => sum + (item.value || 0), 0);
    const area = length && width && qty ? (length * width / 1e6) * qty : null;
    const name = nameRaw ? String(nameRaw).trim() : '';
    if (!name && cost === 0 && (!area || area === 0)) {
      return;
    }
    details.push({
      name: name || `Строка ${rowIndex}`,
      qty: Number.isFinite(qty) ? qty : null,
      area_m2: area,
      thickness,
      cost: round2(cost),
      rowIndex,
      sources: leafItems.map((leaf) => ({
        ref: parseRef(leaf.ref, sheetName)?.cellRef || leaf.ref,
        value: leaf.value,
        formula: leaf.formula || null,
      })),
    });
  });
  return details;
}

function computeCoverage(sumBreakdown, totalValue) {
  if (totalValue === null || totalValue === undefined || totalValue === 0) return null;
  return sumBreakdown / totalValue;
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

function buildCalcSummary(workbook, anchors, mapping, detailsSheetName) {
  const summary = {
    anchors,
    baseValues: {},
    breakdown: {},
    rates: {},
  };
  if (!anchors) return summary;

  const anchorMap = [
    ['weightRef', 'weight'],
    ['dspRef', 'dsp'],
    ['edgeRef', 'edge'],
    ['plasticRef', 'plastic'],
    ['fabricRef', 'fabric'],
    ['hwImpRef', 'hwImp'],
    ['hwRepRef', 'hwRep'],
    ['packRef', 'pack'],
    ['laborRef', 'labor'],
    ['totalCostRef', 'totalCost'],
    ['laborHoursRef', 'laborHours'],
  ];

  anchorMap.forEach(([refKey, valueKey]) => {
    const value = readAnchorValue(workbook, anchors[refKey]);
    if (Number.isFinite(value)) summary.baseValues[valueKey] = value;
  });

  if (anchors.dspRef) {
    const anchorCell = readRefCell(workbook, anchors.dspRef);
    const sheetName = anchorCell?.sheetName || detailsSheetName || mapping.detailsSheet;
    const sheet = sheetName ? workbook.Sheets[sheetName] : null;
    const detectedTable = sheet ? detectMaterialTable(sheet) : null;
    let detectionWarning = null;
    if (!detectedTable?.confidence) {
      detectionWarning = 'Автодетект таблицы ДСП неуверен — проверьте маппинг деталей.';
    }
    const tableInfo = detectedTable?.confidence
      ? detectedTable
      : {
          detailsStartRow: mapping.detailsStartRow || detectedTable?.detailsStartRow,
          detailsEndRow: mapping.detailsEndRow || detectedTable?.detailsEndRow,
          totalsRow: detectedTable?.totalsRow,
          colMap: {
            nameCol: mapping.detailsNameCol ?? detectedTable?.colMap?.nameCol,
            thicknessCol: mapping.detailsThicknessCol ?? detectedTable?.colMap?.thicknessCol,
            lengthCol: mapping.detailsLengthCol ?? detectedTable?.colMap?.lengthCol,
            widthCol: mapping.detailsWidthCol ?? detectedTable?.colMap?.widthCol,
            qtyCol: mapping.detailsQtyCol ?? detectedTable?.colMap?.qtyCol,
          },
        };

    const costInfo = getNonZeroCostColumns(workbook, anchors.dspRef, tableInfo?.totalsRow);
    const costColsList = costInfo.costCols || [];
    const costColsLetters = costColsList.map((index) => colIndexToLetter(index)).filter(Boolean);
    const breakdownResult = sheet && tableInfo
      ? buildDspBreakdownFromTotals(sheet, tableInfo, costColsList)
      : null;

    const details = breakdownResult?.details || [];
    const breakdownSum = details.reduce((sum, item) => sum + (item.cost || 0), 0);
    let totalValue = readAnchorValue(workbook, anchors.dspRef);
    let coverageWarning = costInfo.warning || null;
    let coverage = null;
    if (totalValue === null || totalValue === undefined) {
      totalValue = breakdownSum;
      coverage = 1;
      coverageWarning = coverageWarning || 'У формулы нет cached value, использую сумму по деталям.';
    } else {
      coverage = computeCoverage(breakdownSum, totalValue);
    }
    if (detectionWarning) {
      coverageWarning = coverageWarning ? `${coverageWarning} ${detectionWarning}` : detectionWarning;
    }

    const coverageOutOfRange = coverage === null || coverage < 0.95 || coverage > 1.05;
    const shouldFallback = details.length === 0 || breakdownSum <= 0 || coverageOutOfRange;
    const fallbackDetails = shouldFallback
      ? buildDspBreakdownFromColumns(workbook, mapping, sheetName, costColsLetters)
      : [];
    const fallbackUsed = shouldFallback && fallbackDetails.length > 0;
    const finalDetails = fallbackUsed ? fallbackDetails : details;
    const baseParts = finalDetails.map((item) => ({
      name: item.name,
      qty: item.qty,
      length_mm: item.length_mm || null,
      width_mm: item.width_mm || null,
      thickness: item.thickness || null,
      area_m2: item.area_m2 || null,
      cost: item.cost || 0,
    }));

    summary.breakdown.dsp = {
      leafCount: details.length,
      leafSum: round2(breakdownSum),
      totalValue,
      coverage,
      details: finalDetails,
      usable: !coverageOutOfRange && breakdownSum > 0 && finalDetails.length > 0,
      costColsIndices: costColsList,
      costColsLetters,
      leafRefs: costInfo.refsWithValues?.map((ref) => ref.ref) || [],
      warning: coverageWarning,
      fallbackUsed,
      debug: {
        dspTotalRef: anchors.dspRef,
        dspTotalValue: totalValue,
        formulaRefs: costInfo.refsWithValues || [],
        costColsLetters,
        costColsIndices: costColsList,
        areaColsByCost: breakdownResult?.areaColMap
          ? [...breakdownResult.areaColMap.entries()].map(([costCol, areaCol]) => ({
              costCol,
              areaCol,
            }))
          : [],
      },
      baseParts,
    };
    summary.rates.dsp = buildDspRates(finalDetails);
  }

  if (anchors.edgeRef) {
    const edgeTrace = traceCellLeaves(workbook, anchors.edgeRef, {
      stopPredicate: () => false,
    });
    const leafSum = edgeTrace.leaves.reduce((sum, item) => sum + (item.value || 0), 0);
    const totalValue = summary.baseValues.edge;
    const coverage = computeCoverage(leafSum, totalValue);
    summary.breakdown.edge = {
      leafCount: edgeTrace.leaves.length,
      leafSum: round2(leafSum),
      totalValue,
      coverage,
      usable: coverage !== null ? coverage >= 0.99 && coverage <= 1.01 : false,
    };
  }

  if (anchors.weightRef) {
    const weightTrace = traceCellLeaves(workbook, anchors.weightRef, {
      stopPredicate: () => false,
    });
    const leafSum = weightTrace.leaves.reduce((sum, item) => sum + (item.value || 0), 0);
    const totalValue = summary.baseValues.weight;
    const coverage = computeCoverage(leafSum, totalValue);
    summary.breakdown.weight = {
      leafCount: weightTrace.leaves.length,
      leafSum: round2(leafSum),
      totalValue,
      coverage,
      usable: coverage !== null ? coverage >= 0.99 && coverage <= 1.01 : false,
    };
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

  if (!mapping.detailsStartRow || !mapping.detailsEndRow) {
    const tableInfo = detectMaterialTable(sheet);
    if (tableInfo) {
      mapping.detailsHeaderRow = tableInfo.headerRow;
      mapping.detailsStartRow = tableInfo.detailsStartRow;
      mapping.detailsEndRow = tableInfo.detailsEndRow;
      if (tableInfo.colMap) {
        mapping.detailsNameCol = tableInfo.colMap.nameCol;
        mapping.detailsThicknessCol = tableInfo.colMap.thicknessCol;
        mapping.detailsLengthCol = tableInfo.colMap.lengthCol;
        mapping.detailsWidthCol = tableInfo.colMap.widthCol;
        mapping.detailsQtyCol = tableInfo.colMap.qtyCol;
      }
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

function resolveAnchors(autoAnchors, overrides) {
  const resolved = { ...(autoAnchors || {}) };
  if (!overrides) return resolved;
  Object.keys(overrides).forEach((key) => {
    if (!resolved[key] && overrides[key]) {
      resolved[key] = overrides[key];
    }
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
  const calcSummary = buildCalcSummary(workbook, anchors, mapping, state.activeSheet);
  const baseParts = calcSummary?.breakdown?.dsp?.baseParts || [];
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
    baseParts,
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
  if (mat.includes('лдсп') || mat.includes('дсп')) return 720;
  if (mat.includes('мдф')) return 750;
  if (mat.includes('фанер')) return 650;
  if (mat.includes('двп') || mat.includes('оргалит')) return 850;
  if (mat.includes('стекл')) return 2500;
  return MATERIAL_DENSITY;
}

function calculateWeight(parts) {
  let totalKg = 0;
  parts.forEach((part) => {
    if (!part.thickness || !part.length_mm || !part.qty) return;
    const density = getMaterialDensity(part.material);
    const isShelf = inferPartType(part.name) === 'shelf';
    const lengths = isShelf && Array.isArray(part.widths_mm) && part.widths_mm.length
      ? part.widths_mm
      : [part.length_mm];
    const widths = [part.width_mm];
    for (let i = 0; i < part.qty; i += 1) {
      const lengthMm = lengths[Math.min(i, lengths.length - 1)] || part.length_mm || 0;
      const widthMm = widths[Math.min(i, widths.length - 1)] || part.width_mm || 0;
      const volumeM3 = (lengthMm / 1000) * (widthMm / 1000) * (part.thickness / 1000);
      totalKg += volumeM3 * density;
    }
  });
  return Math.round(totalKg * 100) / 100;
}

function calculateWeightFromBreakdown(parts) {
  if (!parts || parts.length === 0) return null;
  let total = 0;
  parts.forEach((part) => {
    if (!part.area_m2 || !part.thickness) return;
    total += 730 * part.area_m2 * (part.thickness / 1000);
  });
  return Math.round(total * 100) / 100;
}

function calculatePrice(parts, materials) {
  let total = 0;
  parts.forEach((part) => {
    if (!part.length_mm || !part.qty) return;
    const isShelf = inferPartType(part.name) === 'shelf';
    const lengths = isShelf && Array.isArray(part.widths_mm) && part.widths_mm.length
      ? part.widths_mm
      : [part.length_mm];
    const widths = [part.width_mm];
    let areaM2 = 0;
    for (let i = 0; i < part.qty; i += 1) {
      const lengthMm = lengths[Math.min(i, lengths.length - 1)] || part.length_mm || 0;
      const widthMm = widths[Math.min(i, widths.length - 1)] || part.width_mm || 0;
      areaM2 += (lengthMm / 1000) * (widthMm / 1000);
    }
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
  const baseParts = spec.baseParts || [];
  const hasBaseParts = baseParts.length > 0;
  const partsSource = hasBaseParts ? baseParts : (spec.corpus || []);
  const partitionsQty = partsSource.reduce((sum, part) => {
    const name = hasBaseParts ? part.name : part.name;
    return inferPartType(name) === 'partition' ? sum + (part.qty || 0) : sum;
  }, 0);
  if (partitionsQty > 0) {
    return Math.max(1, Math.round(partitionsQty) + 1);
  }
  return hasBaseParts ? 1 : (splitSections(spec.dims.width || 0).length || 1);
}

function getBaseStructure(spec) {
  const baseParts = spec.baseParts || [];
  if (baseParts.length > 0) {
    const shelves = baseParts.reduce((sum, part) => {
      return String(part.name || '').toLowerCase().includes('полк')
        ? sum + (part.qty || 0)
        : sum;
    }, 0);
    const partitions = baseParts.reduce((sum, part) => {
      const name = String(part.name || '').toLowerCase();
      if (name.includes('перегород') || name.includes('боков')) {
        return sum + (part.qty || 0);
      }
      return sum;
    }, 0);
    return {
      sections: partitions > 0 ? partitions + 1 : null,
      partitions: partitions > 0 ? partitions : null,
      shelves,
    };
  }
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
  const baseWeight = calculateWeightFromBreakdown(spec.baseParts)
    ?? calculateWeight(spec.corpus);
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
  const dspBreakdown = spec.calcSummary?.breakdown?.dsp;
  if (dspBreakdown?.fallbackUsed) {
    const percent = dspBreakdown.coverage !== null && dspBreakdown.coverage !== undefined
      ? `${round2(dspBreakdown.coverage * 100)}%`
      : 'нет данных';
    const warning = document.createElement('div');
    warning.textContent = `⚠️ Разбор ДСП покрывает ${percent} — используется fallback.`;
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
    const isShelf = inferPartType(part.name) === 'shelf'
      && Array.isArray(part.widths_mm)
      && part.widths_mm.length > 0;
    const tr = document.createElement('tr');
    [
      part.name,
      part.material,
      isShelf ? part.widths_mm.join(', ') : part.length_mm,
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
  const coverageWarning = document.getElementById('calc-coverage-warning');
  const diagnostics = document.getElementById('calc-dsp-diagnostics');
  const debugToggle = document.getElementById('calc-debug-toggle');
  const debugPanel = document.getElementById('calc-debug-panel');
  const debugMeta = document.getElementById('calc-debug-meta');
  if (!table) return;
  table.innerHTML = '';
  const breakdown = spec.calcSummary?.breakdown?.dsp;
  if (!breakdown) {
    state.calcDebug.selectedRow = null;
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
    if (coverageWarning) coverageWarning.textContent = '';
    if (diagnostics) diagnostics.innerHTML = '';
    if (debugMeta) debugMeta.innerHTML = '';
    if (debugPanel) debugPanel.classList.add('hidden');
    return;
  }
  if (!breakdown.details || breakdown.details.length === 0) {
    state.calcDebug.selectedRow = null;
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 4;
    cell.textContent = breakdown.usable === false
      ? 'Разбор ДСП недоступен: недостаточное покрытие или неполный маппинг.'
      : 'Разбор ДСП не найден.';
    row.appendChild(cell);
    table.appendChild(row);
    if (leafCount) leafCount.textContent = formatNumber(breakdown.leafCount);
    if (coverage) coverage.textContent = breakdown.coverage !== null && breakdown.coverage !== undefined
      ? `${round2(breakdown.coverage * 100)}%`
      : 'нет данных';
    if (leafSum) leafSum.textContent = formatNumber(breakdown.leafSum, '₽');
    if (totalValue) totalValue.textContent = formatNumber(breakdown.totalValue, '₽');
    if (coverageWarning) coverageWarning.textContent = '';
    if (diagnostics) {
      diagnostics.innerHTML = '';
      if (breakdown.usable === false) {
        const refs = (breakdown.leafRefs || []).slice(0, 10).join(', ') || '—';
        const cols = (breakdown.costColsIndices || []).join(', ') || '—';
        const letters = (breakdown.costColsLetters || []).join(', ') || '—';
        diagnostics.innerHTML = `
          <div><strong>Диагностика ДСП:</strong></div>
          <div>totalValue: ${formatNumber(breakdown.totalValue, '₽')}</div>
          <div>leafCount: ${formatNumber(breakdown.leafCount)}</div>
          <div>leafSum: ${formatNumber(breakdown.leafSum, '₽')}</div>
          <div>leafRefs: ${refs}</div>
          <div>costColsIndices: ${cols}</div>
          <div>costColsLetters: ${letters}</div>
        `;
      }
    }
    if (debugMeta) debugMeta.innerHTML = '';
    if (debugPanel) debugPanel.classList.add('hidden');
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

  breakdown.details.forEach((detail, index) => {
    const tr = document.createElement('tr');
    [detail.name, detail.qty ?? '', detail.area_m2 ? round2(detail.area_m2) : '', detail.cost ?? '']
      .forEach((value) => {
        const td = document.createElement('td');
        td.textContent = value;
        tr.appendChild(td);
      });
    if (state.calcDebug.selectedRow === index) {
      tr.classList.add('selected');
    }
    tr.addEventListener('click', () => {
      state.calcDebug.selectedRow = index;
      table.querySelectorAll('tr').forEach((row) => row.classList.remove('selected'));
      tr.classList.add('selected');
      if (state.calcDebug.visible) {
        renderCalcDebugSources(breakdown);
      }
    });
    table.appendChild(tr);
  });

  if (leafCount) leafCount.textContent = formatNumber(breakdown.leafCount);
  if (coverage) coverage.textContent = breakdown.coverage !== null && breakdown.coverage !== undefined
    ? `${round2(breakdown.coverage * 100)}%`
    : 'нет данных';
  if (leafSum) leafSum.textContent = formatNumber(breakdown.leafSum, '₽');
  if (totalValue) totalValue.textContent = formatNumber(breakdown.totalValue, '₽');
  if (coverageWarning) {
    const warnings = [];
    if (breakdown.warning) warnings.push(`⚠️ ${breakdown.warning}`);
    if (breakdown.coverage !== null && breakdown.coverage !== undefined
      && (breakdown.coverage > 1.05 || breakdown.coverage < 0.95)) {
      warnings.push('⚠️ Покрытие вне диапазона 95–105%. Включите «Показать источники» для проверки.');
    }
    coverageWarning.textContent = warnings.join(' ');
  }
  if (diagnostics) {
    diagnostics.innerHTML = '';
    if (breakdown.usable === false) {
      const refs = (breakdown.leafRefs || []).slice(0, 10).join(', ') || '—';
      const cols = (breakdown.costColsIndices || []).join(', ') || '—';
      const letters = (breakdown.costColsLetters || []).join(', ') || '—';
      diagnostics.innerHTML = `
        <div><strong>Диагностика ДСП:</strong></div>
        <div>totalValue: ${formatNumber(breakdown.totalValue, '₽')}</div>
        <div>leafCount: ${formatNumber(breakdown.leafCount)}</div>
        <div>leafSum: ${formatNumber(breakdown.leafSum, '₽')}</div>
        <div>leafRefs: ${refs}</div>
        <div>costColsIndices: ${cols}</div>
        <div>costColsLetters: ${letters}</div>
      `;
    }
  }
  if (debugMeta) {
    debugMeta.innerHTML = '';
    const debug = breakdown.debug || {};
    const refs = (debug.formulaRefs || []).map((item) => `${item.ref}: ${round2(item.value)}`).join(', ') || '—';
    const cols = (debug.costColsLetters || []).join(', ') || '—';
    const areaCols = (debug.areaColsByCost || [])
      .map((pair) => `${colIndexToLetter(pair.costCol)}→${colIndexToLetter(pair.areaCol)}`)
      .join(', ') || '—';
    debugMeta.innerHTML = `
      <div><strong>Debug ДСП:</strong></div>
      <div>dspTotalRef: ${debug.dspTotalRef || '—'}</div>
      <div>dspTotalValue: ${formatNumber(debug.dspTotalValue, '₽')}</div>
      <div>formula refs: ${refs}</div>
      <div>costCols: ${cols}</div>
      <div>areaCols: ${areaCols}</div>
    `;
  }
  if (debugToggle) {
    debugToggle.textContent = state.calcDebug.visible ? 'Скрыть источники' : 'Показать источники';
    debugToggle.onclick = () => {
      state.calcDebug.visible = !state.calcDebug.visible;
      debugToggle.textContent = state.calcDebug.visible ? 'Скрыть источники' : 'Показать источники';
      if (debugPanel) {
        debugPanel.classList.toggle('hidden', !state.calcDebug.visible);
      }
      if (state.calcDebug.visible) {
        renderCalcDebugSources(breakdown);
      }
    };
  }
  if (debugPanel) {
    debugPanel.classList.toggle('hidden', !state.calcDebug.visible);
  }
  if (state.calcDebug.visible) {
    renderCalcDebugSources(breakdown);
  }
}

function renderCalcDebugSources(breakdown) {
  const debugList = document.getElementById('calc-debug-list');
  const debugTitle = document.getElementById('calc-debug-title');
  if (!debugList || !debugTitle) return;
  debugList.innerHTML = '';
  const selectedIndex = state.calcDebug.selectedRow;
  if (selectedIndex === null || selectedIndex === undefined) {
    debugTitle.textContent = 'Источники: выберите строку в таблице';
    return;
  }
  const detail = breakdown.details[selectedIndex];
  if (!detail) {
    debugTitle.textContent = 'Источники: строка не найдена';
    return;
  }
  const qtyInfo = detail.qty ? ` (qty: ${detail.qty})` : '';
  debugTitle.textContent = `Источники для: ${detail.name}${qtyInfo}`;
  if (!detail.sources || detail.sources.length === 0) {
    const item = document.createElement('li');
    item.textContent = 'Нет данных об источниках.';
    debugList.appendChild(item);
    return;
  }
  detail.sources.forEach((source) => {
    const item = document.createElement('li');
    const formula = source.formula ? ` | formula: ${source.formula}` : '';
    const areaInfo = source.areaRef
      ? ` | area ${source.areaRef}: ${source.areaValue !== null ? round2(source.areaValue) : '—'}`
      : '';
    item.textContent = `${source.ref}: ${round2(source.value)}${areaInfo}${formula}`;
    debugList.appendChild(item);
  });
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
  const workbook = XLSX.read(data, { type: 'array' });
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
    try {
      setMappingWarning('');
      state.mapping = collectMapping();
      state.originalSpec = parseExcelWithMapping(state.workbook, state.mapping);
      showScreen('results-screen');
      renderBaseSummary(state.originalSpec);
      renderValidationSummary(state.originalSpec);
    } catch (error) {
      console.error(error);
      const message = error?.message ? `Ошибка при обработке: ${error.message}` : 'Ошибка при обработке файла.';
      setMappingWarning(message);
    }
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
      const weight = calculateWeightFromBreakdown(spec.baseParts)
        ?? calculateWeight(spec.corpus);
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

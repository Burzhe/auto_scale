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
};

const COLUMN_LETTERS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
const MATERIAL_DENSITY = 720;

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
  const normalized = String(value).replace(/\s+/g, '').replace(',', '.');
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
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
  };
}

function parseExcelWithMapping(workbook, mapping) {
  const sheet = workbook.Sheets[state.activeSheet];
  const materials = parseMaterialDictionary(sheet, mapping);
  const corpus = parseCorpusDetails(sheet, mapping, materials);
  const furnitureSheet = workbook.Sheets[mapping.furnitureSheet];
  const furniture = parseFurniture(furnitureSheet, mapping);
  const dims = parseDimensions(sheet, mapping);
  const baseCost = mapping.baseCostCell ? readCellNumber(sheet, mapping.baseCostCell) : null;
  const baseMaterialCost = calculatePrice(corpus, materials);

  return {
    dims,
    corpus,
    furniture,
    materials,
    baseCost: Number.isFinite(baseCost) ? baseCost : null,
    baseMaterialCost,
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
    const widths = part.widths_mm || [part.width_mm];
    for (let i = 0; i < part.qty; i += 1) {
      const w = widths[Math.min(i, widths.length - 1)] || part.width_mm || 0;
      const volumeM3 = (part.length_mm / 1000) * (w / 1000) * (part.thickness / 1000);
      totalKg += volumeM3 * density;
    }
  });
  return Math.round(totalKg * 100) / 100;
}

function calculatePrice(parts, materials) {
  let total = 0;
  parts.forEach((part) => {
    if (!part.length_mm || !part.qty) return;
    const widths = part.widths_mm || [part.width_mm];
    let areaM2 = 0;
    for (let i = 0; i < part.qty; i += 1) {
      const w = widths[Math.min(i, widths.length - 1)] || part.width_mm || 0;
      areaM2 += (part.length_mm / 1000) * (w / 1000);
    }
    const material = materials[part.material_id];
    if (material) {
      const wasteFactor = 1 + (material.waste || 0) / 100;
      total += areaM2 * material.price * wasteFactor;
    }
  });
  return Math.round(total);
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

function renderResultsTable(type, spec) {
  const table = document.getElementById('results-table');
  table.innerHTML = '';
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
      part.widths_mm ? part.widths_mm.join(', ') : part.width_mm,
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

function renderResults(spec, weight, price, warnings) {
  if (state.originalSpec) {
    renderBaseSummary(state.originalSpec);
  }

  document.getElementById('new-dims').textContent = formatDimensions(spec.dims);
  document.getElementById('new-weight').textContent = formatNumber(weight, 'кг');
  document.getElementById('new-price').textContent = formatNumber(price, '₽');

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
    setUploadError('Не удалось загрузить библиотеку чтения Excel. Проверьте доступ к интернету и перезагрузите страницу.');
    return;
  }
  setUploadError('');
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  state.workbook = workbook;
  state.activeSheet = workbook.SheetNames[0];
  renderSheetOptions();
  renderPreview(workbook.Sheets[state.activeSheet]);
  const mapping = autoDetectMapping(workbook.Sheets[state.activeSheet]);
  const costCell = detectBaseCostCell(workbook.Sheets[state.activeSheet]);
  applyMappingToUI({ ...mapping, baseCostCell: costCell });
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
    setUploadError('Не удалось загрузить библиотеку чтения Excel. Проверьте доступ к интернету и перезагрузите страницу.');
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
  });

  document.getElementById('calculate-btn').addEventListener('click', () => {
    const newWidth = Number(document.getElementById('new-width').value) || state.originalSpec.dims.width;
    const newDepth = Number(document.getElementById('new-depth').value) || state.originalSpec.dims.depth;
    const newHeight = Number(document.getElementById('new-height').value) || state.originalSpec.dims.height;
    const newSections = Number(document.getElementById('new-sections').value);
    const newShelves = Number(document.getElementById('new-shelves').value);

    const worker = getWorker();
    worker.onmessage = (event) => {
      if (event.data.type === 'result') {
        const { spec, warnings, weight, price, structure } = event.data.payload;
        state.newSpec = spec;
        renderResults(spec, weight, price, warnings);
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
        overrides: {
          sectionCount: Number.isFinite(newSections) && newSections > 0 ? newSections : null,
          shelfCount: Number.isFinite(newShelves) && newShelves > 0 ? newShelves : null,
        },
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

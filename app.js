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
  showCalcSources: false,
};

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
  const furnitureSheetEl = document.getElementById('furniture-sheet');
  if (furnitureSheetEl && mapping.furnitureSheet !== undefined) furnitureSheetEl.value = mapping.furnitureSheet;
  const furnitureHeaderEl = document.getElementById('furniture-header');
  if (furnitureHeaderEl && mapping.furnitureHeaderRow) furnitureHeaderEl.value = mapping.furnitureHeaderRow;
  const furnitureCodeEl = document.getElementById('furniture-code-col');
  if (furnitureCodeEl && Number.isInteger(mapping.furnitureCodeCol)) furnitureCodeEl.value = colIndexToLetter(mapping.furnitureCodeCol);
  const furnitureQtyEl = document.getElementById('furniture-qty-col');
  if (furnitureQtyEl && Number.isInteger(mapping.furnitureQtyCol)) furnitureQtyEl.value = colIndexToLetter(mapping.furnitureQtyCol);
  const furnitureNameEl = document.getElementById('furniture-name-col');
  if (furnitureNameEl && Number.isInteger(mapping.furnitureNameCol)) furnitureNameEl.value = colIndexToLetter(mapping.furnitureNameCol);
  const furnitureUnitEl = document.getElementById('furniture-unit-col');
  if (furnitureUnitEl && Number.isInteger(mapping.furnitureUnitCol)) furnitureUnitEl.value = colIndexToLetter(mapping.furnitureUnitCol);
  const furniturePriceEl = document.getElementById('furniture-price-col');
  if (furniturePriceEl && Number.isInteger(mapping.furniturePriceCol)) furniturePriceEl.value = colIndexToLetter(mapping.furniturePriceCol);
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
  const mapping = {
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

  const furnitureSheetEl = document.getElementById('furniture-sheet');
  if (furnitureSheetEl) {
    mapping.furnitureSheet = furnitureSheetEl.value;
  }
  const furnitureHeaderEl = document.getElementById('furniture-header');
  if (furnitureHeaderEl) {
    mapping.furnitureHeaderRow = Number(furnitureHeaderEl.value);
  }
  const furnitureCodeEl = document.getElementById('furniture-code-col');
  if (furnitureCodeEl) {
    mapping.furnitureCodeCol = letterToColIndex(furnitureCodeEl.value);
  }
  const furnitureQtyEl = document.getElementById('furniture-qty-col');
  if (furnitureQtyEl) {
    mapping.furnitureQtyCol = letterToColIndex(furnitureQtyEl.value);
  }
  const furnitureNameEl = document.getElementById('furniture-name-col');
  if (furnitureNameEl) {
    mapping.furnitureNameCol = letterToColIndex(furnitureNameEl.value);
  }
  const furnitureUnitEl = document.getElementById('furniture-unit-col');
  if (furnitureUnitEl) {
    mapping.furnitureUnitCol = letterToColIndex(furnitureUnitEl.value);
  }
  const furniturePriceEl = document.getElementById('furniture-price-col');
  if (furniturePriceEl) {
    mapping.furniturePriceCol = letterToColIndex(furniturePriceEl.value);
  }

  return mapping;
}

function formatNumber(value, unit = '') {
  if (value === undefined || value === null || Number.isNaN(value)) return '—';
  return `${Number(value).toLocaleString('ru-RU')} ${unit}`.trim();
}

function formatDimensions(dims) {
  if (!dims || !dims.width || !dims.depth || !dims.height) return '—';
  return `${dims.width}×${dims.depth}×${dims.height}`;
}

function renderBaseSummary(spec) {
  const baseWeight = calculateWeight(spec.corpus, inferDensityFromSpec(spec));
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
  const totalCost = Number.isFinite(baseValues.totalCost) ? baseValues.totalCost : baseCost;
  const baseHardwareCost = Number.isFinite(baseValues.hwImp) || Number.isFinite(baseValues.hwRep)
    ? (Number(baseValues.hwImp || 0) + Number(baseValues.hwRep || 0))
    : baseHardwareCostFallback;
  const baseOther = baseCost !== null && baseCost !== undefined
    ? baseCost - (baseMaterialsCost + baseHardwareCost)
    : null;
  document.getElementById('validation-dsp').textContent = formatNumber(baseValues.dsp, '₽');
  document.getElementById('validation-edge').textContent = formatNumber(baseValues.edge, '₽');
  document.getElementById('validation-plastic').textContent = formatNumber(baseValues.plastic, '₽');
  document.getElementById('validation-fabric').textContent = formatNumber(baseValues.fabric, '₽');
  document.getElementById('validation-hw-imp').textContent = formatNumber(baseValues.hwImp, '₽');
  document.getElementById('validation-hw-rep').textContent = formatNumber(baseValues.hwRep, '₽');
  document.getElementById('validation-pack').textContent = formatNumber(baseValues.pack, '₽');
  document.getElementById('validation-labor').textContent = formatNumber(baseValues.labor, '₽');
  document.getElementById('validation-total').textContent = formatNumber(totalCost, '₽');

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
  const reasonBox = document.getElementById('calc-breakdown-reason');
  const toggleSourcesBtn = document.getElementById('calc-toggle-sources');
  if (!table) return;
  table.innerHTML = '';
  const breakdown = spec.calcSummary?.breakdown?.dsp;
  if (reasonBox) reasonBox.textContent = '';
  if (!breakdown || !breakdown.details || breakdown.details.length === 0) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 5;
    cell.textContent = 'Разбор ДСП не найден.';
    row.appendChild(cell);
    table.appendChild(row);
    if (leafCount) leafCount.textContent = '—';
    if (coverage) coverage.textContent = '—';
    if (leafSum) leafSum.textContent = '—';
    if (totalValue) totalValue.textContent = '—';
    if (reasonBox && breakdown?.reason) {
      reasonBox.textContent = `Причина: ${breakdown.reason}`;
    }
    if (toggleSourcesBtn) toggleSourcesBtn.classList.add('hidden');
    return;
  }
  if (toggleSourcesBtn) {
    toggleSourcesBtn.classList.remove('hidden');
    toggleSourcesBtn.textContent = state.showCalcSources ? 'Скрыть источники' : 'Показать источники';
  }
  const headers = ['Деталь', 'Qty', 'Area (м²)', 'Cost (₽)', 'Источники'];
  const headerRow = document.createElement('tr');
  headers.forEach((text) => {
    const th = document.createElement('th');
    th.textContent = text;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  breakdown.details.forEach((detail) => {
    const tr = document.createElement('tr');
    const sourceText = detail.sources
      ? `total: ${detail.sources.totalCell || '—'}; term: ${detail.sources.colTotalCell || '—'}; leaf: ${detail.sources.leafCell || '—'}`
      : '';
    const values = [
      detail.name,
      detail.qty ?? '',
      detail.area_m2 ? round2(detail.area_m2) : '',
      detail.cost_rub ?? detail.cost ?? '',
      state.showCalcSources ? sourceText : '—',
    ];
    values.forEach((value) => {
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
  if (reasonBox && breakdown.reason) {
    reasonBox.textContent = `Причина: ${breakdown.reason}`;
  }
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
    const furnitureSheetEl = document.getElementById('furniture-sheet');
    if (!furnitureSheetEl) return;
    const selectedSheet = furnitureSheetEl.value || state.activeSheet;
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
  const workbook = XLSX.read(data, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
    cellText: true,
  });
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
  if (furnitureSelect) furnitureSelect.innerHTML = '<option value="">—</option>';
  state.workbook.SheetNames.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    sheetSelect.appendChild(option);
    if (furnitureSelect) {
      const furnitureOption = option.cloneNode(true);
      furnitureSelect.appendChild(furnitureOption);
    }
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
      const weight = calculateWeight(spec.corpus, inferDensityFromSpec(spec));
      const baseMaterialsCost = spec.baseMaterialCost || calculatePrice(spec.corpus, spec.materials || {});
      const baseValues = spec.calcSummary?.baseValues || {};
      const baseHardwareCost = Number.isFinite(baseValues.hwImp) || Number.isFinite(baseValues.hwRep)
        ? (Number(baseValues.hwImp || 0) + Number(baseValues.hwRep || 0))
        : calculateFurnitureCost(spec.furniture || []);
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

  const toggleSourcesBtn = document.getElementById('calc-toggle-sources');
  if (toggleSourcesBtn) {
    toggleSourcesBtn.addEventListener('click', () => {
      state.showCalcSources = !state.showCalcSources;
      if (state.newSpec) {
        renderCalcBreakdown(state.newSpec);
      }
    });
  }

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

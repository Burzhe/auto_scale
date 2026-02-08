import { state } from './state.js';
import {
  formatNumber,
  formatDimensions,
  calculateWeight,
  inferDensityFromSpec,
  getBasePriceFromSpec,
  calculatePrice,
  calculateFurnitureCost,
  round2,
} from './calculations.js';
import {
  colIndexToLetter,
  normalizeCellRef,
} from './excel.js';

export function setUploadError(message) {
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

export function showScreen(id) {
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

export function populateColumnSelects() {
  document.querySelectorAll('select[data-highlight]').forEach((select) => {
    select.innerHTML = Array.from({ length: 26 }, (_, i) => colIndexToLetter(i))
      .map((letter) => `<option value="${letter}">${letter}</option>`)
      .join('');
  });
}

export function renderPreview(sheet) {
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

export function highlightColumn(colIndex) {
  const table = document.getElementById('preview-table');
  if (!table) return;
  table.querySelectorAll('td, th').forEach((cell) => {
    cell.classList.remove('highlighted');
  });
  table.querySelectorAll(`td:nth-child(${colIndex + 2}), th:nth-child(${colIndex + 2})`).forEach((cell) => {
    cell.classList.add('highlighted');
  });
}

export function attachPreviewHover() {
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

export function attachCellSelection() {
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

export function applyMappingToUI(mapping) {
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

export function saveTemplate(name, mapping) {
  const templates = JSON.parse(localStorage.getItem('mapping-templates') || '{}');
  templates[name] = mapping;
  localStorage.setItem('mapping-templates', JSON.stringify(templates));
  renderTemplateOptions();
}

export function loadTemplate(name) {
  const templates = JSON.parse(localStorage.getItem('mapping-templates') || '{}');
  return templates[name];
}

export function renderTemplateOptions() {
  const select = document.getElementById('template-select');
  if (!select) return;
  const templates = JSON.parse(localStorage.getItem('mapping-templates') || '{}');
  select.innerHTML = '<option value="">Выбрать шаблон...</option>';
  Object.keys(templates).forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });
}

export function renderCalcSummaryAnchors(anchors) {
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

export function renderSheetOptions() {
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

export function renderBaseSummary(spec) {
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

export function renderValidationSummary(spec) {
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

export function renderResultsTable(type, spec) {
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

export function setActiveResultsTab(type) {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('active', tab.dataset.tab === type);
  });
  state.activeResultsTab = type;
}

export function renderCalcBreakdown(spec) {
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

export function renderMaterialSpecOptions(spec) {
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

export function renderResults(spec, weight, price, warnings, breakdown) {
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

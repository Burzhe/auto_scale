import { state } from './state.js';
import {
  calculateWeight,
  inferDensityFromSpec,
  calculatePrice,
  calculateFurnitureCost,
  getBaseStructure,
  formatNumber,
  round2,
} from './calculations.js';
import {
  autoDetectMapping,
  autoDetectFurnitureMapping,
  detectBaseCostCell,
  exportToExcel,
  findCalcSummaryAnchors,
  letterToColIndex,
  normalizeAnchorRef,
  normalizeCellRef,
  parseExcelWithMapping,
} from './excel.js';
import {
  applyMappingToUI,
  highlightColumn,
  renderBaseSummary,
  renderCalcBreakdown,
  renderCalcSummaryAnchors,
  renderPreview,
  renderResults,
  renderResultsTable,
  renderSheetOptions,
  renderValidationSummary,
  setActiveResultsTab,
  setUploadError,
  showScreen,
  saveTemplate,
  loadTemplate,
} from './ui.js';

const XLSX = globalThis.XLSX;

function getWorker() {
  if (state.worker) return state.worker;
  const workerSource = document.getElementById('worker-src').textContent;
  const blob = new Blob([workerSource], { type: 'text/javascript' });
  const url = URL.createObjectURL(blob);
  state.worker = new Worker(url);
  return state.worker;
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

export function attachEventHandlers() {
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
      const value = event.target.value;
      const index = letterToColIndex(value);
      if (Number.isInteger(index) && index >= 0) {
        highlightColumn(index);
      }
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
        materials: round2(baseMaterialsCost),
        hardware: round2(baseHardwareCost),
        other: round2(baseOther),
        total: round2(price),
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
    if (state.newSpec) {
      exportToExcel(state.newSpec);
    }
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

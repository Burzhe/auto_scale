export const state = {
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

export const COLUMN_LETTERS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
export const MATERIAL_DENSITY = 730;
export const MAX_SECTION_WIDTH = 1200;
export const PARTITION_THRESHOLD = 800;

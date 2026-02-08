function findSummaryAnchorsInSheet(sheet, sheetName, labelMap) {
  const anchors = {};
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const labelEntries = Object.entries(labelMap).sort((a, b) => b[0].length - a[0].length);

  for (let r = range.s.r; r <= range.e.r; r += 1) {
    const c = 0;
    if (c < range.s.c || c > range.e.c) continue;
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr];
    if (!cell || cell.v == null) continue;
    const normalized = normalizeLabelText(cell.v);
    if (!normalized) continue;

    for (const [labelKey, anchorKey] of labelEntries) {
      if (anchors[anchorKey]) continue;
      if (normalized === labelKey || normalized.startsWith(labelKey) || normalized.includes(labelKey)) {
        let valueRef = null;
        const rightAddr = XLSX.utils.encode_cell({ r, c: c + 1 });
        const rightCell = sheet[rightAddr];
        if (rightCell && (rightCell.v != null || rightCell.f)) {
          valueRef = `${sheetName}!${rightAddr}`;
        }
        if (valueRef) anchors[anchorKey] = valueRef;
        break;
      }
    }
  }

  return anchors;
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
    стоимостьрасчетасуммарно: 'totalCostRef',
  };

  const normalizedSheetName = (name) => normalizeLabelText(name);
  const preferred = workbook.SheetNames.find((name) => normalizedSheetName(name).includes('плитнматериал'));
  const orderedSheets = preferred
    ? [preferred, ...workbook.SheetNames.filter((name) => name !== preferred)]
    : workbook.SheetNames.slice();

  const anchors = {};
  orderedSheets.forEach((sheetName) => {
    const ws = workbook.Sheets[sheetName];
    if (!ws) return;
    const found = findSummaryAnchorsInSheet(ws, sheetName, labels);
    Object.entries(found).forEach(([key, value]) => {
      if (!anchors[key]) anchors[key] = value;
    });
  });

  return anchors;
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
      const refs = parseRefsFromFormula(cell.f);
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
    const area = Number(item.area_m2 || 0);
    const cost = Number(item.cost_rub ?? item.cost ?? 0);
    if (!area || !cost) return;
    const thicknessKey = item.thickness_mm ? String(Math.round(item.thickness_mm)) : (item.thickness ? String(item.thickness) : null);
    if (thicknessKey) {
      totals[thicknessKey] = totals[thicknessKey] || { area: 0, cost: 0 };
      totals[thicknessKey].area += area;
      totals[thicknessKey].cost += cost;
    }
    totalArea += area;
    totalCost += cost;
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

function aggregateItemsByName(items) {
  const grouped = new Map();
  items.forEach((item) => {
    const key = String(item.name || '').trim() || 'Без названия';
    const entry = grouped.get(key) || {
      name: key,
      qty: 0,
      area_m2: 0,
      cost_rub: 0,
    };
    entry.qty += Number(item.qty || 0);
    entry.area_m2 += Number(item.area_m2 || 0);
    entry.cost_rub += Number(item.cost_rub || 0);
    grouped.set(key, entry);
  });
  return Array.from(grouped.values());
}

function buildChipboardBreakdown(workbook, totalRef) {
  const parsed = parseRef(totalRef);
  if (!parsed || !parsed.sheetName || !parsed.cellRef) {
    return {
      totalValue: 0,
      leafRefs: [],
      leafSum: 0,
      coverage: null,
      items: [],
      aggregated: [],
      debug: { reason: 'bad total cell ref' },
    };
  }

  const sheet = workbook.Sheets[parsed.sheetName];
  if (!sheet) {
    return {
      totalValue: 0,
      leafRefs: [],
      leafSum: 0,
      coverage: null,
      items: [],
      aggregated: [],
      debug: { reason: 'sheet not found' },
    };
  }

  const totalCell = sheet[parsed.cellRef];
  const totalValue = parseNumericValue(totalCell?.v) || 0;
  const totalFormula = getCellFormulaString(totalCell);
  if (!totalFormula) {
    return {
      totalValue,
      leafRefs: [],
      leafSum: 0,
      coverage: null,
      items: [],
      aggregated: [],
      debug: { reason: 'total cell has no formula' },
    };
  }

  const termCells = extractA1Refs(totalFormula);
  if (!termCells.length) {
    return {
      totalValue,
      leafRefs: [],
      leafSum: 0,
      coverage: null,
      items: [],
      aggregated: [],
      debug: { reason: 'term cells not found', formula: totalFormula },
    };
  }

  const leafSet = new Set();
  const leafToTerm = new Map();
  termCells.forEach((termCellRef) => {
    const termCell = sheet[termCellRef];
    const termFormula = getCellFormulaString(termCell);
    const sumRefs = parseSumFormulaRefs(termFormula);
    if (sumRefs.length) {
      sumRefs.forEach((leafRef) => {
        leafSet.add(leafRef);
        if (!leafToTerm.has(leafRef)) leafToTerm.set(leafRef, termCellRef);
      });
    } else {
      leafSet.add(termCellRef);
      if (!leafToTerm.has(termCellRef)) leafToTerm.set(termCellRef, termCellRef);
    }
  });

  const leafRefs = Array.from(leafSet);
  const items = [];
  let leafSum = 0;

  leafRefs.forEach((leafRef) => {
    const leafCell = sheet[leafRef];
    const cost = parseNumericValue(leafCell?.v);
    if (!Number.isFinite(cost) || cost === 0) return;
    const addr = XLSX.utils.decode_cell(leafRef);
    const row = addr.r + 1;
    const name = sheet[`A${row}`]?.v;
    const thickness = parseNumericValue(sheet[`B${row}`]?.v);
    const length = parseNumericValue(sheet[`C${row}`]?.v);
    const width = parseNumericValue(sheet[`D${row}`]?.v);
    const qty = parseNumericValue(sheet[`I${row}`]?.v);

    let area = null;
    let areaCell = null;
    if (Number.isFinite(length) && Number.isFinite(width) && Number.isFinite(qty) && qty > 0) {
      area = (length * width * qty) / 1e6;
    } else {
      const areaAddr = XLSX.utils.encode_cell({ r: addr.r, c: Math.max(0, addr.c - 1) });
      const areaVal = parseNumericValue(sheet[areaAddr]?.v);
      if (Number.isFinite(areaVal) && areaVal > 0) {
        area = areaVal;
        areaCell = areaAddr;
      }
    }

    items.push({
      name: String(name || '').trim() || `Строка ${row}`,
      length_mm: Number.isFinite(length) ? length : null,
      width_mm: Number.isFinite(width) ? width : null,
      thickness_mm: Number.isFinite(thickness) ? thickness : null,
      qty: Number.isFinite(qty) ? qty : null,
      area_m2: Number.isFinite(area) ? area : null,
      cost_rub: round2(cost),
      sources: {
        totalCell: `${parsed.sheetName}!${parsed.cellRef}`,
        colTotalCell: `${parsed.sheetName}!${leafToTerm.get(leafRef) || ''}`,
        leafCell: `${parsed.sheetName}!${leafRef}`,
        areaCell: areaCell ? `${parsed.sheetName}!${areaCell}` : null,
        unitPriceCell: null,
      },
    });

    leafSum += cost;
  });

  const coverage = totalValue ? leafSum / totalValue : null;
  const aggregated = aggregateItemsByName(items);
  return {
    totalValue,
    leafRefs,
    leafSum,
    coverage,
    items,
    aggregated,
    debug: {
      totalCell: `${parsed.sheetName}!${parsed.cellRef}`,
      formula: totalFormula,
      termCells,
      leafCount: leafRefs.length,
    },
  };
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

function buildCostTableBreakdown(workbook, anchorRef, materials) {
  const anchorParts = decodeRefParts(anchorRef);
  if (!anchorParts) {
    return { anchorValue: 0, leaves: [], details: [], rates: { usable: false }, debug: { reason: 'bad anchor' } };
  }

  const ws = workbook.Sheets[anchorParts.sheetName];
  if (!ws) {
    return { anchorValue: 0, leaves: [], details: [], rates: { usable: false }, debug: { reason: 'sheet missing' } };
  }

  const anchorCell = ws[anchorParts.cellRef];
  const anchorValue = parseNumericValue(anchorCell?.v) || 0;
  const anchorFormula = getCellFormulaString(anchorCell);
  const totalsRefsRaw = anchorFormula ? parseRefsFromFormula(anchorFormula) : [];

  if (!totalsRefsRaw.length) {
    const fallback = inferCostTableFallback(workbook, anchorParts.sheetName, anchorValue, materials);
    if (fallback) {
      return {
        anchorValue,
        leaves: fallback.leafRefs,
        details: (fallback.items || []).map((item) => ({
          name: item.name,
          thickness_mm: item.thicknessMm ?? null,
          length_mm: null,
          width_mm: null,
          qty: null,
          area_m2: item.areaM2 ?? null,
          cost: item.costRub ?? null,
          rowIndex: null,
        })),
        rates: fallback.rates,
        debug: {
          method: 'table-scan',
          ...fallback,
        },
      };
    }
  }

  const totalsFullRefs = expandRanges(totalsRefsRaw, anchorParts.sheetName)
    .map((r) => toFullRef(r, anchorParts.sheetName))
    .filter(Boolean);

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

  const nonZeroTotals = totalsWithValues.filter((x) => typeof x.value === 'number' && x.value !== 0);
  const activeTotals = (nonZeroTotals.length ? nonZeroTotals : totalsWithValues);

  const leafSet = new Set();
  const activeCostCols = new Set();

  activeTotals.forEach((t) => {
    if (!t.p) return;
    activeCostCols.add(t.p.col);
    const refs = parseRefsFromFormula(t.formula);
    const expanded = expandRanges(refs, t.p.sheetName);
    expanded.forEach((r) => {
      const leafFull = toFullRef(r, t.p.sheetName);
      if (!leafFull) return;
      const lp = decodeRefParts(leafFull, t.p.sheetName);
      if (!lp) return;
      if (lp.r === t.p.r && lp.col === t.p.col) return;
      leafSet.add(leafFull);
    });
  });

  const leaves = Array.from(leafSet).sort((a, b) => {
    const pa = decodeRefParts(a, anchorParts.sheetName);
    const pb = decodeRefParts(b, anchorParts.sheetName);
    if (pa.sheetName !== pb.sheetName) return pa.sheetName.localeCompare(pb.sheetName);
    if (pa.r !== pb.r) return pa.r - pb.r;
    return pa.c - pb.c;
  });

  const details = [];
  const rowMap = new Map();
  let leafSum = 0;
  let areaSum = 0;

  leaves.forEach((fullRef) => {
    const p = decodeRefParts(fullRef, anchorParts.sheetName);
    if (!p) return;
    const sheet = workbook.Sheets[p.sheetName];
    const cost = parseNumericValue(sheet?.[p.cellRef]?.v) || 0;
    const rowKey = `${p.sheetName}!${p.r}`;
    const item = rowMap.get(rowKey) || { rowKey, sheetName: p.sheetName, row: p.r, cost: 0 };
    item.cost += cost;
    rowMap.set(rowKey, item);
    leafSum += cost;
  });

  rowMap.forEach((item) => {
    const sheet = workbook.Sheets[item.sheetName];
    const row = item.row;
    const nameVal = sheet?.[`A${row + 1}`]?.v;
    const thicknessVal = parseNumericValue(sheet?.[`B${row + 1}`]?.v);
    const lengthVal = parseNumericValue(sheet?.[`C${row + 1}`]?.v);
    const widthVal = parseNumericValue(sheet?.[`D${row + 1}`]?.v);
    const qtyVal = parseNumericValue(sheet?.[`I${row + 1}`]?.v);

    const qty = Number.isFinite(qtyVal) ? qtyVal : 0;
    const area = (lengthVal && widthVal && qty)
      ? (lengthVal * widthVal * qty) / 1e6
      : null;

    details.push({
      name: String(nameVal ?? `Строка ${row + 1}`),
      thickness_mm: thicknessVal || null,
      length_mm: lengthVal || null,
      width_mm: widthVal || null,
      qty: qty || null,
      area_m2: area,
      cost: round2(item.cost || 0),
      rowIndex: row + 1,
    });

    if (area) areaSum += area;
  });

  const byThickness = {};
  details.forEach((it) => {
    if (!it.area_m2 || !it.cost) return;
    const key = it.thickness_mm ? `${Math.round(it.thickness_mm)}mm` : 'unknown';
    if (!byThickness[key]) byThickness[key] = { area: 0, cost: 0 };
    byThickness[key].area += it.area_m2;
    byThickness[key].cost += it.cost;
  });

  const rateByThickness = {};
  Object.keys(byThickness).forEach((k) => {
    const a = byThickness[k].area;
    const c = byThickness[k].cost;
    if (a > 0 && c > 0) rateByThickness[k] = c / a;
  });

  const avgRate = areaSum > 0 ? (leafSum / areaSum) : null;
  const coverage = anchorValue > 0 ? (leafSum / anchorValue) : null;
  const usable = Boolean(avgRate && avgRate > 0 && coverage && coverage > 0.95 && coverage < 1.05);

  const debug = {
    anchorValue,
    leafCount: leaves.length,
    leafSum,
    areaSum,
    coverage,
    leafRefs: leaves.slice(0, 200).join(', '),
    costColsLetters: Array.from(activeCostCols).join(', ') || '—',
  };

  return {
    anchorValue,
    leaves,
    details,
    rates: { usable, avgRate, byThickness: rateByThickness },
    debug,
  };
}

function buildCalcSummary(workbook, anchors, materials) {
  const baseValues = {
    weight: anchors.weightRef ? readAnchorValue(workbook, anchors.weightRef) : null,
    totalCost: anchors.totalCostRef ? readAnchorValue(workbook, anchors.totalCostRef) : null,
    dsp: anchors.dspRef ? readAnchorValue(workbook, anchors.dspRef) : null,
    edge: anchors.edgeRef ? readAnchorValue(workbook, anchors.edgeRef) : null,
    plastic: anchors.plasticRef ? readAnchorValue(workbook, anchors.plasticRef) : null,
    fabric: anchors.fabricRef ? readAnchorValue(workbook, anchors.fabricRef) : null,
    hwImp: anchors.hwImpRef ? readAnchorValue(workbook, anchors.hwImpRef) : null,
    hwRep: anchors.hwRepRef ? readAnchorValue(workbook, anchors.hwRepRef) : null,
    pack: anchors.packRef ? readAnchorValue(workbook, anchors.packRef) : null,
    labor: anchors.laborRef ? readAnchorValue(workbook, anchors.laborRef) : null,
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
    const dsp = buildChipboardBreakdown(workbook, anchors.dspRef);
    summary.breakdown.dsp = {
      usable: dsp.coverage !== null && dsp.coverage >= 0.95,
      details: dsp.items,
      aggregated: dsp.aggregated,
      leaves: dsp.leafRefs,
      leafCount: dsp.leafRefs?.length ?? 0,
      leafSum: dsp.leafSum ?? 0,
      totalValue: dsp.totalValue ?? 0,
      coverage: dsp.coverage ?? null,
      reason: dsp.debug?.reason || null,
    };
    summary.rates.dsp = buildDspRates(dsp.items || []);
    summary.debug.dsp = dsp.debug;
    if (summary.debug.dsp) {
      console.log('[DSP] totalCell:', summary.debug.dsp.totalCell, 'formula:', summary.debug.dsp.formula, 'termCells:', summary.debug.dsp.termCells, 'leafCount:', summary.debug.dsp.leafCount);
    }
  }

  if (anchors.edgeRef) {
    const edge = buildCostTableBreakdown(workbook, anchors.edgeRef, materials);
    summary.breakdown.edge = {
      usable: edge.rates.usable,
      details: edge.details,
      leaves: edge.leaves,
      leafCount: edge.debug?.leafCount ?? edge.leaves?.length ?? 0,
      leafSum: edge.debug?.leafSum ?? 0,
      totalValue: edge.anchorValue ?? 0,
      coverage: edge.debug?.coverage ?? null,
    };
    summary.rates.edge = edge.rates;
    summary.debug.edge = edge.debug;
  }

  summary.debug.hardwareAnchors = {
    hwImpAnchorRef: anchors.hwImpRef || null,
    hwRepAnchorRef: anchors.hwRepRef || null,
  };

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
  const keywords = ['прямые затраты', 'стоимость расчета суммарно'];
  for (let rowIndex = 0; rowIndex < json.length; rowIndex += 1) {
    const row = json[rowIndex] || [];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cellText = normalizeLabelText(row[colIndex]);
      if (!keywords.some((keyword) => cellText.includes(normalizeLabelText(keyword)))) continue;
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

function findFurnitureHeaderRow(json) {
  for (let row = 0; row < json.length; row += 1) {
    const rowData = json[row] || [];
    const a = normalizeLabelText(rowData[0]);
    const b = normalizeLabelText(rowData[1]);
    const c = normalizeLabelText(rowData[2]);
    const d = normalizeLabelText(rowData[3]);
    if (a.includes('коэф') && b.includes('код') && c.includes('кол') && d.includes('наимен')) {
      return row + 1;
    }
  }
  return null;
}

function parseFurniture(sheet, mapping) {
  if (!sheet) return [];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  const detectedHeader = findFurnitureHeaderRow(json);
  const headerRow = detectedHeader || mapping.furnitureHeaderRow || null;
  if (!headerRow) return [];

  const hasTemplateHeader = Boolean(detectedHeader);
  const columns = hasTemplateHeader
    ? {
      coef: 0,
      code: 1,
      qty: 2,
      name: 3,
      unit: 4,
      origin: 5,
      priceCurrent: 6,
      priceUpdated: 7,
      sum: 8,
    }
    : {
      coef: null,
      code: mapping.furnitureCodeCol,
      qty: mapping.furnitureQtyCol,
      name: mapping.furnitureNameCol,
      unit: mapping.furnitureUnitCol,
      origin: null,
      priceCurrent: mapping.furniturePriceCol,
      priceUpdated: null,
      sum: null,
    };

  const items = [];
  let emptyRows = 0;
  for (let row = headerRow + 1; row < headerRow + 400; row += 1) {
    const rowData = json[row - 1];
    if (!rowData) continue;
    const code = rowData[columns.code];
    const name = rowData[columns.name];
    if (!code && !name) {
      emptyRows += 1;
      if (hasTemplateHeader || emptyRows >= 5) break;
      continue;
    }
    emptyRows = 0;
    const coef = parseNumericValue(rowData[columns.coef]) || 1;
    const qty = parseNumericValue(rowData[columns.qty]) || 0;
    const unit = rowData[columns.unit] || 'шт';
    const originRaw = parseNumericValue(rowData[columns.origin]);
    const origin = Number.isFinite(originRaw) ? originRaw : null;
    const basePrice = parseNumericValue(rowData[columns.priceUpdated])
      ?? parseNumericValue(rowData[columns.priceCurrent]);
    const sumCell = columns.sum != null ? parseNumericValue(rowData[columns.sum]) : null;
    const sumFromJ = parseNumericValue(rowData[9]);
    const sumValue = Number.isFinite(sumCell)
      ? sumCell
      : (Number.isFinite(sumFromJ) ? sumFromJ : null);
    const sumColumnIndex = Number.isFinite(sumCell) ? columns.sum : (Number.isFinite(sumFromJ) ? 9 : null);
    let price = Number.isFinite(basePrice) ? basePrice : 0;
    let priceDerivedFromSum = false;
    if ((!price || !Number.isFinite(price)) && qty > 0 && Number.isFinite(sumValue) && sumValue > 0) {
      const denom = qty * (Number(coef) || 1);
      if (denom > 0) {
        price = sumValue / denom;
        priceDerivedFromSum = true;
      }
    }
    const sum = Number.isFinite(sumValue) ? sumValue : qty * price * coef;
    items.push({
      code,
      name,
      qty,
      unit,
      origin,
      price,
      sum,
      coef,
      priceDerivedFromSum,
      priceDerivedFromSumColumn: priceDerivedFromSum && sumColumnIndex != null ? colIndexToLetter(sumColumnIndex) : null,
    });
  }
  return items;
}

function findFurnitureSheetName(workbook, preferredName) {
  if (preferredName && workbook.Sheets[preferredName]) return preferredName;
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    if (findFurnitureHeaderRow(json)) return sheetName;
  }
  return preferredName || null;
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

  const headerRow = findFurnitureHeaderRow(json);
  if (headerRow) {
    mapping.furnitureHeaderRow = headerRow;
    mapping.furnitureCodeCol = 1;
    mapping.furnitureQtyCol = 2;
    mapping.furnitureNameCol = 3;
    mapping.furnitureUnitCol = 4;
    mapping.furniturePriceCol = 7;
    return mapping;
  }

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
  const furnitureSheetName = findFurnitureSheetName(workbook, mapping.furnitureSheet);
  const furnitureSheet = furnitureSheetName ? workbook.Sheets[furnitureSheetName] : null;
  const furniture = parseFurniture(furnitureSheet, mapping);
  const dims = parseDimensions(sheet, mapping);
  const anchors = resolveAnchors(state.calcSummary?.anchors, mapping.anchorOverrides);
  const calcSummary = buildCalcSummary(workbook, anchors, materials);
  const baseValues = calcSummary.baseValues || {};
  if (Array.isArray(furniture) && furniture.length) {
    const hwTotals = furniture.reduce((acc, row) => {
      const sum = Number(row.sum || 0);
      if (row.origin === 0) acc.hwImp += sum;
      if (row.origin === 1) acc.hwRep += sum;
      return acc;
    }, { hwImp: 0, hwRep: 0 });
    calcSummary.debug.hardwareTable = {
      hwImp: hwTotals.hwImp,
      hwRep: hwTotals.hwRep,
      total: hwTotals.hwImp + hwTotals.hwRep,
    };
    if (!Number.isFinite(baseValues.hwImp)) baseValues.hwImp = hwTotals.hwImp;
    if (!Number.isFinite(baseValues.hwRep)) baseValues.hwRep = hwTotals.hwRep;
  }
  const baseCostFromAnchors = baseValues.totalCost;
  const baseCost = Number.isFinite(baseCostFromAnchors)
    ? baseCostFromAnchors
    : (mapping.baseCostCell ? readCellNumber(sheet, mapping.baseCostCell) : null);
  const materialSumCandidate = [baseValues.dsp, baseValues.edge, baseValues.plastic, baseValues.fabric]
    .some((v) => Number.isFinite(v))
    ? [baseValues.dsp, baseValues.edge, baseValues.plastic, baseValues.fabric]
      .reduce((sum, v) => sum + (Number(v) || 0), 0)
    : null;
  const baseMaterialCost = Number.isFinite(materialSumCandidate)
    ? materialSumCandidate
    : calculatePrice(corpus, materials);

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

const MATERIAL_DENSITY = 730;
const MAX_SECTION_WIDTH = 1200;
const PARTITION_THRESHOLD = 800;

function getMaterialDensity(materialName) {
  const mat = String(materialName || '').toLowerCase();
  if (mat.includes('лдсп') || mat.includes('дсп')) return 730;
  if (mat.includes('мдф')) return 750;
  if (mat.includes('фанер')) return 650;
  if (mat.includes('двп') || mat.includes('оргалит')) return 850;
  if (mat.includes('стекл')) return 2500;
  return MATERIAL_DENSITY;
}

function inferDensityFromSpec(spec) {
  const baseWeight = Number(spec?.calcSummary?.baseValues?.weight);
  if (!Number.isFinite(baseWeight) || baseWeight <= 0) return MATERIAL_DENSITY;
  const volume = (spec?.corpus || []).reduce((sum, part) => {
    if (!part.thickness || !part.length_mm || !part.width_mm || !part.qty) return sum;
    const areaM2 = (part.length_mm / 1000) * (part.width_mm / 1000);
    const thicknessM = part.thickness / 1000;
    return sum + areaM2 * thicknessM * part.qty;
  }, 0);
  if (!volume) return MATERIAL_DENSITY;
  const density = baseWeight / volume;
  return Number.isFinite(density) && density > 0 ? density : MATERIAL_DENSITY;
}

function calculateWeight(parts, density = MATERIAL_DENSITY) {
  let totalKg = 0;
  parts.forEach((part) => {
    if (!part.thickness || !part.length_mm || !part.width_mm || !part.qty) return;
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

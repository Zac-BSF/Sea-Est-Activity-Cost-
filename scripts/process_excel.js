/**
 * Process production Excel files into structured JSON for the Activity Cost Dashboard.
 *
 * Usage:
 *   node scripts/process_excel.js "path/to/excel_file.xlsx"
 *   node scripts/process_excel.js "path/to/excel_file.xlsx" --append
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const LABOR_RATE = 22.00;
const OUTPUT_PATH = path.join(__dirname, '..', 'data', 'production_data_v2.json');

const PROTEIN_PRICES = {
  "2026-03-09": { skin_on: 6.27, abf: 6.55, coho: 5.45, steelhead: 5.90, sockeye: 10.47, grouper: 10.55, snapper: 7.97 },
  "2026-03-16": { skin_on: 6.27, abf: 6.55, coho: 5.45, steelhead: 5.90, sockeye: 10.47, grouper: 10.55, snapper: 7.97 },
  "2026-03-23": { skin_on: 6.37, abf: 6.68, coho: 5.45, steelhead: 5.90, sockeye: 10.47, grouper: 10.55, snapper: 7.97 },
  "2026-03-30": { skin_on: 6.37, abf: 6.68, coho: 5.45, steelhead: 5.90, sockeye: 10.47, grouper: 10.55, snapper: 7.97 },
  "2026-04-06": { skin_on: 6.37, abf: 6.68, coho: 5.45, steelhead: 5.90, sockeye: 10.47, grouper: 10.55, snapper: 7.97 },
  "2026-04-13": { skin_on: 6.46, abf: 6.74, coho: 5.45, steelhead: 5.90, sockeye: 10.47, grouper: 10.55, snapper: 7.97 },
};

const TARGET_COSTS = {
  "2-4 lb Skin-On Atlantic Salmon Fillets": 7.52,
  "2-4 lb Skin-On ABF Atlantic Salmon Fillets": 7.84,
};

const ACTIVITY_NAMES = {
  "Skinner": "Skinning",
  "Slicer Skin-on": "Slicing - Skin-On Salmon",
  "Slicer Skinless": "Slicing - Skinless Salmon",
  "Stripping": "Stripping",
};

// ---- Helpers ----

function round(v, decimals) {
  if (v == null) return null;
  const f = Math.pow(10, decimals);
  return Math.round(v * f) / f;
}

function safeFloat(v) {
  if (v == null) return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).trim());
  return isNaN(n) ? null : n;
}

function getWeekMonday(dt) {
  const d = new Date(dt);
  const day = d.getUTCDay();
  const diff = (day === 0 ? -6 : 1) - day;
  d.setUTCDate(d.getUTCDate() + diff);
  return d.toISOString().slice(0, 10);
}

function getWeekLabel(dt) {
  const d = new Date(dt);
  const jan4 = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const dayOfYear = Math.floor((d - new Date(Date.UTC(d.getUTCFullYear(), 0, 1))) / 86400000) + 1;
  const jan4DayOfWeek = jan4.getUTCDay() || 7;
  const weekNum = Math.ceil((dayOfYear + jan4DayOfWeek - 1) / 7);
  const isoWeek = Math.max(1, weekNum);
  return `${d.getUTCFullYear()}-W${String(isoWeek).padStart(2, '0')}`;
}

function formatDate(dt) {
  const d = new Date(dt);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function isDateValue(v) {
  if (v instanceof Date) return true;
  if (typeof v === 'number' && v > 40000 && v < 60000) return true;
  return false;
}

function toDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + v * 86400000);
  }
  return null;
}

function parseTime(tStr) {
  if (!tStr) return null;
  tStr = String(tStr).trim().toUpperCase().replace(/\s+/g, ' ');
  const m12 = tStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
  if (m12) {
    let h = parseInt(m12[1]);
    const min = parseInt(m12[2]);
    const ampm = m12[3];
    if (ampm === 'PM' && h !== 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return h + min / 60.0;
  }
  const m12b = tStr.match(/^(\d{1,2}):(\d{2})(AM|PM)$/);
  if (m12b) {
    let h = parseInt(m12b[1]);
    const min = parseInt(m12b[2]);
    const ampm = m12b[3];
    if (ampm === 'PM' && h !== 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return h + min / 60.0;
  }
  const m24 = tStr.match(/^(\d{1,2}):(\d{2})$/);
  if (m24) {
    return parseInt(m24[1]) + parseInt(m24[2]) / 60.0;
  }
  return null;
}

function parseLaborTime(laborStr) {
  if (!laborStr || typeof laborStr !== 'string') return null;
  laborStr = laborStr.trim();
  if (!laborStr) return null;

  let breakMinutes = 0;
  const breakMatch = laborStr.match(/(\d+)['’]\s*BREAK/i);
  if (breakMatch) breakMinutes = parseInt(breakMatch[1]);

  const parts = laborStr.split(/\s*-\s*/);
  const times = [];
  for (const p of parts) {
    const trimmed = p.trim();
    if (!trimmed || /BREAK|LUNCH/i.test(trimmed)) continue;
    const t = parseTime(trimmed);
    if (t !== null) times.push(t);
  }

  if (times.length < 2) return null;

  let totalHours = 0;
  for (let i = 0; i < times.length - 1; i += 2) {
    let start = times[i];
    let end = times[i + 1];
    if (end < start) end += 12;
    const diff = end - start;
    if (diff > 0) totalHours += diff;
  }

  totalHours -= breakMinutes / 60.0;
  return totalHours > 0 ? totalHours : null;
}

function normalizeSupplier(s) {
  if (!s || typeof s !== 'string') return null;
  s = s.trim().toUpperCase().replace(/\s+/g, ' ').replace(/[`']/g, '');
  if (['MULTIX', 'MULTI X', 'MULTI  X'].includes(s)) return 'Multi-X';
  if (['AQUA', 'AQUA'].includes(s)) return 'AquaChile';
  if (s === 'CERMAQ') return 'Cermaq';
  if (s === 'BLUGLACIER') return 'BluGlacier';
  if (s === 'TRAPANANDA') return 'Trapananda';
  return s;
}

function mean(arr) {
  if (!arr.length) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function median(arr) {
  if (!arr.length) return null;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function stdev(arr) {
  if (arr.length < 2) return 0;
  const m = mean(arr);
  const variance = arr.reduce((sum, v) => sum + (v - m) ** 2, 0) / (arr.length - 1);
  return Math.sqrt(variance);
}

function cellVal(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
  const cell = ws[addr];
  return cell ? cell.v : null;
}

// ---- Tab Parsers ----

function processSkinner(ws) {
  const records = [];
  let currentDate = null;
  let currentSupplier = null;
  let currentPeople = null;
  const range = XLSX.utils.decode_range(ws['!ref']);

  for (let row = 5; row <= range.e.r + 1; row++) {
    const dateVal = cellVal(ws, row, 1);
    const supplierVal = cellVal(ws, row, 2);
    const lotVal = cellVal(ws, row, 3);
    const palletVal = cellVal(ws, row, 4);
    const incoming = safeFloat(cellVal(ws, row, 6));
    const outgoing = safeFloat(cellVal(ws, row, 7));
    const productFormat = cellVal(ws, row, 9);
    const peopleVal = cellVal(ws, row, 11);
    const laborVal = cellVal(ws, row, 12);

    if (isDateValue(dateVal)) {
      currentDate = toDate(dateVal);
      currentPeople = null;
    }
    if (supplierVal && typeof supplierVal === 'string' && supplierVal.trim()) {
      currentSupplier = supplierVal;
    }
    if (peopleVal != null) {
      const p = parseInt(peopleVal);
      if (!isNaN(p)) currentPeople = p;
    }

    if (incoming == null || outgoing == null) continue;
    if (incoming <= 0 || outgoing <= 0) continue;
    if (!currentDate) continue;

    const hours = parseLaborTime(laborVal != null ? String(laborVal) : null);
    if (hours == null || currentPeople == null) continue;

    let fmt = productFormat ? String(productFormat).trim() : '';
    fmt = fmt.toUpperCase() === 'ABF' ? 'ABF' : 'Conventional';

    const totalLaborHours = currentPeople * hours;
    const laborCost = totalLaborHours * LABOR_RATE;
    const costPerLb = outgoing > 0 ? laborCost / outgoing : null;
    const yieldPct = incoming > 0 ? (outgoing / incoming * 100) : null;

    records.push({
      activity: 'Skinner',
      date: formatDate(currentDate),
      week: getWeekLabel(currentDate),
      supplier: normalizeSupplier(currentSupplier),
      lot: lotVal != null ? String(lotVal).trim() : null,
      pallet: palletVal != null ? String(palletVal).trim() : null,
      product_format: fmt,
      incoming_lbs: round(incoming, 2),
      finished_lbs: round(outgoing, 2),
      yield_pct: yieldPct != null ? round(yieldPct, 2) : null,
      people: currentPeople,
      hours_worked: round(hours, 4),
      total_labor_hours: round(totalLaborHours, 4),
      labor_cost: round(laborCost, 2),
      cost_per_finished_lb: costPerLb != null ? round(costPerLb, 4) : null,
    });
  }
  return records;
}

function processSlicerSkinOn(ws) {
  const records = [];
  let currentDate = null;
  let currentSupplier = null;
  let currentPeople = null;
  const range = XLSX.utils.decode_range(ws['!ref']);

  for (let row = 6; row <= range.e.r + 1; row++) {
    const dateVal = cellVal(ws, row, 1);
    const supplierVal = cellVal(ws, row, 2);
    const lotVal = cellVal(ws, row, 3);
    const palletVal = cellVal(ws, row, 4);
    const incoming = safeFloat(cellVal(ws, row, 6));
    const sides = safeFloat(cellVal(ws, row, 7)) || 0;
    const portions = safeFloat(cellVal(ws, row, 8)) || 0;
    const pesto = safeFloat(cellVal(ws, row, 9)) || 0;
    const pieces = safeFloat(cellVal(ws, row, 10)) || 0;
    const productFormat = cellVal(ws, row, 12);
    const peopleVal = cellVal(ws, row, 14);
    const laborVal = cellVal(ws, row, 15);

    if (isDateValue(dateVal)) {
      currentDate = toDate(dateVal);
      currentPeople = null;
    }
    if (supplierVal && typeof supplierVal === 'string' && supplierVal.trim()) {
      currentSupplier = supplierVal;
    }
    if (peopleVal != null) {
      const p = parseInt(peopleVal);
      if (!isNaN(p)) currentPeople = p;
    }

    if (incoming == null || incoming <= 0) continue;
    if (!currentDate) continue;

    const totalOutput = sides + portions + pesto + pieces;
    if (totalOutput <= 0) continue;

    const hours = parseLaborTime(laborVal != null ? String(laborVal) : null);
    if (hours == null || currentPeople == null) continue;

    let fmt = productFormat ? String(productFormat).trim() : '';
    if (!fmt || fmt === 'None') fmt = 'Skin on (ungraded)';

    const totalLaborHours = currentPeople * hours;
    const laborCost = totalLaborHours * LABOR_RATE;
    const costPerLb = totalOutput > 0 ? laborCost / totalOutput : null;
    const yieldPct = incoming > 0 ? (totalOutput / incoming * 100) : null;

    records.push({
      activity: 'Slicer Skin-on',
      date: formatDate(currentDate),
      week: getWeekLabel(currentDate),
      supplier: normalizeSupplier(currentSupplier),
      lot: lotVal != null ? String(lotVal).trim() : null,
      pallet: palletVal != null ? String(palletVal).trim() : null,
      product_format: fmt,
      incoming_lbs: round(incoming, 2),
      finished_lbs: round(totalOutput, 2),
      yield_pct: yieldPct != null ? round(yieldPct, 2) : null,
      people: currentPeople,
      hours_worked: round(hours, 4),
      total_labor_hours: round(totalLaborHours, 4),
      labor_cost: round(laborCost, 2),
      cost_per_finished_lb: costPerLb != null ? round(costPerLb, 4) : null,
    });
  }
  return records;
}

function processSlicerSkinless(ws) {
  const records = [];
  let currentDate = null;
  let currentSupplier = null;
  let currentPeople = null;
  const range = XLSX.utils.decode_range(ws['!ref']);

  for (let row = 6; row <= range.e.r + 1; row++) {
    const dateVal = cellVal(ws, row, 1);
    const supplierVal = cellVal(ws, row, 2);
    const lotVal = cellVal(ws, row, 3);
    const palletVal = cellVal(ws, row, 4);
    const incoming = safeFloat(cellVal(ws, row, 5));
    const skinlessOut = safeFloat(cellVal(ws, row, 6)) || 0;
    const piecesOut = safeFloat(cellVal(ws, row, 7)) || 0;
    const productFormat = cellVal(ws, row, 9);
    const peopleVal = cellVal(ws, row, 11);
    const laborVal = cellVal(ws, row, 12);

    if (isDateValue(dateVal)) {
      currentDate = toDate(dateVal);
      currentPeople = null;
    } else if (dateVal && typeof dateVal === 'string') {
      const dateMatch = String(dateVal).match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
      if (dateMatch) {
        const parts = dateMatch[1].split('/');
        currentDate = new Date(Date.UTC(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1])));
        currentPeople = null;
      }
    }

    if (supplierVal && typeof supplierVal === 'string' && supplierVal.trim()) {
      currentSupplier = supplierVal;
    }
    if (peopleVal != null) {
      const p = parseInt(peopleVal);
      if (!isNaN(p)) currentPeople = p;
    }

    if (incoming == null || incoming <= 0) continue;
    const totalOutput = skinlessOut + piecesOut;
    if (totalOutput <= 0) continue;
    if (!currentDate) continue;

    const hours = parseLaborTime(laborVal != null ? String(laborVal) : null);
    if (hours == null || currentPeople == null) continue;

    let fmt = productFormat ? String(productFormat).trim().replace(/\s+/g, ' ') : '';
    if (!fmt || fmt === 'None') fmt = 'Conventional';
    if (fmt === 'From skin on') fmt = 'From Skin-on (Conventional)';
    else if (fmt === 'From skin on ABF') fmt = 'From Skin-on (ABF)';

    const totalLaborHours = currentPeople * hours;
    const laborCost = totalLaborHours * LABOR_RATE;
    const costPerLb = totalOutput > 0 ? laborCost / totalOutput : null;
    const yieldPct = incoming > 0 ? (totalOutput / incoming * 100) : null;

    records.push({
      activity: 'Slicer Skinless',
      date: formatDate(currentDate),
      week: getWeekLabel(currentDate),
      supplier: normalizeSupplier(currentSupplier),
      lot: lotVal != null ? String(lotVal).trim() : null,
      pallet: palletVal != null ? String(palletVal).trim() : null,
      product_format: fmt,
      incoming_lbs: round(incoming, 2),
      finished_lbs: round(totalOutput, 2),
      yield_pct: yieldPct != null ? round(yieldPct, 2) : null,
      people: currentPeople,
      hours_worked: round(hours, 4),
      total_labor_hours: round(totalLaborHours, 4),
      labor_cost: round(laborCost, 2),
      cost_per_finished_lb: costPerLb != null ? round(costPerLb, 4) : null,
    });
  }
  return records;
}

function processStripping(ws) {
  const records = [];
  const range = XLSX.utils.decode_range(ws['!ref']);

  for (let row = 6; row <= range.e.r + 1; row++) {
    const dateVal = cellVal(ws, row, 1);
    const productVal = cellVal(ws, row, 2);
    const lbsVal = safeFloat(cellVal(ws, row, 3));
    const peopleVal = cellVal(ws, row, 4);
    const laborVal = cellVal(ws, row, 5);

    if (!isDateValue(dateVal)) continue;
    const dt = toDate(dateVal);
    if (lbsVal == null || lbsVal <= 0) continue;

    const people = parseInt(peopleVal);
    if (isNaN(people)) continue;

    const hours = parseLaborTime(laborVal != null ? String(laborVal) : null);
    if (hours == null) continue;

    const fmt = productVal ? String(productVal).trim() : 'Unknown';
    const totalLaborHours = people * hours;
    const laborCost = totalLaborHours * LABOR_RATE;
    const costPerLb = lbsVal > 0 ? laborCost / lbsVal : null;

    records.push({
      activity: 'Stripping',
      date: formatDate(dt),
      week: getWeekLabel(dt),
      supplier: null,
      lot: null,
      pallet: null,
      product_format: fmt,
      incoming_lbs: round(lbsVal, 2),
      finished_lbs: round(lbsVal, 2),
      yield_pct: null,
      people,
      hours_worked: round(hours, 4),
      total_labor_hours: round(totalLaborHours, 4),
      labor_cost: round(laborCost, 2),
      cost_per_finished_lb: costPerLb != null ? round(costPerLb, 4) : null,
    });
  }
  return records;
}

// ---- Enrichment ----

function getProteinPrice(dt, activity, productFormat) {
  const monday = getWeekMonday(dt);
  const prices = PROTEIN_PRICES[monday];
  if (!prices) return null;

  const fmtLower = (productFormat || '').toLowerCase();

  if (activity === 'Skinner') return fmtLower.includes('abf') ? prices.abf : prices.skin_on;
  if (activity === 'Slicer Skin-on') return prices.skin_on;
  if (activity === 'Slicer Skinless') return fmtLower.includes('abf') ? prices.abf : prices.skin_on;
  if (activity === 'Stripping') {
    if (fmtLower.includes('coho')) return prices.coho;
    if (fmtLower.includes('steelhead')) return prices.steelhead;
    if (fmtLower.includes('sockeye')) return prices.sockeye;
    if (fmtLower.includes('grouper')) return prices.grouper;
    if (fmtLower.includes('snapper')) return prices.snapper;
    if (fmtLower.includes('salmon') || fmtLower.includes('skin on')) return prices.skin_on;
    return null;
  }
  return null;
}

function enrichWithProteinCost(r) {
  const dt = new Date(r.date + 'T00:00:00Z');
  const price = getProteinPrice(dt, r.activity, r.product_format);
  r.raw_protein_cost_per_lb = price;

  if (r.activity === 'Stripping') {
    r.protein_cost_per_finished_lb = price;
    r.yield_loss_cost_per_lb = 0;
    if (price && r.cost_per_finished_lb) {
      r.total_cost_per_finished_lb = round(price + r.cost_per_finished_lb, 4);
    } else if (r.cost_per_finished_lb) {
      r.total_cost_per_finished_lb = r.cost_per_finished_lb;
    } else {
      r.total_cost_per_finished_lb = null;
    }
  } else if (price && r.yield_pct && r.yield_pct > 0) {
    const proteinCostPerFinished = price / (r.yield_pct / 100.0);
    const yieldLossCost = proteinCostPerFinished - price;
    r.protein_cost_per_finished_lb = round(proteinCostPerFinished, 4);
    r.yield_loss_cost_per_lb = round(yieldLossCost, 4);
    if (r.cost_per_finished_lb) {
      r.total_cost_per_finished_lb = round(proteinCostPerFinished + r.cost_per_finished_lb, 4);
    } else {
      r.total_cost_per_finished_lb = null;
    }
  } else {
    r.protein_cost_per_finished_lb = null;
    r.yield_loss_cost_per_lb = null;
    r.total_cost_per_finished_lb = null;
  }
  return r;
}

function classifyRecord(r) {
  const fmtLower = (r.product_format || '').toLowerCase();
  const prevFrozen = ['grouper', 'snapper', 'steelhead', 'coho', 'sockeye'];

  if (r.activity === 'Stripping' && prevFrozen.some(sp => fmtLower.includes(sp))) {
    r.classification = 'Previously Frozen';
  } else {
    r.classification = 'Fresh';
  }

  if (r.activity === 'Skinner') {
    r.product_format = fmtLower === 'abf' ? '2-4 lb Skin-On ABF Atlantic Salmon Fillets' : '2-4 lb Skin-On Atlantic Salmon Fillets';
  } else if (r.activity === 'Slicer Skin-on') {
    if (r.product_format.includes('3-4')) r.product_format = '3-4 lb Skin-On Atlantic Salmon Fillets';
    else if (r.product_format.includes('2-3')) r.product_format = '2-3 lb Skin-On Atlantic Salmon Fillets';
    else r.product_format = '2-4 lb Skin-On Atlantic Salmon Fillets';
  } else if (r.activity === 'Slicer Skinless') {
    if (fmtLower.includes('abf') || r.product_format.includes('From Skin-on (ABF)')) {
      r.product_format = '2-4 lb Skin-On ABF Atlantic Salmon Fillets';
    } else {
      r.product_format = '2-4 lb Skin-On Atlantic Salmon Fillets';
    }
  } else if (r.activity === 'Stripping') {
    if (fmtLower.includes('skin on') && fmtLower.includes('salmon')) {
      r.product_format = '2-4 lb Skin-On Atlantic Salmon Fillets';
    }
  }
  return r;
}

function computeChainedCosts(records) {
  // Step 1: Weekly avg stripping labor for Fresh Atlantic salmon
  const stripLaborByWeek = {};
  for (const r of records) {
    if (r.activity !== 'Stripping' || r.classification !== 'Fresh') continue;
    if (!r.cost_per_finished_lb) continue;
    if (!stripLaborByWeek[r.week]) stripLaborByWeek[r.week] = [];
    stripLaborByWeek[r.week].push(r.cost_per_finished_lb);
  }
  const avgStripLabor = {};
  for (const [w, vals] of Object.entries(stripLaborByWeek)) {
    avgStripLabor[w] = mean(vals);
  }

  // Step 2: Recompute Skinner records with upstream stripping cost
  const skinnerOutputByWeekProduct = {};
  for (const r of records) {
    if (r.activity !== 'Skinner') continue;
    if (!r.raw_protein_cost_per_lb || !r.yield_pct) continue;

    const rawPrice = r.raw_protein_cost_per_lb;
    const stripLabor = avgStripLabor[r.week] || 0;
    const inputCost = rawPrice + stripLabor;
    const yieldFrac = r.yield_pct / 100.0;
    const yieldedInputCost = inputCost / yieldFrac;
    const labor = r.cost_per_finished_lb || 0;
    const outputCost = yieldedInputCost + labor;

    r.upstream_strip_labor = round(stripLabor, 4);
    r.input_cost_per_lb = round(inputCost, 4);
    r.total_cost_per_finished_lb = round(outputCost, 4);
    r.yield_loss_cost_per_lb = round(yieldedInputCost - inputCost, 4);
    r.protein_cost_per_finished_lb = round(yieldedInputCost, 4);

    const kpi = TARGET_COSTS[r.product_format];
    if (kpi) {
      const spread = kpi - outputCost;
      r.target_cost = kpi;
      r.production_spread_per_lb = round(spread, 4);
      r.extended_production_spread = round(spread * r.finished_lbs, 2);
    } else {
      r.target_cost = null;
      r.production_spread_per_lb = null;
      r.extended_production_spread = null;
    }

    const key = `${r.week}|${r.product_format}`;
    if (!skinnerOutputByWeekProduct[key]) skinnerOutputByWeekProduct[key] = [];
    skinnerOutputByWeekProduct[key].push(outputCost);
  }

  const avgSkinnerOutput = {};
  for (const [k, v] of Object.entries(skinnerOutputByWeekProduct)) {
    avgSkinnerOutput[k] = mean(v);
  }

  // Step 3: Recompute Slicer Skinless with Skinner output as input cost
  for (const r of records) {
    if (r.activity !== 'Slicer Skinless') continue;
    if (!r.yield_pct) continue;

    let upstreamCost = avgSkinnerOutput[`${r.week}|${r.product_format}`];
    if (!upstreamCost) {
      const weekCosts = Object.entries(avgSkinnerOutput)
        .filter(([k]) => k.startsWith(r.week + '|'))
        .map(([, v]) => v);
      if (weekCosts.length) upstreamCost = mean(weekCosts);
    }
    if (!upstreamCost) continue;

    const yieldFrac = r.yield_pct / 100.0;
    const yieldedInputCost = upstreamCost / yieldFrac;
    const labor = r.cost_per_finished_lb || 0;
    const outputCost = yieldedInputCost + labor;

    r.input_cost_per_lb = round(upstreamCost, 4);
    r.total_cost_per_finished_lb = round(outputCost, 4);
    r.yield_loss_cost_per_lb = round(yieldedInputCost - upstreamCost, 4);
    r.protein_cost_per_finished_lb = round(yieldedInputCost, 4);
    r.raw_protein_cost_per_lb = round(upstreamCost, 4);

    const kpi = TARGET_COSTS[r.product_format];
    if (kpi) {
      const spread = kpi - outputCost;
      r.target_cost = kpi;
      r.production_spread_per_lb = round(spread, 4);
      r.extended_production_spread = round(spread * r.finished_lbs, 2);
    }
  }

  // Step 4: KPI/spread defaults for other activities
  for (const r of records) {
    if (r.activity === 'Slicer Skin-on' && !r.target_cost) {
      r.target_cost = null;
      r.production_spread_per_lb = null;
      r.extended_production_spread = null;
    }
    if (r.activity === 'Stripping') {
      r.target_cost = null;
      r.production_spread_per_lb = null;
      r.extended_production_spread = null;
    }
  }

  return records;
}

// ---- Summary ----

function computeSummary(records) {
  const groups = {};
  for (const r of records) {
    const key = `${r.activity}|${r.product_format}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  }

  const summary = {};
  for (const [key, recs] of Object.entries(groups)) {
    const costs = recs.filter(r => r.cost_per_finished_lb != null && r.cost_per_finished_lb > 0).map(r => r.cost_per_finished_lb);
    const yields = recs.filter(r => r.yield_pct != null).map(r => r.yield_pct);
    const totalFinished = recs.reduce((s, r) => s + r.finished_lbs, 0);

    if (costs.length) {
      const sortedCosts = [...costs].sort((a, b) => a - b);
      const n = sortedCosts.length;
      const p25Idx = Math.floor(n * 0.25);
      const p75Idx = Math.min(Math.floor(n * 0.75), n - 1);

      const totalCosts = recs.filter(r => r.total_cost_per_finished_lb).map(r => r.total_cost_per_finished_lb);
      const yieldLossCosts = recs.filter(r => r.yield_loss_cost_per_lb).map(r => r.yield_loss_cost_per_lb);

      summary[key] = {
        count: recs.length,
        avg_cost: round(mean(costs), 4),
        median_cost: round(median(costs), 4),
        min_cost: round(Math.min(...costs), 4),
        max_cost: round(Math.max(...costs), 4),
        p25_cost: round(sortedCosts[p25Idx], 4),
        p75_cost: round(sortedCosts[p75Idx], 4),
        std_cost: round(stdev(costs), 4),
        avg_yield: yields.length ? round(mean(yields), 2) : null,
        avg_yield_loss_cost: yieldLossCosts.length ? round(mean(yieldLossCosts), 4) : null,
        avg_total_cost: totalCosts.length ? round(mean(totalCosts), 4) : null,
        total_finished_lbs: round(totalFinished, 2),
      };
    }
  }
  return summary;
}

// ---- Main ----

function main() {
  const args = process.argv.slice(2);
  if (!args.length) {
    console.log('Usage: node scripts/process_excel.js <excel_file> [--append]');
    process.exit(1);
  }

  const excelPath = args[0];
  const appendMode = args.includes('--append');

  if (!fs.existsSync(excelPath)) {
    console.log(`File not found: ${excelPath}`);
    process.exit(1);
  }

  console.log(`Reading: ${excelPath}`);
  const wb = XLSX.readFile(excelPath);

  let allRecords = [];

  if (wb.SheetNames.includes('Skinner')) {
    const recs = processSkinner(wb.Sheets['Skinner']);
    console.log(`  Skinner: ${recs.length} records`);
    allRecords.push(...recs);
  }

  const skinOnSheet = wb.SheetNames.find(n => n.trim().startsWith('Slicer for Skin on'));
  if (skinOnSheet) {
    const recs = processSlicerSkinOn(wb.Sheets[skinOnSheet]);
    console.log(`  Slicer Skin-on: ${recs.length} records`);
    allRecords.push(...recs);
  }

  if (wb.SheetNames.includes('Slicer for Skinless')) {
    const recs = processSlicerSkinless(wb.Sheets['Slicer for Skinless']);
    console.log(`  Slicer Skinless: ${recs.length} records`);
    allRecords.push(...recs);
  }

  if (wb.SheetNames.includes('Stripping')) {
    const recs = processStripping(wb.Sheets['Stripping']);
    console.log(`  Stripping: ${recs.length} records`);
    allRecords.push(...recs);
  }

  // Enrich with protein cost first (needs original format names), then classify/rename
  allRecords = allRecords.map(enrichWithProteinCost);
  allRecords = allRecords.map(classifyRecord);

  // Compute chained costs: stripping -> skinning -> slicing
  allRecords = computeChainedCosts(allRecords);

  // Rename activities for display
  for (const r of allRecords) {
    r.activity = ACTIVITY_NAMES[r.activity] || r.activity;
  }

  if (appendMode && fs.existsSync(OUTPUT_PATH)) {
    const existing = JSON.parse(fs.readFileSync(OUTPUT_PATH, 'utf8'));
    const existingKeys = new Set();
    for (const r of existing.records) {
      existingKeys.add(`${r.date}|${r.activity}|${r.lot || ''}|${r.pallet || ''}`);
    }
    let newCount = 0;
    for (const r of allRecords) {
      const k = `${r.date}|${r.activity}|${r.lot || ''}|${r.pallet || ''}`;
      if (!existingKeys.has(k)) {
        existing.records.push(r);
        newCount++;
      }
    }
    allRecords = existing.records;
    console.log(`\nAppend mode: added ${newCount} new records`);
  }

  const summary = computeSummary(allRecords);

  const output = {
    generated_at: new Date().toISOString(),
    labor_rate: LABOR_RATE,
    protein_prices: PROTEIN_PRICES,
    source_file: path.basename(excelPath),
    total_records: allRecords.length,
    records: allRecords.sort((a, b) => a.date.localeCompare(b.date) || a.activity.localeCompare(b.activity)),
    summary,
  };

  fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(output, null, 2));

  console.log(`\nTotal records: ${allRecords.length}`);
  console.log(`Output: ${OUTPUT_PATH}`);
  console.log('\nSummary by Activity|Product:');
  for (const [key, stats] of Object.entries(summary).sort()) {
    console.log(`  ${key}: avg $${stats.avg_cost.toFixed(4)}/lb, range $${stats.min_cost.toFixed(4)}-$${stats.max_cost.toFixed(4)}, n=${stats.count}`);
  }
}

main();

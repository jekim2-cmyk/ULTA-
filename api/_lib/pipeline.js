// ULTA 원본 리포트(Store-Sales_*.xlsx, Sales_Inv_Perf__*.xlsx) 파싱 + 월간/주간 롤업 계산.
// ulta_update_pipeline.py 의 로직을 그대로 JS로 포팅한 것으로, 실제 13주치 원본 파일로
// 파이썬 버전과 1:1 비교 검증을 마쳤습니다 (storeRanking/skuSales/skuWeekly/skuOos/financials/qtdYtd 완전 일치).
import * as XLSX from 'xlsx';

// ---------- 공통 유틸 ----------

export function parseFileDate(filename) {
  const base = filename.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '');
  const dateStr = base.slice(-10);
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() - 1);
  return d.toISOString().slice(0, 10);
}

function findHeaderRow(rows, marker, col = 0) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r && r.length > col && r[col] === marker) return i;
  }
  return null;
}

function num(v) {
  return typeof v === 'number' ? v : null;
}

function round2(v) {
  return Math.round(v * 100) / 100;
}
function round1(v) {
  return Math.round(v * 10) / 10;
}

function sheetRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`시트를 찾을 수 없습니다: ${sheetName} (파일 형식을 확인해주세요)`);
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

// ---------- Store-Sales 파싱 ----------

export function parseStoreSales(buffer, filename) {
  const weekEnd = parseFileDate(filename);
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const rows = sheetRows(wb, 'StoreSalesReport');

  let fiscalYear = null, fiscalWeekNum = null;
  for (let i = 0; i < Math.min(6, rows.length); i++) {
    const r = rows[i];
    if (r && r[0] && String(r[0]).includes('Fiscal Week')) {
      const txt = String(r[0]).split(':').pop().trim();
      fiscalYear = txt.slice(0, 4);
      fiscalWeekNum = txt.slice(4);
      break;
    }
  }

  const hdrIdx = findHeaderRow(rows, 'Store Number');
  if (hdrIdx === null) throw new Error('Store-Sales 파일에서 "Store Number" 헤더를 찾지 못했습니다.');
  const header = rows[hdrIdx];
  const dataRows = rows.slice(hdrIdx + 1);

  const nCols = header.length;
  const blockStarts = [];
  for (let c = 2; c < nCols; c += 4) blockStarts.push(c);
  const lastBlock = blockStarts[blockStarts.length - 1];
  const unitsCol = lastBlock, salesCol = lastBlock + 1;

  const totalRow = dataRows.find((r) => r[1] && String(r[1]).startsWith('Total:'));
  if (!totalRow) throw new Error('Store-Sales 파일에서 합계(Total:) 행을 찾지 못했습니다.');
  const comRow = dataRows.find((r) => r[0] === '0902') || null;
  const storeRows = dataRows.filter(
    (r) => r[0] !== null && r[0] !== '' && r[0] !== undefined && !(r[1] && String(r[1]).startsWith('Total:'))
  );

  const totalUnits = num(totalRow[unitsCol]) || 0;
  const totalSales = round2(num(totalRow[salesCol]) || 0);
  const comUnits = comRow ? num(comRow[unitsCol]) || 0 : 0;
  const comSales = round2(comRow ? num(comRow[salesCol]) || 0 : 0);
  const bmUnits = totalUnits - comUnits;
  const bmSales = round2(totalSales - comSales);
  const storeCount = storeRows.length - (comRow ? 1 : 0);

  const topStores = [];
  for (const r of storeRows) {
    if (r[0] === '0902') continue;
    const sales = num(r[salesCol]) || 0;
    const units = num(r[unitsCol]) || 0;
    if (sales || units) {
      topStores.push({ storeNum: r[0], storeName: r[1], sales: round2(sales), units });
    }
  }

  return {
    weekEndDate: weekEnd, fiscalYear, fiscalWeekNum,
    totalSales, totalUnits, comSales, comUnits, bmSales, bmUnits,
    storeCount, topStores,
  };
}

// ---------- Sales_Inv_Perf 파싱 (Last Closed Week 시트) ----------

const COL_DESC = 2, COL_COST = 8, COL_PRICE = 9;
const COL_UNITS = 10, COL_SALES = 13;
const COL_EOHUNITS = 34;
const COL_EOHDOLLARS = 37, COL_WOS = 41;
const COL_DCEOHUNITS = 42, COL_DCONORDERUNITS = 48;
const COL_STOREEOHUNITS = 97;
const COL_STORECOUNT = 107, COL_STOREOUTS = 108, COL_OOSPCT = 109;

export function parseSalesInvPerfWeek(buffer, filename) {
  const weekEnd = parseFileDate(filename);
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const rows = sheetRows(wb, 'Last Closed Week');
  const hdrIdx = findHeaderRow(rows, 'UPC');
  if (hdrIdx === null) throw new Error('Sales_Inv_Perf 파일에서 "UPC" 헤더를 찾지 못했습니다.');
  const overall = rows[hdrIdx + 1];
  const skuRows = rows.slice(hdrIdx + 2);

  const skus = [];
  for (const r of skuRows) {
    const desc = r[COL_DESC];
    if (!desc) continue;
    const sales = num(r[COL_SALES]);
    if (sales === null) continue;
    const oosRaw = num(r[COL_OOSPCT]);
    skus.push({
      desc: String(desc).trim(),
      cost: num(r[COL_COST]), price: num(r[COL_PRICE]),
      units: num(r[COL_UNITS]), sales: round2(sales),
      eohUsd: num(r[COL_EOHDOLLARS]), wos: num(r[COL_WOS]),
      storeCount: num(r[COL_STORECOUNT]), storeOuts: num(r[COL_STOREOUTS]),
      oosPct: oosRaw !== null ? oosRaw * 100 : null,
    });
  }

  const overallSales = num(overall[COL_SALES]) || 0;
  const overallEoh = num(overall[COL_EOHDOLLARS]) || 0;
  const overallWos = num(overall[COL_WOS]);
  return {
    weekEndDate: weekEnd,
    totalSales: round2(overallSales),
    totalUnits: num(overall[COL_UNITS]),
    invEohUsd: round2(overallEoh),
    invWos: overallWos !== null ? round1(overallWos) : null,
    eohUnits: num(overall[COL_EOHUNITS]),
    dcEohUnits: num(overall[COL_DCEOHUNITS]),
    dcOnOrderUnits: num(overall[COL_DCONORDERUNITS]),
    storeEohUnits: num(overall[COL_STOREEOHUNITS]),
    skus,
  };
}

export function parseQtdYtd(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const result = {};
  for (const [sheetName, key] of [['Quarter to Date', 'qtd'], ['Year to Date', 'ytd']]) {
    const rows = sheetRows(wb, sheetName);
    let label = null;
    for (let i = 0; i < Math.min(8, rows.length); i++) {
      const r = rows[i] || [];
      for (const cell of r) {
        if (typeof cell === 'string' && (cell.startsWith('QTD') || cell.startsWith('YTD'))) label = cell;
      }
    }
    let overallRowIdx = null;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      for (let j = 0; j < r.length; j++) {
        if (r[j] === 'Overall Result') overallRowIdx = i;
      }
    }
    if (overallRowIdx === null) throw new Error(`${sheetName} 시트에서 "Overall Result" 행을 찾지 못했습니다.`);
    const hdrRowIdx = overallRowIdx - 1;
    const header = rows[hdrRowIdx];
    const unitsCol = header.findIndex((v) => typeof v === 'string' && v.includes('TOTAL') && v.includes('Sales TY') && v.includes('Units'));
    const salesCol = header.findIndex((v) => typeof v === 'string' && v.includes('TOTAL') && v.includes('Sales TY') && v.includes('$'));
    const overall = rows[overallRowIdx];
    const salesVal = num(overall[salesCol]);
    result[key] = {
      label,
      totalUnits: num(overall[unitsCol]),
      totalSalesUsd: salesVal !== null ? round2(salesVal) : null,
    };
  }
  return result;
}

// ---------- 월간/주간 롤업 계산 ----------

function monthKey(dateIso) {
  return dateIso.slice(0, 7);
}

export function buildStoreRanking(storeWeeks) {
  const byMonth = {};
  for (const w of storeWeeks) {
    const mk = monthKey(w.weekEndDate);
    if (!byMonth[mk]) byMonth[mk] = { totalSales: 0, storeCount: 0, storeAgg: {}, weeksIncluded: [] };
    const agg = byMonth[mk];
    agg.totalSales += w.totalSales;
    agg.storeCount = Math.max(agg.storeCount, w.storeCount);
    agg.weeksIncluded.push(w.weekEndDate);
    for (const s of w.topStores) {
      const key = s.storeNum;
      if (!agg.storeAgg[key]) agg.storeAgg[key] = { storeNum: s.storeNum, storeName: s.storeName, sales: 0, units: 0 };
      const e = agg.storeAgg[key];
      e.sales += s.sales;
      e.units += s.units;
    }
  }
  const out = {};
  for (const mk of Object.keys(byMonth)) {
    const agg = byMonth[mk];
    const stores = Object.values(agg.storeAgg).sort((a, b) => b.sales - a.sales).slice(0, 20);
    const total = round2(agg.totalSales);
    for (const s of stores) {
      s.sales = round2(s.sales);
      s.pctOfTotal = total ? round2((s.sales / total) * 100) : 0;
    }
    out[mk] = {
      totalSales: total,
      storeCount: agg.storeCount,
      topStores: stores,
      weeksIncluded: [...agg.weeksIncluded].sort(),
    };
  }
  return out;
}

// 주차별 매장 순위 스냅샷 (전주 대비 순위 변동 비교용). 월간 롤업과 달리 주 단위를 그대로 보존.
export function buildStoreWeekly(storeWeeks) {
  const byWeek = {};
  for (const w of storeWeeks) {
    const sorted = [...w.topStores].sort((a, b) => b.sales - a.sales).slice(0, 20);
    const ranked = sorted.map((s, i) => ({
      rank: i + 1, storeNum: s.storeNum, storeName: s.storeName,
      sales: round2(s.sales), units: s.units,
    }));
    byWeek[w.weekEndDate] = { totalSales: w.totalSales, storeCount: w.storeCount, topStores: ranked };
  }
  return byWeek;
}

export function buildSkuSales(invWeeks) {
  const byMonth = {};
  for (const w of invWeeks) {
    const mk = monthKey(w.weekEndDate);
    if (!byMonth[mk]) byMonth[mk] = {};
    const agg = byMonth[mk];
    for (const sku of w.skus) {
      agg[sku.desc] = (agg[sku.desc] || 0) + sku.sales;
    }
  }
  const out = {};
  for (const mk of Object.keys(byMonth)) {
    const agg = byMonth[mk];
    const items = Object.entries(agg).sort((a, b) => b[1] - a[1]);
    const total = round2(items.reduce((s, [, v]) => s + v, 0));
    const top10 = items.slice(0, 10);
    const top3Sales = round2(items.slice(0, 3).reduce((s, [, v]) => s + v, 0));
    out[mk] = {
      totalSales: total,
      skuCount: items.length,
      top3Sales,
      top3Pct: total ? round1((top3Sales / total) * 100) : 0,
      topSkus: top10.map(([d, v]) => ({ desc: d, sales: round2(v), pct: total ? round2((v / total) * 100) : 0 })),
    };
  }
  return out;
}

export function buildSkuWeekly(invWeeks) {
  const byWeek = {};
  for (const w of invWeeks) {
    const items = [];
    for (const sku of w.skus) {
      if (sku.sales === null) continue;
      items.push({ desc: sku.desc, sales: round2(sku.sales), units: sku.units });
    }
    items.sort((a, b) => b.sales - a.sales);
    const total = round2(items.reduce((s, i) => s + i.sales, 0));
    byWeek[w.weekEndDate] = { totalSales: total, skuCount: items.length, skus: items };
  }
  return byWeek;
}

export function buildSkuOos(invWeeks) {
  const byMonth = {};
  for (const w of invWeeks) {
    const mk = monthKey(w.weekEndDate);
    if (!byMonth[mk]) byMonth[mk] = {};
    const agg = byMonth[mk];
    for (const sku of w.skus) {
      if (sku.oosPct === null || sku.storeCount === null) continue;
      if (!agg[sku.desc]) agg[sku.desc] = { weighted: 0, weight: 0, maxPct: 0, weeks: 0 };
      const e = agg[sku.desc];
      e.weighted += sku.oosPct * sku.storeCount;
      e.weight += sku.storeCount;
      e.maxPct = Math.max(e.maxPct, sku.oosPct);
      e.weeks += 1;
    }
  }
  const out = {};
  for (const mk of Object.keys(byMonth)) {
    const agg = byMonth[mk];
    const top = [];
    let overallWeighted = 0, overallWeight = 0;
    for (const desc of Object.keys(agg)) {
      const e = agg[desc];
      const avgPct = e.weight ? e.weighted / e.weight : 0;
      top.push({ desc, avgOosPct: round2(avgPct), maxOosPct: round2(e.maxPct), weeks: e.weeks });
      overallWeighted += e.weighted;
      overallWeight += e.weight;
    }
    top.sort((a, b) => b.avgOosPct - a.avgOosPct);
    out[mk] = {
      overallAvgOosPct: overallWeight ? round2(overallWeighted / overallWeight) : 0,
      trackedSkuCount: top.length,
      topOosSkus: top,
    };
  }
  return out;
}

// 주차별 재고(EOH)-판매 다이버전스 트렌드용 스냅샷.
// 판매(Sell-through Units)는 평평한데 재고(EOH Units)나 DC 발주(DC On Order Units)가
// 늘어나는 구간을 잡아내기 위한 시계열. 월간 롤업이 아니라 주 단위를 그대로 보존.
export function buildInventoryWeekly(invWeeks) {
  const byWeek = {};
  const sortedWeeks = [...invWeeks].sort((a, b) => (a.weekEndDate < b.weekEndDate ? -1 : 1));
  for (const w of sortedWeeks) {
    byWeek[w.weekEndDate] = {
      totalSalesUnits: w.totalUnits,
      eohUnits: w.eohUnits ?? null,
      eohUsd: w.invEohUsd,
      wos: w.invWos,
      dcEohUnits: w.dcEohUnits ?? null,
      dcOnOrderUnits: w.dcOnOrderUnits ?? null,
      storeEohUnits: w.storeEohUnits ?? null,
    };
  }

  const weekKeys = Object.keys(byWeek).sort();
  let divergence = null;
  const N = 6;
  if (weekKeys.length >= N) {
    const recentKeys = weekKeys.slice(-N);
    const recent = recentKeys.map((k) => byWeek[k]);
    const half = Math.floor(N / 2);
    const firstHalf = recent.slice(0, half);
    const secondHalf = recent.slice(half);

    const avg = (vals) => {
      const filtered = vals.filter((v) => v !== null && v !== undefined);
      return filtered.length ? filtered.reduce((s, v) => s + v, 0) / filtered.length : null;
    };
    const pctChange = (a, b) => {
      if (a === null || b === null || a === 0) return null;
      return Math.round(((b - a) / a) * 1000) / 10;
    };

    const salesFirst = avg(firstHalf.map((r) => r.totalSalesUnits));
    const salesSecond = avg(secondHalf.map((r) => r.totalSalesUnits));
    const eohFirst = avg(firstHalf.map((r) => r.eohUnits));
    const eohSecond = avg(secondHalf.map((r) => r.eohUnits));
    const dcFirst = avg(firstHalf.map((r) => r.dcOnOrderUnits));
    const dcSecond = avg(secondHalf.map((r) => r.dcOnOrderUnits));

    const salesChg = pctChange(salesFirst, salesSecond);
    const eohChg = pctChange(eohFirst, eohSecond);
    const dcChg = pctChange(dcFirst, dcSecond);

    let flag = false;
    const reasons = [];
    if (salesChg !== null && eohChg !== null && Math.abs(salesChg) <= 10 && eohChg >= 10) {
      flag = true;
      reasons.push(`판매는 ${salesChg > 0 ? '+' : ''}${salesChg}%로 평평한데 재고(EOH)는 ${eohChg > 0 ? '+' : ''}${eohChg}% 증가`);
    }
    if (salesChg !== null && dcChg !== null && Math.abs(salesChg) <= 10 && dcChg >= 15) {
      flag = true;
      reasons.push(`판매는 ${salesChg > 0 ? '+' : ''}${salesChg}%로 평평한데 DC 발주는 ${dcChg > 0 ? '+' : ''}${dcChg}% 증가`);
    }

    divergence = {
      windowWeeks: N,
      fromWeek: recentKeys[0], toWeek: recentKeys[recentKeys.length - 1],
      salesUnitsChangePct: salesChg, eohUnitsChangePct: eohChg, dcOnOrderChangePct: dcChg,
      flag, reasons,
    };
  }

  return { byWeek, divergence };
}

export function buildFinancials(invWeeks) {
  const byMonth = {};
  for (const w of invWeeks) {
    const mk = monthKey(w.weekEndDate);
    if (!byMonth[mk]) byMonth[mk] = { marginNum: 0, marginDen: 0, lostSales: 0, lastWeek: null };
    const agg = byMonth[mk];
    for (const sku of w.skus) {
      const { price, cost, sales } = sku;
      if (price && cost !== null && sales) {
        const marginPct = (price - cost) / price;
        agg.marginNum += marginPct * sales;
        agg.marginDen += sales;
      }
      if (sku.oosPct !== null && sales) {
        const frac = sku.oosPct / 100;
        if (frac < 1) agg.lostSales += (sales * frac) / (1 - frac);
      }
    }
    if (!agg.lastWeek || w.weekEndDate > agg.lastWeek.weekEndDate) agg.lastWeek = w;
  }
  const out = {};
  for (const mk of Object.keys(byMonth)) {
    const agg = byMonth[mk];
    const lw = agg.lastWeek;
    out[mk] = {
      marginPct: agg.marginDen ? round2((agg.marginNum / agg.marginDen) * 100) : null,
      invEohUsd: lw.invEohUsd,
      invWos: lw.invWos,
      invAsOfWeek: lw.weekEndDate,
      lostSalesUsd: round2(agg.lostSales),
    };
  }
  return out;
}

// ---------- 스냅샷(data.json) 전체 재계산 ----------
// existingSnapshot: 현재 data.json 파싱 결과 (target/marketing/kpiMonthly 등은 그대로 보존)
// storeFiles / invFiles: [{ filename, buffer }] 배열 (raw-data 폴더 내 전체 원본 파일)
export function recomputeSnapshot(existingSnapshot, storeFiles, invFiles, updatedByEmail) {
  const snapshot = JSON.parse(JSON.stringify(existingSnapshot));

  const storeWeeks = storeFiles.map((f) => parseStoreSales(f.buffer, f.filename));
  const invWeeks = invFiles.map((f) => parseSalesInvPerfWeek(f.buffer, f.filename));

  const weekly = snapshot.weekly || {};
  for (const w of storeWeeks) {
    const existing = weekly[w.weekEndDate];
    if (existing && existing.source === 'manual_table') continue;
    weekly[w.weekEndDate] = {
      weekEndDate: w.weekEndDate, fiscalYear: w.fiscalYear, fiscalWeekNum: w.fiscalWeekNum,
      totalSales: w.totalSales, totalUnits: w.totalUnits,
      comSales: w.comSales, comUnits: w.comUnits,
      bmSales: w.bmSales, bmUnits: w.bmUnits,
      storeCount: w.storeCount, source: 'file',
      uploadedAt: new Date().toISOString(),
    };
  }
  snapshot.weekly = weekly;

  if (storeWeeks.length) {
    snapshot.storeRanking = {
      byMonth: buildStoreRanking(storeWeeks),
      uploadedAt: new Date().toISOString(),
      note: 'Store-Sales 원본 xlsx 파일이 업로드된 주차만 집계됨. 일부 월은 부분 데이터일 수 있음.',
    };
    snapshot.storeWeekly = {
      byWeek: buildStoreWeekly(storeWeeks),
      note: 'Store-Sales 원본 리포트 기준 주간 매장별 매출 순위 (전주 대비 순위 변동 비교용)',
      uploadedAt: new Date().toISOString(),
    };
  }

  if (invWeeks.length) {
    snapshot.skuSales = {
      byMonth: buildSkuSales(invWeeks),
      note: 'Sales_Inv_Perf 원본 리포트의 Last Closed Week 시트 기준 SKU별 매출 합산',
      uploadedAt: new Date().toISOString(),
    };
    snapshot.skuWeekly = {
      byWeek: buildSkuWeekly(invWeeks),
      note: 'Sales_Inv_Perf 원본 리포트의 Last Closed Week 시트 기준 주간 SKU별 매출 (전주 대비 비교용)',
      uploadedAt: new Date().toISOString(),
    };
    snapshot.skuOos = {
      byMonth: buildSkuOos(invWeeks),
      note: 'Sales_Inv_Perf 원본 리포트의 Store Out of Stock % 컬럼 기준 (판매 상위 8~10개 SKU만 추적됨, storeCount 가중평균)',
      uploadedAt: new Date().toISOString(),
    };
    snapshot.financials = {
      byMonth: buildFinancials(invWeeks),
      note: 'ULTA 소매 마진은 (Retail Price - Purch Cost)/Retail Price 기준으로, ULTA 채널 마진 근사치입니다 (Celimax 자체 원가/마진과는 다름). 재고 자산가치/WOS는 해당 월 마지막 업로드 주차 기준 시점 스냅샷입니다. 품절 손실 매출은 OOS%를 이용한 근사 추정치입니다.',
      uploadedAt: new Date().toISOString(),
    };
    snapshot.inventoryWeekly = {
      ...buildInventoryWeekly(invWeeks),
      note: 'Sales_Inv_Perf 원본 리포트 기준 주간 판매(Sell-through Units) vs 재고(Total EOH Units) vs DC 발주(DC On Order Units) 추이. 판매는 평평한데 재고/DC발주만 느는 구간을 잡아내기 위한 다이버전스 트렌드입니다.',
      uploadedAt: new Date().toISOString(),
    };

    const latestInv = invFiles[invFiles.length - 1];
    const qtdYtd = parseQtdYtd(latestInv.buffer);
    const latestWeekEnd = parseFileDate(latestInv.filename);
    snapshot.qtdYtd = {
      qtd: qtdYtd.qtd, ytd: qtdYtd.ytd,
      asOfDate: latestWeekEnd,
      note: 'ULTA Sales_Inv_Perf 리포트의 Quarter to Date / Year to Date 시트에서 그대로 가져온 누적 실적입니다.',
    };
  }

  snapshot.snapshotUpdatedAt = new Date().toISOString();
  snapshot.updatedBy = updatedByEmail || snapshot.updatedBy || '-';

  return snapshot;
}

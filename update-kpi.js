// /api/update-kpi - Sell-in(내부 출고 실적) / COP(목표) 숫자를 수기로 입력해 data.json에 반영
// 담당자(ADMIN_EMAILS)만 사용 가능. 원본 파일 파싱과 무관하게 kpiMonthly / target 섹션만 갱신합니다.
import { requireAdmin } from './_lib/session.js';
import { getFile, putFile } from './_lib/github.js';

const DATA_PATH = 'data.json';

function monthKeyToKoreanLabel(key) {
  const [y, m] = key.split('-');
  return `${y.slice(2)}년 ${parseInt(m, 10)}월`;
}

function parseKoreanMonthLabel(label) {
  const m = String(label).match(/(\d{2,4})\s*년\s*(\d{1,2})\s*월/);
  if (!m) return null;
  let y = m[1];
  if (y.length === 2) y = '20' + y;
  const mo = String(parseInt(m[2], 10)).padStart(2, '0');
  return `${y}-${mo}`;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ message: 'POST 요청만 허용됩니다.' });
    return;
  }

  const email = await requireAdmin(req, res);
  if (!email) return;

  try {
    const { monthKey, usdSales, krwSales, copTargetKrw } = req.body || {};
    if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) {
      res.status(400).json({ message: 'monthKey는 "YYYY-MM" 형식이어야 합니다. (예: 2026-06)' });
      return;
    }
    const hasSellIn = usdSales !== undefined && usdSales !== null && usdSales !== '';
    const hasCop = copTargetKrw !== undefined && copTargetKrw !== null && copTargetKrw !== '';
    if (!hasSellIn && !hasCop) {
      res.status(400).json({ message: 'Sell-in(USD) 또는 COP 목표(KRW) 중 하나 이상 입력해야 합니다.' });
      return;
    }
    if (hasSellIn && (krwSales === undefined || krwSales === null || krwSales === '')) {
      res.status(400).json({ message: 'Sell-in을 입력하려면 USD 금액과 KRW 금액을 함께 입력해야 합니다 (COP 달성률 계산에 KRW 금액이 필요합니다).' });
      return;
    }

    const existingDataFile = await getFile(DATA_PATH);
    if (!existingDataFile) {
      res.status(500).json({ message: 'GitHub 저장소에서 기존 data.json을 찾을 수 없습니다.' });
      return;
    }
    const snapshot = JSON.parse(existingDataFile.text);

    if (hasSellIn) {
      const label = monthKeyToKoreanLabel(monthKey);
      snapshot.kpiMonthly = Array.isArray(snapshot.kpiMonthly) ? snapshot.kpiMonthly : [];
      let entry = snapshot.kpiMonthly.find((m) => parseKoreanMonthLabel(m.month) === monthKey);
      if (!entry) {
        entry = { month: label, qty: null, usdSales: null, krwSales: null };
        snapshot.kpiMonthly.push(entry);
      }
      entry.usdSales = Number(usdSales);
      entry.krwSales = Number(krwSales);
    }

    if (hasCop) {
      snapshot.target = snapshot.target && typeof snapshot.target === 'object' ? snapshot.target : { byMonth: {} };
      snapshot.target.byMonth = snapshot.target.byMonth || {};
      snapshot.target.byMonth[monthKey] = Number(copTargetKrw);
    }

    snapshot.snapshotUpdatedAt = new Date().toISOString();
    snapshot.updatedBy = email;

    await putFile(
      DATA_PATH,
      JSON.stringify(snapshot),
      `data: ${monthKey} Sell-in/COP 수동 입력 (by ${email})`,
      existingDataFile.sha
    );

    res.status(200).json({ ok: true, message: `${monthKeyToKoreanLabel(monthKey)} 데이터가 저장되었습니다. 잠시 후 대시보드에 반영됩니다.` });
  } catch (err) {
    res.status(500).json({ message: `저장 중 오류: ${err.message}` });
  }
}

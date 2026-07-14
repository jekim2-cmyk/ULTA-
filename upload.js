// /api/upload - ULTA 원본 리포트(xlsx) 업로드 -> GitHub raw-data/ 폴더에 저장
// -> raw-data/ 폴더 전체를 다시 읽어 data.json을 재계산 -> GitHub에 커밋 (Vercel 자동 재배포)
// 담당자(ADMIN_EMAILS)만 사용 가능.
import { requireAdmin } from './_lib/session.js';
import { getFile, listDir, putFile } from './_lib/github.js';
import { recomputeSnapshot } from './_lib/pipeline.js';

export const config = {
  api: { bodyParser: { sizeLimit: '8mb' } },
};

const RAW_DIR = 'raw-data';
const DATA_PATH = 'data.json';

function isStoreSalesFilename(name) {
  return /^Store-Sales_.*-\d{4}-\d{2}-\d{2}\.xlsx$/i.test(name);
}
function isSalesInvPerfFilename(name) {
  return /^Sales_Inv_Perf.*-\d{4}-\d{2}-\d{2}\.xlsx$/i.test(name);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ message: 'POST 요청만 허용됩니다.' });
    return;
  }

  const email = await requireAdmin(req, res);
  if (!email) return; // requireAdmin이 이미 응답을 보냄

  try {
    const { type, filename, contentBase64 } = req.body || {};
    if (!type || !filename || !contentBase64) {
      res.status(400).json({ message: 'type, filename, contentBase64 값이 모두 필요합니다.' });
      return;
    }
    if (type !== 'store-sales' && type !== 'sales-inv-perf') {
      res.status(400).json({ message: 'type은 store-sales 또는 sales-inv-perf 여야 합니다.' });
      return;
    }
    if (type === 'store-sales' && !isStoreSalesFilename(filename)) {
      res.status(400).json({ message: `파일명이 예상 형식과 다릅니다. "Store-Sales_...-YYYY-MM-DD.xlsx" 형식이어야 합니다. (받은 파일명: ${filename})` });
      return;
    }
    if (type === 'sales-inv-perf' && !isSalesInvPerfFilename(filename)) {
      res.status(400).json({ message: `파일명이 예상 형식과 다릅니다. "Sales_Inv_Perf...-YYYY-MM-DD.xlsx" 형식이어야 합니다. (받은 파일명: ${filename})` });
      return;
    }

    const fileBuffer = Buffer.from(contentBase64, 'base64');
    if (fileBuffer.length === 0) {
      res.status(400).json({ message: '파일 내용이 비어 있습니다.' });
      return;
    }

    // 1) 원본 파일을 raw-data/ 폴더에 저장 (이미 있으면 덮어쓰기)
    const rawPath = `${RAW_DIR}/${filename}`;
    const existingRaw = await getFile(rawPath);
    await putFile(
      rawPath,
      fileBuffer,
      `data: upload ${filename} (by ${email})`,
      existingRaw ? existingRaw.sha : undefined
    );

    // 2) raw-data/ 폴더 전체를 다시 읽어와 파싱 (Store-Sales + Sales_Inv_Perf 둘 다)
    const entries = await listDir(RAW_DIR);
    const storeEntries = entries.filter((e) => e.type === 'file' && isStoreSalesFilename(e.name)).sort((a, b) => a.name.localeCompare(b.name));
    const invEntries = entries.filter((e) => e.type === 'file' && isSalesInvPerfFilename(e.name)).sort((a, b) => a.name.localeCompare(b.name));

    const storeFiles = [];
    for (const e of storeEntries) {
      const f = await getFile(`${RAW_DIR}/${e.name}`);
      if (f) storeFiles.push({ filename: e.name, buffer: f.buffer });
    }
    const invFiles = [];
    for (const e of invEntries) {
      const f = await getFile(`${RAW_DIR}/${e.name}`);
      if (f) invFiles.push({ filename: e.name, buffer: f.buffer });
    }

    // 3) 기존 data.json을 불러와 target/marketing/kpiMonthly 등은 보존한 채 재계산
    const existingDataFile = await getFile(DATA_PATH);
    if (!existingDataFile) {
      res.status(500).json({ message: 'GitHub 저장소에서 기존 data.json을 찾을 수 없습니다.' });
      return;
    }
    const existingSnapshot = JSON.parse(existingDataFile.text);
    const newSnapshot = recomputeSnapshot(existingSnapshot, storeFiles, invFiles, email);

    // 4) data.json 커밋
    await putFile(
      DATA_PATH,
      JSON.stringify(newSnapshot),
      `data: ${filename} 업로드 반영 (by ${email})`,
      existingDataFile.sha
    );

    res.status(200).json({
      ok: true,
      message: `${filename} 업로드 완료. data.json이 갱신되어 잠시 후 대시보드에 반영됩니다.`,
      storeFileCount: storeFiles.length,
      invFileCount: invFiles.length,
    });
  } catch (err) {
    res.status(500).json({ message: `업로드 처리 중 오류: ${err.message}` });
  }
}

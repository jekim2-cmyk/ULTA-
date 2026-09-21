// /api/recompute - 새 파일 업로드 없이 raw-data/ 폴더 전체를 "지금 GitHub에 있는 최신 pipeline.js 로직"으로
// 다시 계산해서 data.json을 갱신합니다. pipeline.js/pipeline.py 코드를 업데이트한 직후,
// data.json을 손으로 만들어 올리는 대신 이 버튼 하나로 안전하게 전체 재계산을 반영하기 위한 용도입니다.
// 담당자(ADMIN_EMAILS)만 사용 가능.
import { requireAdmin } from './_lib/session.js';
import { getFile, listDir } from './_lib/github.js';
import { putFile } from './_lib/github.js';
import { recomputeSnapshot } from './_lib/pipeline.js';

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
    // 1) raw-data/ 폴더 전체를 다시 읽어와 파싱 (Store-Sales + Sales_Inv_Perf 둘 다)
    const entries = await listDir(RAW_DIR);
    const storeEntries = entries.filter((e) => e.type === 'file' && isStoreSalesFilename(e.name)).sort((a, b) => a.name.localeCompare(b.name));
    const invEntries = entries.filter((e) => e.type === 'file' && isSalesInvPerfFilename(e.name)).sort((a, b) => a.name.localeCompare(b.name));

    if (!storeEntries.length && !invEntries.length) {
      res.status(400).json({ message: 'raw-data 폴더에 원본 파일이 없습니다. 먼저 파일을 업로드해주세요.' });
      return;
    }

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

    // 2) 기존 data.json을 불러와 target/marketing/kpiMonthly 등은 보존한 채 재계산
    const existingDataFile = await getFile(DATA_PATH);
    if (!existingDataFile) {
      res.status(500).json({ message: 'GitHub 저장소에서 기존 data.json을 찾을 수 없습니다.' });
      return;
    }
    const existingSnapshot = JSON.parse(existingDataFile.text);
    const newSnapshot = recomputeSnapshot(existingSnapshot, storeFiles, invFiles, email);

    // 3) data.json 커밋
    await putFile(
      DATA_PATH,
      JSON.stringify(newSnapshot),
      `data: 전체 재계산 (raw-data ${storeFiles.length + invFiles.length}개 파일, by ${email})`,
      existingDataFile.sha
    );

    res.status(200).json({
      ok: true,
      message: `재계산 완료 (Store-Sales ${storeFiles.length}개, Sales_Inv_Perf ${invFiles.length}개 파일 기준). 잠시 후 대시보드에 반영됩니다.`,
      storeFileCount: storeFiles.length,
      invFileCount: invFiles.length,
    });
  } catch (err) {
    res.status(500).json({ message: `재계산 처리 중 오류: ${err.message}` });
  }
}

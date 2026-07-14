// GitHub Contents API 헬퍼 (파일 읽기/쓰기 -> 커밋 -> Vercel 자동 재배포)
// 환경변수 필요: GITHUB_TOKEN (repo 쓰기 권한 PAT), GITHUB_REPO ("owner/repo" 형식)
// 선택: GITHUB_BRANCH (기본값 main)

const API_BASE = 'https://api.github.com';

function getRepoConfig() {
  const repo = process.env.GITHUB_REPO;
  const branch = process.env.GITHUB_BRANCH || 'main';
  const token = process.env.GITHUB_TOKEN;
  if (!repo || !token) {
    throw new Error('서버에 GITHUB_REPO / GITHUB_TOKEN 환경변수가 설정되지 않았습니다. Vercel 프로젝트 설정에서 등록해주세요.');
  }
  return { repo, branch, token };
}

function encodePath(p) {
  return p.split('/').map(encodeURIComponent).join('/');
}

async function ghFetch(pathAndQuery, options = {}) {
  const { token } = getRepoConfig();
  const res = await fetch(`${API_BASE}${pathAndQuery}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      ...(options.headers || {}),
    },
  });
  return res;
}

// 단일 파일 조회. 없으면 null 반환.
export async function getFile(filePath) {
  const { repo, branch } = getRepoConfig();
  const res = await ghFetch(`/repos/${repo}/contents/${encodePath(filePath)}?ref=${encodeURIComponent(branch)}`);
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub 파일 조회 실패 (${res.status}): ${await res.text()}`);
  const data = await res.json();
  const buffer = Buffer.from(data.content, 'base64');
  return { sha: data.sha, buffer, text: buffer.toString('utf-8') };
}

// 폴더 목록 조회. 없으면 빈 배열.
export async function listDir(dirPath) {
  const { repo, branch } = getRepoConfig();
  const res = await ghFetch(`/repos/${repo}/contents/${encodePath(dirPath)}?ref=${encodeURIComponent(branch)}`);
  if (res.status === 404) return [];
  if (!res.ok) throw new Error(`GitHub 폴더 조회 실패 (${res.status}): ${await res.text()}`);
  const data = await res.json();
  return Array.isArray(data) ? data : [];
}

// 파일 생성/수정 (커밋 1건 발생 -> Vercel Git 연동이 자동 재배포)
export async function putFile(filePath, contentBufferOrString, message, sha) {
  const { repo, branch } = getRepoConfig();
  const contentBase64 = Buffer.isBuffer(contentBufferOrString)
    ? contentBufferOrString.toString('base64')
    : Buffer.from(contentBufferOrString, 'utf-8').toString('base64');
  const body = { message, content: contentBase64, branch };
  if (sha) body.sha = sha;
  const res = await ghFetch(`/repos/${repo}/contents/${encodePath(filePath)}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`GitHub 파일 저장 실패 (${res.status}): ${await res.text()}`);
  return res.json();
}

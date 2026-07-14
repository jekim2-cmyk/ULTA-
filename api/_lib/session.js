// 공통 세션/권한 헬퍼 (Node.js 런타임 API 함수들이 공유)
// middleware.js / api/verify.js 와 동일한 쿠키 형식(ulta_session)을 검증합니다.

const COOKIE_NAME = 'ulta_session';

function toBase64Url(bytes) {
  let binary = '';
  const arr = new Uint8Array(bytes);
  for (let i = 0; i < arr.length; i++) binary += String.fromCharCode(arr[i]);
  return Buffer.from(binary, 'binary').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function fromBase64Url(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  return Buffer.from(str, 'base64');
}

async function hmacSign(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(message));
  return toBase64Url(sig);
}

async function verifySessionToken(token, secret) {
  if (!token) return null;
  const parts = token.split('.');
  if (parts.length !== 2) return null;
  const [payloadB64, sigB64] = parts;
  const expectedSig = await hmacSign(secret, payloadB64);
  if (expectedSig !== sigB64) return null;
  try {
    const payloadBytes = fromBase64Url(payloadB64);
    const payload = JSON.parse(payloadBytes.toString('utf-8'));
    if (!payload.exp || Date.now() > payload.exp) return null;
    return payload;
  } catch (e) {
    return null;
  }
}

export function parseCookies(cookieHeader) {
  const out = {};
  (cookieHeader || '').split(';').forEach((pair) => {
    const idx = pair.indexOf('=');
    if (idx === -1) return;
    const k = pair.slice(0, idx).trim();
    const v = pair.slice(idx + 1).trim();
    if (k) out[k] = decodeURIComponent(v);
  });
  return out;
}

// req: Node.js 스타일 요청 객체 (req.headers.cookie 사용)
export async function getSessionEmail(req) {
  const secret = process.env.SESSION_SECRET;
  if (!secret) return null;
  const cookieHeader = (req.headers && (req.headers.cookie || req.headers.Cookie)) || '';
  const cookies = parseCookies(cookieHeader);
  const token = cookies[COOKIE_NAME];
  const session = await verifySessionToken(token, secret);
  return session ? session.email : null;
}

export function isAdminEmail(email) {
  if (!email) return false;
  const list = (process.env.ADMIN_EMAILS || '')
    .split(',')
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
  return list.includes(email.toLowerCase());
}

// 관리자 여부까지 함께 확인하고, 아니면 표준 오류 응답을 내려주는 헬퍼
export async function requireAdmin(req, res) {
  const email = await getSessionEmail(req);
  if (!email) {
    res.status(401).json({ message: '로그인이 필요합니다.' });
    return null;
  }
  if (!isAdminEmail(email)) {
    res.status(403).json({ message: `이 기능은 담당자만 사용할 수 있습니다 (${email} 계정은 권한이 없습니다).` });
    return null;
  }
  return email;
}

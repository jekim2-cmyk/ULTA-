// Vercel Edge Function: /api/verify
// 프론트(login.html)에서 받은 Google ID 토큰(credential)을 검증하고,
// 회사 도메인(@celimax.co.kr, Google Workspace) 계정일 때만 로그인 세션 쿠키를 발급합니다.

export const config = { runtime: 'edge' };

const COOKIE_NAME = 'ulta_session';
const SESSION_DAYS = 30;

function toBase64Url(bytes) {
  let binary = '';
  const arr = new Uint8Array(bytes);
  for (let i = 0; i < arr.length; i++) binary += String.fromCharCode(arr[i]);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function hmacSign(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(message));
  return toBase64Url(sig);
}

async function createSessionToken(secret, email) {
  const payload = {
    email,
    exp: Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000,
  };
  const payloadBytes = new TextEncoder().encode(JSON.stringify(payload));
  const payloadB64 = toBase64Url(payloadBytes);
  const sigB64 = await hmacSign(secret, payloadB64);
  return `${payloadB64}.${sigB64}`;
}

export default async function handler(request) {
  if (request.method !== 'POST') {
    return new Response(JSON.stringify({ message: 'Method not allowed' }), { status: 405 });
  }

  const clientId = process.env.GOOGLE_CLIENT_ID;
  const allowedDomain = process.env.ALLOWED_DOMAIN || 'celimax.co.kr';
  const secret = process.env.SESSION_SECRET;

  if (!clientId || !secret) {
    return new Response(JSON.stringify({ message: '서버 환경변수(GOOGLE_CLIENT_ID / SESSION_SECRET)가 설정되지 않았습니다.' }), { status: 500 });
  }

  let body;
  try {
    body = await request.json();
  } catch (e) {
    return new Response(JSON.stringify({ message: '잘못된 요청입니다.' }), { status: 400 });
  }

  const credential = body && body.credential;
  if (!credential) {
    return new Response(JSON.stringify({ message: '로그인 토큰이 없습니다.' }), { status: 400 });
  }

  // Google 토큰 검증 (서명/만료는 Google 서버가 확인해줌)
  let tokenInfo;
  try {
    const res = await fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential));
    if (!res.ok) {
      return new Response(JSON.stringify({ message: '구글 로그인 확인에 실패했습니다. 다시 시도해주세요.' }), { status: 401 });
    }
    tokenInfo = await res.json();
  } catch (e) {
    return new Response(JSON.stringify({ message: '구글 서버 확인 중 오류가 발생했습니다.' }), { status: 502 });
  }

  if (tokenInfo.aud !== clientId) {
    return new Response(JSON.stringify({ message: '허용되지 않은 클라이언트입니다.' }), { status: 401 });
  }
  if (tokenInfo.email_verified !== 'true' && tokenInfo.email_verified !== true) {
    return new Response(JSON.stringify({ message: '이메일이 확인되지 않은 계정입니다.' }), { status: 401 });
  }
  const hd = tokenInfo.hd;
  const email = tokenInfo.email || '';
  const isAllowedDomain = hd === allowedDomain || email.toLowerCase().endsWith('@' + allowedDomain.toLowerCase());
  if (!isAllowedDomain) {
    return new Response(JSON.stringify({ message: `회사 계정(@${allowedDomain})으로만 로그인할 수 있습니다.` }), { status: 403 });
  }

  const token = await createSessionToken(secret, email);
  const cookie = `${COOKIE_NAME}=${encodeURIComponent(token)}; Path=/; Max-Age=${SESSION_DAYS * 24 * 60 * 60}; HttpOnly; Secure; SameSite=Lax`;

  return new Response(JSON.stringify({ ok: true, email }), {
    status: 200,
    headers: {
      'Content-Type': 'application/json',
      'Set-Cookie': cookie,
    },
  });
}

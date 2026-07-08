// Vercel Routing Middleware (프레임워크 없는 정적 사이트용)
// 모든 요청에 대해 로그인 세션 쿠키를 검사합니다.
// 세션이 없거나 만료/위조되었으면 /login.html 로 보냅니다.
// login.html, /api/* 는 검사 대상에서 제외합니다 (matcher 설정 참고).

import { next } from '@vercel/functions';

export const config = {
  matcher: ['/((?!api/|login\\.html).*)'],
};

const COOKIE_NAME = 'ulta_session';

function toBase64Url(bytes) {
  let binary = '';
  const arr = new Uint8Array(bytes);
  for (let i = 0; i < arr.length; i++) binary += String.fromCharCode(arr[i]);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function fromBase64Url(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  const binary = atob(str);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
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
    const payload = JSON.parse(new TextDecoder().decode(payloadBytes));
    if (!payload.exp || Date.now() > payload.exp) return null;
    return payload;
  } catch (e) {
    return null;
  }
}

export default async function middleware(request) {
  const secret = process.env.SESSION_SECRET;
  const cookieHeader = request.headers.get('cookie') || '';
  const match = cookieHeader.match(new RegExp('(?:^|; )' + COOKIE_NAME + '=([^;]*)'));
  const token = match ? decodeURIComponent(match[1]) : null;

  const session = secret ? await verifySessionToken(token, secret) : null;

  if (session && session.email) {
    return next(); // 통과
  }

  const url = new URL(request.url);
  const loginUrl = new URL('/login.html', url.origin);
  loginUrl.searchParams.set('next', url.pathname);
  return Response.redirect(loginUrl, 302);
}

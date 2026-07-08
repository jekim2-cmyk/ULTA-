// Vercel Edge Function: /api/config
// login.html이 구글 클라이언트 ID를 코드에 하드코딩하지 않고 받아갈 수 있도록 함.

export const config = { runtime: 'edge' };

export default async function handler(request) {
  return new Response(JSON.stringify({
    clientId: process.env.GOOGLE_CLIENT_ID || '',
    allowedDomain: process.env.ALLOWED_DOMAIN || 'celimax.co.kr',
  }), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

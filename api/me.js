// /api/me - 현재 로그인한 사용자 이메일 및 관리자(데이터 업로드 권한) 여부를 프론트에 알려줌
import { getSessionEmail, isAdminEmail } from './_lib/session.js';

export default async function handler(req, res) {
  const email = await getSessionEmail(req);
  if (!email) {
    res.status(200).json({ loggedIn: false, email: null, isAdmin: false });
    return;
  }
  res.status(200).json({ loggedIn: true, email, isAdmin: isAdminEmail(email) });
}

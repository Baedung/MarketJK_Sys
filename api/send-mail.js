import nodemailer from 'nodemailer';
import { kv } from '@vercel/kv';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const { to, subject, body, fileName, base64 } = req.body || {};
  if (!to || !subject || !fileName || !base64) {
    res.status(400).json({ error: '필수 값(받는 사람/제목/파일)이 누락되었습니다.' });
    return;
  }

  // 1순위: '설정' 화면에서 Vercel KV에 저장한 계정 정보
  // 2순위: Vercel 프로젝트 환경변수 (NAVER_EMAIL / NAVER_APP_PASSWORD)
  let user = process.env.NAVER_EMAIL;
  let pass = process.env.NAVER_APP_PASSWORD;
  try {
    const cfg = await kv.get('mail-config');
    if (cfg?.email) user = cfg.email;
    if (cfg?.appPassword) pass = cfg.appPassword;
  } catch (err) {
    // KV 미연결 시 환경변수만 사용
  }

  if (!user || !pass) {
    res.status(500).json({
      error: '네이버 메일 계정이 설정되어 있지 않습니다. "설정" 탭에서 이메일·앱 비밀번호를 저장하거나, Vercel 환경변수(NAVER_EMAIL / NAVER_APP_PASSWORD)를 추가해주세요.',
    });
    return;
  }

  try {
    const transporter = nodemailer.createTransport({
      host: 'smtp.naver.com',
      port: 587,
      secure: false, // STARTTLS
      auth: { user, pass },
    });

    await transporter.sendMail({
      from: `"발주 담당" <${user}>`,
      to,
      subject,
      text: body,
      attachments: [
        {
          filename: fileName,
          content: Buffer.from(base64, 'base64'),
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
      ],
    });

    res.status(200).json({ ok: true });
  } catch (err) {
    console.error('send-mail error:', err);
    res.status(500).json({ error: err.message || '메일 발송 중 오류가 발생했습니다.' });
  }
}

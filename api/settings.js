import { Redis } from '@upstash/redis';

const redis = Redis.fromEnv();

const CONFIG_KEY = 'mail-config';

export default async function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const cfg = await redis.get(CONFIG_KEY);
      res.status(200).json({ email: cfg?.email || '', hasPassword: !!cfg?.appPassword });
    } catch (err) {
      // Upstash Redis가 아직 프로젝트에 연결되지 않은 경우
      res.status(200).json({ email: '', hasPassword: false, kvMissing: true });
    }
    return;
  }

  if (req.method === 'POST') {
    const { email, appPassword } = req.body || {};
    if (!email || !String(email).trim()) {
      res.status(400).json({ error: '네이버 이메일을 입력해주세요.' });
      return;
    }
    try {
      const existing = (await redis.get(CONFIG_KEY)) || {};
      const next = {
        email: String(email).trim(),
        appPassword: appPassword ? String(appPassword).trim() : existing.appPassword || '',
      };
      await redis.set(CONFIG_KEY, next);
      res.status(200).json({ email: next.email, hasPassword: !!next.appPassword });
    } catch (err) {
      console.error('settings save error:', err);
      res.status(500).json({
        error: 'Upstash Redis가 연결되지 않았습니다. 프로젝트의 Storage 탭에서 Upstash Redis를 만들어 연결한 뒤 다시 시도해주세요.',
      });
    }
    return;
  }

  res.status(405).json({ error: 'Method not allowed' });
}

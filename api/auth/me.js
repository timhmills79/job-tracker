// api/auth/me.js
// Returns the current authenticated user, or 401 if not logged in

import crypto from 'crypto';

function verify(signed, secret) {
  const [value, sig] = signed.split('.');
  if (!value || !sig) return null;
  const expected = crypto.createHmac('sha256', secret).update(value).digest('base64url');
  if (sig !== expected) return null;
  try {
    return JSON.parse(Buffer.from(value, 'base64url').toString('utf8'));
  } catch {
    return null;
  }
}

export function getSession(req) {
  const cookieHeader = req.headers.cookie || '';
  const cookies = Object.fromEntries(
    cookieHeader.split(';').map(c => {
      const [k, ...v] = c.trim().split('=');
      return [k, v.join('=')];
    })
  );
  const raw = cookies['session'];
  if (!raw) return null;
  return verify(raw, process.env.SESSION_SECRET);
}

export default function handler(req, res) {
  const session = getSession(req);
  if (!session) return res.status(401).json({ error: 'Not authenticated' });

  // Refresh check (token expiry) — client will re-trigger login if needed
  if (session.expires_at < Date.now() + 60_000) {
    return res.status(401).json({ error: 'Token expired' });
  }

  res.json({
    email: session.email,
    name: session.name,
    picture: session.picture,
  });
}

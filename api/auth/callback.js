// api/auth/callback.js
// Exchanges the Google auth code for tokens and stores them in a signed cookie

import crypto from 'crypto';

function sign(value, secret) {
  const sig = crypto.createHmac('sha256', secret).update(value).digest('base64url');
  return `${value}.${sig}`;
}

export default async function handler(req, res) {
  const { code, error } = req.query;

  if (error || !code) {
    return res.redirect('/?error=auth_denied');
  }

  try {
    // Exchange code for tokens
    const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        client_id: process.env.GOOGLE_CLIENT_ID,
        client_secret: process.env.GOOGLE_CLIENT_SECRET,
        redirect_uri: `${process.env.VITE_APP_URL}/api/auth/callback`,
        grant_type: 'authorization_code',
      }),
    });

    const tokens = await tokenRes.json();
    if (!tokens.access_token) throw new Error('No access token received');

    // Get user profile
    const profileRes = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${tokens.access_token}` },
    });
    const profile = await profileRes.json();

    // Build session payload
    const session = {
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_at: Date.now() + (tokens.expires_in || 3600) * 1000,
      email: profile.email,
      name: profile.name,
      picture: profile.picture,
    };

    const sessionStr = Buffer.from(JSON.stringify(session)).toString('base64url');
    const signed = sign(sessionStr, process.env.SESSION_SECRET);

    res.setHeader('Set-Cookie', [
      `session=${signed}; Path=/; HttpOnly; SameSite=Lax; Max-Age=604800${process.env.NODE_ENV === 'production' ? '; Secure' : ''}`,
    ]);

    res.redirect('/');
  } catch (err) {
    console.error('Auth callback error:', err);
    res.redirect('/?error=auth_failed');
  }
}

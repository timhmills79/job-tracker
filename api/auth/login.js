// api/auth/login.js
// Vercel serverless function — redirects user to Google OAuth consent screen

export default function handler(req, res) {
  const params = new URLSearchParams({
    client_id: process.env.GOOGLE_CLIENT_ID,
    redirect_uri: `${process.env.VITE_APP_URL}/api/auth/callback`,
    response_type: 'code',
    scope: [
      'https://www.googleapis.com/auth/drive.file',
      'https://www.googleapis.com/auth/spreadsheets',
      'openid',
      'email',
      'profile',
    ].join(' '),
    access_type: 'offline',
    prompt: 'consent',
  });

  res.redirect(`https://accounts.google.com/o/oauth2/v2/auth?${params}`);
}

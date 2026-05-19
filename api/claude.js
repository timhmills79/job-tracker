import { getSession } from './auth/me.js';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  const session = getSession(req);
  if (!session) return res.status(401).json({ error: 'Not authenticated' });
  try {
    const upstream = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'mcp-client-2025-04-04,interleaved-thinking-2025-05-14',
      },
      body: JSON.stringify(req.body),
    });
    const data = await upstream.json();
    console.log('Claude API status:', upstream.status, 'error:', data.error);
    res.status(upstream.status).json(data);
  } catch (err) {
    console.error('Claude proxy error:', err);
    res.status(500).json({ error: 'Upstream request failed' });
  }
}

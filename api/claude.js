// api/claude.js
// Secure proxy for Anthropic API — keeps ANTHROPIC_API_KEY server-side only.
// Accepts the same body shape as /v1/messages but strips any client-supplied
// api-key headers and injects the server key instead.

import { getSession } from './auth/me.js';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  // Require authentication
  const session = getSession(req);
  if (!session) return res.status(401).json({ error: 'Not authenticated' });

  try {
    const body = req.body;

    // Forward to Anthropic
    const upstream = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'mcp-client-2025-04-04',
      },
      body: JSON.stringify(body),
    });

    const data = await upstream.json();
    res.status(upstream.status).json(data);
  } catch (err) {
    console.error('Claude proxy error:', err);
    res.status(500).json({ error: 'Upstream request failed' });
  }
}

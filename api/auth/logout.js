// api/auth/logout.js
export default function handler(req, res) {
  res.setHeader('Set-Cookie', 'session=; Path=/; HttpOnly; Max-Age=0');
  res.redirect('/');
}

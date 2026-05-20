// api/search.js
// Free job board API aggregator — no Claude tokens used for search.
// Sources: Adzuna (free tier), RemoteOK (free), The Muse (free), USAJobs (free)

import { getSession } from './auth/me.js';

// ── Adzuna ────────────────────────────────────────────────────────────────────
async function searchAdzuna(title, location, locationType, salaryMin) {
  const appId = process.env.ADZUNA_APP_ID;
  const appKey = process.env.ADZUNA_APP_KEY;
  if (!appId || !appKey) return [];

  const country = 'us';
  const params = new URLSearchParams({
    app_id: appId,
    app_key: appKey,
    results_per_page: 6,
    what: title,
    content_type: 'application/json',
  });

  if (location) params.set('where', location);
  if (salaryMin) params.set('salary_min', salaryMin);
  if (locationType === 'REMOTE') params.set('what_and', 'remote');

  const res = await fetch(`https://api.adzuna.com/v1/api/jobs/${country}/search/1?${params}`);
  if (!res.ok) return [];
  const data = await res.json();

  return (data.results || []).map((j, i) => ({
    id: `adzuna-${j.id || i}`,
    title: j.title || '',
    company: j.company?.display_name || 'Unknown',
    location: j.location?.display_name || '',
    salary: j.salary_min ? `$${Math.round(j.salary_min).toLocaleString()}${j.salary_max ? `–$${Math.round(j.salary_max).toLocaleString()}` : '+'}` : 'Not specified',
    url: j.redirect_url || '',
    snippet: j.description ? j.description.slice(0, 200) + '...' : '',
    posted: j.created ? new Date(j.created).toLocaleDateString() : '',
    source: 'adzuna',
  }));
}

// ── RemoteOK ──────────────────────────────────────────────────────────────────
async function searchRemoteOK(title) {
  const res = await fetch('https://remoteok.com/api', {
    headers: { 'User-Agent': 'JobSearchTracker/1.0' },
  });
  if (!res.ok) return [];
  const data = await res.json();

  const titleWords = title.toLowerCase().split(' ').filter(w => w.length > 2);
  const filtered = data
    .filter(j => j.position && titleWords.some(w => j.position.toLowerCase().includes(w)))
    .slice(0, 6);

  return filtered.map((j, i) => ({
    id: `remoteok-${j.id || i}`,
    title: j.position || '',
    company: j.company || 'Unknown',
    location: 'Remote',
    salary: j.salary ? `$${j.salary}` : 'Not specified',
    url: j.url || `https://remoteok.com/remote-jobs/${j.id}`,
    snippet: j.description ? j.description.replace(/<[^>]*>/g, '').slice(0, 200) + '...' : '',
    posted: j.date ? new Date(j.date).toLocaleDateString() : '',
    source: 'remoteok',
  }));
}

// ── The Muse ──────────────────────────────────────────────────────────────────
async function searchTheMuse(title, location) {
  const params = new URLSearchParams({
    query: title,
    page: 0,
    api_key: process.env.MUSE_API_KEY || '',
  });
  if (location) params.set('location', location);

  const res = await fetch(`https://www.themuse.com/api/public/jobs?${params}`);
  if (!res.ok) return [];
  const data = await res.json();

  return (data.results || []).slice(0, 6).map((j, i) => ({
    id: `muse-${j.id || i}`,
    title: j.name || '',
    company: j.company?.name || 'Unknown',
    location: j.locations?.map(l => l.name).join(', ') || 'Not specified',
    salary: 'Not specified',
    url: j.refs?.landing_page || '',
    snippet: j.contents ? j.contents.replace(/<[^>]*>/g, '').slice(0, 200) + '...' : '',
    posted: j.publication_date ? new Date(j.publication_date).toLocaleDateString() : '',
    source: 'themuse',
  }));
}

// ── USAJobs ───────────────────────────────────────────────────────────────────
async function searchUSAJobs(title, location, locationType) {
  const email = process.env.USAJOBS_EMAIL;
  const apiKey = process.env.USAJOBS_API_KEY;
  if (!email || !apiKey) return [];

  const params = new URLSearchParams({
    Keyword: title,
    ResultsPerPage: 6,
  });
  if (location) params.set('LocationName', location);
  if (locationType === 'REMOTE') params.set('RemoteIndicator', 'true');

  const res = await fetch(`https://data.usajobs.gov/api/search?${params}`, {
    headers: {
      'Authorization-Key': apiKey,
      'User-Agent': email,
      'Host': 'data.usajobs.gov',
    },
  });
  if (!res.ok) return [];
  const data = await res.json();

  const items = data.SearchResult?.SearchResultItems || [];
  return items.map((item, i) => {
    const j = item.MatchedObjectDescriptor;
    const pay = j.PositionRemuneration?.[0];
    const salary = pay ? `$${parseInt(pay.MinimumRange).toLocaleString()}–$${parseInt(pay.MaximumRange).toLocaleString()}` : 'Not specified';
    return {
      id: `usajobs-${j.PositionID || i}`,
      title: j.PositionTitle || '',
      company: j.OrganizationName || 'Federal Government',
      location: j.PositionLocationDisplay || '',
      salary,
      url: j.PositionURI || '',
      snippet: j.QualificationSummary ? j.QualificationSummary.slice(0, 200) + '...' : '',
      posted: j.PublicationStartDate ? new Date(j.PublicationStartDate).toLocaleDateString() : '',
      source: 'usajobs',
    };
  });
}

// ── Main handler ──────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const session = getSession(req);
  if (!session) return res.status(401).json({ error: 'Not authenticated' });

  const { source, title, location, locationType, seniority, salaryMin } = req.body;

  try {
    let results = [];

    switch (source) {
      case 'adzuna':
        results = await searchAdzuna(title, location, locationType, salaryMin);
        break;
      case 'remoteok':
        results = await searchRemoteOK(title);
        break;
      case 'themuse':
        results = await searchTheMuse(title, location);
        break;
      case 'usajobs':
        results = await searchUSAJobs(title, location, locationType);
        break;
      default:
        return res.status(400).json({ error: `Unknown source: ${source}` });
    }

    res.json({ results });
  } catch (err) {
    console.error(`Search error (${source}):`, err);
    res.json({ results: [], error: err.message });
  }
}

# Job Search Tracker

Search 6 job boards, score against your resume, tailor applications, and sync everything to your own Google Sheets — shareable with anyone via a free Vercel deployment.

---

## Quickstart (about 20 minutes total)

### Step 1 — Google Cloud Setup (~10 min)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (e.g. "Job Tracker")
3. Go to **APIs & Services → Library** and enable:
   - **Google Drive API**
   - **Google Sheets API**
4. Go to **APIs & Services → OAuth consent screen**
   - User type: **External**
   - Fill in app name, support email, developer email
   - Add scopes: `.../auth/drive.file` and `.../auth/spreadsheets`
   - Add your own email as a test user (while in development)
5. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**
   - Application type: **Web application**
   - Authorized redirect URIs:
     - `http://localhost:3001/api/auth/callback` (local dev)
     - `https://YOUR-APP.vercel.app/api/auth/callback` (production — fill in after deploy)
   - Copy the **Client ID** and **Client Secret**

### Step 2 — Anthropic API Key (~2 min)

1. Go to [console.anthropic.com/settings/keys](https://console.anthropic.com/settings/keys)
2. Create a new API key
3. Copy it — you'll need it in Step 4

### Step 3 — Push to GitHub (~5 min)

```bash
cd job-tracker
git init
git add .
git commit -m "Initial commit"
```

Create a new repo at [github.com/new](https://github.com/new), then:

```bash
git remote add origin https://github.com/YOUR_USERNAME/job-tracker.git
git branch -M main
git push -u origin main
```

### Step 4 — Deploy to Vercel (~3 min)

1. Go to [vercel.com](https://vercel.com) and sign in with GitHub
2. Click **Add New Project** → select your `job-tracker` repo
3. Under **Environment Variables**, add:

| Key | Value |
|-----|-------|
| `ANTHROPIC_API_KEY` | your Anthropic key |
| `GOOGLE_CLIENT_ID` | from Step 1 |
| `GOOGLE_CLIENT_SECRET` | from Step 1 |
| `SESSION_SECRET` | any 32+ char random string |
| `VITE_APP_URL` | `https://your-app.vercel.app` (check your Vercel URL first) |

4. Click **Deploy**

### Step 5 — Add redirect URI to Google (~1 min)

After deploying, go back to Google Cloud → Credentials → your OAuth client and add:
```
https://YOUR-APP.vercel.app/api/auth/callback
```

---

## Local Development

```bash
cp .env.example .env.local
# Fill in your values in .env.local

npm install
npm run dev        # frontend on :5173
# In another terminal:
node api-dev.js    # backend on :3001 (see below)
```

For local API testing, Vercel's dev CLI is the easiest:
```bash
npm install -g vercel
vercel dev         # runs both frontend and API functions together
```

---

## How It Works

Each user who visits your URL:
1. Clicks "Sign in with Google"
2. Authorizes the app to access **only their own** Drive and Sheets
3. The app creates a "Job Search Tracker" sheet in **their** Google Drive
4. All their data is completely separate from other users

**Your API keys** (Anthropic, Google OAuth client secret) stay on the server — users never see them.

---

## Sharing

Once deployed, just share your Vercel URL. Anyone can:
- Sign in with their own Google account
- Search jobs and build their own tracker
- Export to their own Google Sheet

---

## Environment Variables Reference

| Variable | Description |
|----------|-------------|
| `ANTHROPIC_API_KEY` | Powers AI search, scoring, tailoring, salary research |
| `GOOGLE_CLIENT_ID` | OAuth app client ID from Google Cloud |
| `GOOGLE_CLIENT_SECRET` | OAuth app client secret from Google Cloud |
| `SESSION_SECRET` | Random string for signing session cookies |
| `VITE_APP_URL` | Your full deployment URL, no trailing slash |

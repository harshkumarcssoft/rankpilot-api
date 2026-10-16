RankPilot API connection fix (v08)

Replace these two files in the root of the GitHub repository:
- api.py
- src/main.jsx

Then redeploy the Render API service so the updated CORS settings take effect, and redeploy the Vercel frontend so the configured API base is compiled into the site.

Vercel Production environment variable:
VITE_API_BASE_URL=https://rankpilot-api-ovfh.onrender.com/api

This patch fixes the immediate issue: the frontend previously hard-coded requests to /api/... on the Vercel domain, ignoring VITE_API_BASE_URL. The patch routes those calls to the Render API and allows the production Vercel origin in CORS.

Note: this only fixes API connection/routing. The current API is an early prototype with local SQLite/files and no authentication; do not use it for sensitive or public client workloads without hardening and persistent storage.

RANKPILOT COMPLETE CRAWLER + DASHBOARD PACKAGE v10

Upload the extracted contents of this folder to the root of your GitHub repository. Keep `src/` and `tests/` as folders. Do not upload only this README or the ZIP itself.

The package includes the React/Vite dashboard, FastAPI API, crawler, requirements files, tests and deployment documentation.

Key changes in v10:
- Dashboard crawl settings are connected to the API/crawler: max URLs, delay, robots.txt, sitemap discovery, user-agent, include/exclude patterns, JavaScript rendering option, pasted URLs and `.txt`/`.csv`/`.xlsx`/`.xlsm` URL inventory imports (up to 4 MB), RSS/Atom feeds, and previous-crawl URL import.
- Images, Links and Technical dashboard panels read the extracted page fields instead of showing placeholder counts.
- Sitemap/feed/robots and raw/rendered link-discovery outcomes are recorded in `Discovery Diagnostics`.
- `Crawl Summary` reports discovered, attempted, successfully extracted and blocked URLs, plus crawl quality.
- Cloudflare/common access challenge pages are recorded as blocked; their challenge HTML is not treated as the website's real SEO content. The crawler does not bypass bot protection.
- The dashboard shows health as N/A if no successful page extraction was confirmed.
- Workbook columns remain dynamic with no fixed 75-column limit.

Deploy the Python API separately from the Vercel frontend. On Render, build with `pip install -r requirements.txt` and start with `uvicorn api:app --host 0.0.0.0 --port $PORT`. On the Vercel frontend project, set `VITE_API_BASE_URL` to the full Render API URL ending in `/api`, then redeploy.

JavaScript rendering requires installing `requirements-rendering.txt` and Playwright Chromium in the API environment. Search Console, Analytics, backlink, CMS and server-log data remain unavailable until authorised integrations are configured. The API prototype still needs authentication and persistent storage before confidential or public production use.

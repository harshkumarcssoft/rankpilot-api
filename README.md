# Rankpilot SEO Audit Platform — GitHub upload package

**Important deployment note:** this is a React/Vite frontend plus a FastAPI/Python crawler API. It is not a single-file Streamlit app, so uploading this project to Streamlit Community Cloud as though it were the old Streamlit app will not run this dashboard. Deploy the frontend and API as separate services (for example, a static site plus a Python web service), or run it locally using the steps below.

This package has the application files at the repository root (not nested inside an extra folder) and includes a combined `requirements.txt` for platforms that expect that filename. For the local setup, use the dedicated requirement files listed below.


This build connects the React/Vite dashboard to the existing Python SEO crawler through a local FastAPI service. Projects and crawl history are saved in a local SQLite database. Crawl reports retain the `1 - All` worksheet, keep baseline fields first, and automatically include every additional collected field; there is no fixed 75-column limit.

## Requirements

- Node.js LTS (for the React dashboard)
- Python 3.10+ (for the crawler API)
- Internet access for installing dependencies and crawling public sites

## Windows setup (first time)

1. Extract this ZIP to a normal folder, for example `C:\\Users\\YourName\\Downloads\\rankpilot_seo_audit_platform`.
2. Open the folder in File Explorer.
3. Click the address bar, type `cmd`, and press Enter.
4. Install the frontend packages:

   ```bat
   npm install
   ```

5. In that same Command Prompt, install the Python API packages:

   ```bat
   py -m pip install -r requirements-api.txt
   ```

## Start it (each time you want to use the app)

Keep **two Command Prompt windows open**.

**Window 1 — Python crawler API**

Double-click `start_api.bat`, or run this in the project folder:

```bat
py -m uvicorn api:app --host 127.0.0.1 --port 8000
```

The API should be available at `http://127.0.0.1:8000/api/health`.

**Window 2 — React dashboard**

Open another Command Prompt in the same project folder and run:

```bat
npm run dev
```

Open the local URL Vite prints, normally `http://localhost:5173/`.

## Use the connected dashboard

1. Confirm the top status says **Crawler API connected**.
2. Click **New project**.
3. Enter a project name and full public URL such as `https://example.com`.
4. Select the project from the project dropdown.
5. In Settings, choose a crawl limit and delay. Start with 20–100 URLs and at least a 1-second delay while testing.
6. Click **Re-crawl**. The Python crawler will run in the background; the progress banner polls live counts, current URL, speed and ETA.
7. The report updates as saved crawl batches become available and loads the final report when the crawl completes. Use Export to download CSV/Excel.

## Data and limitations

- Project metadata and crawl history are stored in `data/rankpilot.sqlite3`; crawl workbooks and logs are stored in `data/projects/`.
- `1 - All` uses a dynamic schema: baseline columns appear first, then all additional extracted fields are appended automatically. There is no fixed 75-column ceiling.
- Crawl depth and link metrics are finalized at the end of a crawl; interim dashboard data can be partial until the crawl completes.
- This crawler does not render JavaScript and does not implement true pause/resume. Stop is supported.
- The queue count reflects URLs currently waiting in the crawler queue; the ETA is an estimate based on observed speed and the requested limit.
- This local build is for development and internal testing. The API is deliberately bound to `127.0.0.1` and has no user authentication. Do not expose it to the internet or share it with external clients as-is. Before client-facing hosting, add authentication, per-user/project authorization, crawl quotas, stronger SSRF protections, and managed report storage.
- The PDF control uses browser print-to-PDF. Some advanced dashboard panels are only as complete as the source crawler data.

## API endpoints

- `GET /api/health`
- `GET /api/projects`, `POST /api/projects`
- `GET /api/projects/{project_id}/crawls`, `POST /api/projects/{project_id}/crawls`
- `GET /api/crawls/{crawl_id}/progress`
- `GET /api/crawls/{crawl_id}/pages`
- `POST /api/crawls/{crawl_id}/stop`
- `GET /api/crawls/{crawl_id}/download`


## Metadata extraction accuracy update

The crawler now:
- checks all standard `meta[name="description"]` elements and uses the first non-empty value (an empty duplicate no longer hides a later populated tag);
- detects HTML by response content type and a conservative HTML-body sniff when servers omit/mislabel Content-Type;
- writes an `Audit Diagnostics` worksheet containing extraction evidence, whether the body was truncated, whether JavaScript rendering was checked, and raw-vs-selected title/description/H1 values;
- keeps baseline `1 - All` fields first and automatically appends new fields discovered by future crawler updates;
- avoids generating missing title/description/H1 issues for blocked URLs, non-HTML responses, HTTP errors, failed fetches, and incomplete/truncated extraction;
- labels missing metadata from raw HTML as “not found in raw HTML” with a notice severity unless rendered HTML was checked. This avoids claiming that a tag is definitely absent from the browser-rendered page.

### Optional JavaScript rendering

By default, the crawler audits the raw HTML returned by the server. Some sites inject or change metadata using JavaScript. To enable optional rendering:

```bat
.venv\Scripts\python.exe -m pip install -r requirements-rendering.txt
.venv\Scripts\python.exe -m playwright install chromium
```

Then open Settings in the dashboard, enable **JavaScript rendering**, and start a new crawl. Rendering can be slower and causes a real browser to load additional resources; use it only on sites you are authorized to audit and keep crawl limits/delays conservative. If Playwright or Chromium is unavailable, the diagnostics worksheet records the rendering error and the report does not silently claim rendered extraction.

### Important interpretation notes

- `Meta Description 1` represents a standard SEO meta description (`meta[name="description"]`), not an Open Graph `og:description` tag. The diagnostics sheet records if an Open Graph description exists separately.
- Pixel width is still a rough estimate, not a browser-measured pixel width.
- The crawler is not guaranteed to match Googlebot or every user session. Cookie consent, geography, user-agent, server personalization, blocked resources and changing content can produce different HTML.
- The `1 - All` worksheet is dynamic: all available fields are included, with detailed evidence also retained on `Audit Diagnostics` and `Page Data`.


## Dashboard data explorer update (v06)

- The dashboard starts with an empty workspace instead of fabricated demo crawl results.
- **Data Explorer** lists every worksheet in the actual crawler workbook and reads the selected worksheet using dynamic headers. It does not hard-code a column count.
- **URL Discovery** displays the URL Discovery worksheet, including discovery source and crawl status.
- Worksheet APIs are `GET /api/crawls/{crawl_id}/sheets` and `GET /api/crawls/{crawl_id}/sheets/{sheet_name}?limit=100&offset=0&search=`. They return actual workbook values and support pagination/search.
- The dashboard's Excel export downloads the original complete crawler workbook from the API, rather than reconstructing a limited workbook in the browser.
- Historical comparison UI no longer shows invented sample numbers. Automated crawl-to-crawl diffing remains unimplemented and is labelled as such.
- External Search Console / Analytics integrations are labelled disconnected until credentials and integration code are configured.

Validation performed: Python compilation passed; the crawler and dashboard-workbook API test suite passed (7 tests). A production Vite build could not be verified because npm dependency installation timed out in this environment.

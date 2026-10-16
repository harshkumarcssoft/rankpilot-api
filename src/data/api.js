/**
 * Shared RankPilot API client. Route names mirror api.py.
 * Production: set VITE_API_BASE_URL to the Render API base URL (including /api).
 */
const configuredApiBase = (import.meta.env.VITE_API_BASE_URL || '').trim();
const selectedApiBase = configuredApiBase || (import.meta.env.DEV ? '/api' : 'https://rankpilot-api-ovfh.onrender.com/api');
const trimmedApiBase = selectedApiBase.replace(/\/+$/, '');
const API_BASE = trimmedApiBase === '/api' || /\/api$/i.test(trimmedApiBase) ? trimmedApiBase : `${trimmedApiBase}/api`;

async function request(path, options = {}) {
  const url = `${API_BASE}${path.startsWith('/') ? path : `/${path}`}`;
  let response;
  try {
    response = await fetch(url, {
      ...options,
      headers: { Accept: 'application/json', ...(options.body ? { 'Content-Type': 'application/json' } : {}), ...(options.headers || {}) },
    });
  } catch (error) {
    throw new Error(`Could not reach RankPilot API at ${url}: ${error?.message || 'network or CORS error'}`);
  }
  const text = await response.text();
  let payload = null;
  if (text) {
    try { payload = JSON.parse(text); }
    catch { throw new Error(`RankPilot API returned non-JSON (HTTP ${response.status}) at ${url}: ${text.replace(/\s+/g, ' ').slice(0, 180)}`); }
  }
  if (!response.ok) {
    const detail = payload && typeof payload === 'object' ? (payload.detail || payload.message) : '';
    throw new Error(detail || `RankPilot API request failed (HTTP ${response.status}) at ${url}`);
  }
  return payload;
}

export const apiRequest = request;

export const auditApi = {
  health: () => request('/health'),
  projects: () => request('/projects'),
  createProject: (payload) => request('/projects', { method: 'POST', body: JSON.stringify(payload) }),
  deleteProject: (projectId) => request(`/projects/${encodeURIComponent(projectId)}`, { method: 'DELETE' }),
  crawls: (projectId) => request(`/projects/${encodeURIComponent(projectId)}/crawls`),
  startCrawl: (projectId, payload) => request(`/projects/${encodeURIComponent(projectId)}/crawls`, { method: 'POST', body: JSON.stringify(payload) }),
  pages: (crawlId, params = {}) => request(`/crawls/${encodeURIComponent(crawlId)}/pages?${new URLSearchParams(params)}`),
  sheets: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/sheets`),
  sheetData: (crawlId, sheetName, params = {}) => request(`/crawls/${encodeURIComponent(crawlId)}/sheets/${encodeURIComponent(sheetName)}?${new URLSearchParams(params)}`),
  progress: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/progress`),
  stop: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/stop`, { method: 'POST' }),
  downloadUrl: (crawlId) => `${API_BASE}/crawls/${encodeURIComponent(crawlId)}/download`,
};

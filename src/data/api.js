/**
 * API boundary for the SEO audit UI.
 * Components should consume these normalized models rather than depend on crawler internals.
 * Configure VITE_API_BASE_URL when a real backend is available.
 */
const API_BASE = (import.meta.env.VITE_API_BASE_URL || '/api').replace(/\/$/, '');

async function request(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, {
    headers: { Accept: 'application/json', ...(options.body ? { 'Content-Type': 'application/json' } : {}), ...options.headers },
    ...options,
  });
  if (!response.ok) throw new Error(`API request failed (${response.status})`);
  return response.status === 204 ? null : response.json();
}

export const auditApi = {
  projects: () => request('/projects'),
  crawls: (projectId) => request(`/projects/${encodeURIComponent(projectId)}/crawls`),
  pages: (crawlId, params = {}) => request(`/crawls/${encodeURIComponent(crawlId)}/pages?${new URLSearchParams(params)}`),
  issues: (crawlId, params = {}) => request(`/crawls/${encodeURIComponent(crawlId)}/issues?${new URLSearchParams(params)}`),
  crawl: (payload) => request('/crawls', { method: 'POST', body: JSON.stringify(payload) }),
  progress: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/progress`),
  pause: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/pause`, { method: 'POST' }),
  resume: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/resume`, { method: 'POST' }),
  stop: (crawlId) => request(`/crawls/${encodeURIComponent(crawlId)}/stop`, { method: 'POST' }),
  compare: (crawlA, crawlB) => request(`/crawls/compare?${new URLSearchParams({ a: crawlA, b: crawlB })}`),
};

/** JSON model contract shared by the frontend and crawler API. */
export const auditSchema = {
  page: ['url', 'status', 'indexability', 'title', 'titleLength', 'meta', 'metaLength', 'h1', 'wordCount', 'canonical', 'depth', 'inlinks', 'outlinks', 'response', 'size', 'issues'],
  issue: ['name', 'category', 'severity', 'url', 'evidence', 'fix', 'status'],
  crawlProgress: ['crawlId', 'state', 'crawled', 'queued', 'speedPerSecond', 'etaSeconds', 'updatedAt'],
};

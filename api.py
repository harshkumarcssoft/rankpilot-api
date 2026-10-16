"""Local SEO crawler API. Intended for local development only; add authentication before remote exposure."""
from __future__ import annotations
import json, os, sqlite3, subprocess, sys, time, uuid
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)
DB = DATA / "rankpilot.sqlite3"
RUNS: dict[str, subprocess.Popen] = {}

def now(): return datetime.now(timezone.utc).isoformat(timespec="seconds")
def connect():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init_db():
    with connect() as c:
        c.execute('''CREATE TABLE IF NOT EXISTS projects(id TEXT PRIMARY KEY,name TEXT NOT NULL,url TEXT NOT NULL,created_at TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS crawls(id TEXT PRIMARY KEY,project_id TEXT NOT NULL,started_at TEXT NOT NULL,status TEXT NOT NULL,max_pages INTEGER NOT NULL,delay REAL NOT NULL,output_path TEXT NOT NULL,progress_path TEXT NOT NULL,process_id INTEGER)''')
        c.execute('CREATE INDEX IF NOT EXISTS idx_crawls_project ON crawls(project_id,started_at DESC)')
init_db()
app=FastAPI(title="Rankpilot SEO Crawler API", version="0.3.0")
# Vite dev server only. Do not expose this unauthenticated API publicly.
# Allow local development and the deployed Vercel frontend. Add additional origins
# as a comma-separated CORS_ORIGINS environment variable on Render if needed.
_default_origins = "http://localhost:5173,http://127.0.0.1:5173,https://rankpilot-api-mu.vercel.app"
_allowed_origins = [origin.strip() for origin in os.getenv("CORS_ORIGINS", _default_origins).split(",") if origin.strip()]
# Vercel creates unique deployment aliases. Allow HTTPS Vercel origins in addition
# to the explicit production domain; custom domains can be added via CORS_ORIGINS.
app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed_origins,
    allow_origin_regex=r"https://([a-zA-Z0-9-]+\.)*vercel\.app",
    allow_credentials=False,
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)

class ProjectIn(BaseModel):
    name: str = Field(min_length=1, max_length=100)
    url: str = Field(min_length=8, max_length=2048)
class CrawlIn(BaseModel):
    max_pages: int = Field(default=100, ge=1, le=100000)
    delay: float = Field(default=1.0, ge=0.5, le=30)
    use_sitemap: bool = True
    respect_robots: bool = True
    render_js: bool = False
    user_agent: str = Field(default="SiteAuditCrawler/1.0 (+authorised technical SEO audit)", min_length=3, max_length=300)
    include_patterns_text: str = Field(default="", max_length=20000)
    exclude_patterns_text: str = Field(default="", max_length=20000)
    url_import_text: str = Field(default="", max_length=200000)
    url_import_file_base64: str = Field(default="", max_length=6000000)
    url_import_filename: str = Field(default="", max_length=255)
    feed_urls_text: str = Field(default="", max_length=10000)
    use_previous_crawl: bool = True

def validate_public_url(raw):
    from seo_crawler import public_host
    parsed=urlparse(raw.strip())
    if parsed.scheme not in ("http","https") or not parsed.hostname:
        raise HTTPException(400,"Enter a full http:// or https:// website URL.")
    if parsed.username or parsed.password:
        raise HTTPException(400,"URLs containing credentials are not allowed.")
    if not public_host(parsed.hostname):
        raise HTTPException(400,"The URL must resolve only to a public host. Local/private network targets are blocked.")
    return raw.strip().rstrip("/") or raw.strip()

def project_dict(r): return {"id":r["id"],"name":r["name"],"url":r["url"],"created_at":r["created_at"]}
def crawl_dict(r): return {"id":r["id"],"project_id":r["project_id"],"started_at":r["started_at"],"status":r["status"],"max_pages":r["max_pages"],"delay":r["delay"]}

@app.get('/')
def root(): return {"ok": True, "service": "rankpilot-crawler-api", "health": "/api/health", "docs": "/docs"}

@app.get('/api/health')
def health(): return {"ok":True,"service":"rankpilot-crawler-api","time":now()}
@app.get('/api/projects')
def list_projects():
    with connect() as c: return [project_dict(r) for r in c.execute('SELECT * FROM projects ORDER BY created_at DESC').fetchall()]
@app.post('/api/projects')
def create_project(body: ProjectIn):
    url=validate_public_url(body.url)
    pid=uuid.uuid4().hex
    with connect() as c: c.execute('INSERT INTO projects VALUES(?,?,?,?)',(pid,body.name.strip(),url,now()))
    return {"id":pid,"name":body.name.strip(),"url":url,"created_at":now()}
@app.delete('/api/projects/{project_id}')
def delete_project(project_id: str):
    with connect() as c:
        if not c.execute('SELECT id FROM projects WHERE id=?',(project_id,)).fetchone(): raise HTTPException(404,'Project not found')
        active=c.execute("SELECT id FROM crawls WHERE project_id=? AND status='running'",(project_id,)).fetchall()
        if active: raise HTTPException(409,'Stop the active crawl before deleting this project.')
        c.execute('DELETE FROM crawls WHERE project_id=?',(project_id,)); c.execute('DELETE FROM projects WHERE id=?',(project_id,))
    return {"ok":True}
@app.get('/api/projects/{project_id}/crawls')
def list_crawls(project_id: str):
    with connect() as c:
        if not c.execute('SELECT id FROM projects WHERE id=?',(project_id,)).fetchone(): raise HTTPException(404,'Project not found')
        return [crawl_dict(r) for r in c.execute('SELECT * FROM crawls WHERE project_id=? ORDER BY started_at DESC',(project_id,)).fetchall()]
@app.post('/api/projects/{project_id}/crawls')
def start_crawl(project_id: str, body: CrawlIn):
    with connect() as c: project=c.execute('SELECT * FROM projects WHERE id=?',(project_id,)).fetchone()
    if not project: raise HTTPException(404,'Project not found')
    url=validate_public_url(project['url'])
    with connect() as c:
        active=c.execute("SELECT id FROM crawls WHERE project_id=? AND status='running'",(project_id,)).fetchone()
    if active: raise HTTPException(409,'This project already has a crawl running.')
    cid=uuid.uuid4().hex
    folder=DATA/'projects'/project_id/'crawls'/cid; folder.mkdir(parents=True,exist_ok=True)
    output=folder/'seo_audit_report.xlsx'; progress=folder/'progress.json'; log=folder/'crawler.log'
    cmd=[sys.executable,str(BASE/'seo_crawler.py'),url,'--max-pages',str(body.max_pages),'--delay',str(body.delay),'--output',str(output),'--progress',str(progress)]
    if not body.use_sitemap: cmd.append('--no-sitemap')
    if not body.respect_robots: cmd.append('--ignore-robots')
    if body.render_js: cmd.append('--render-js')
    cmd.extend(['--user-agent', body.user_agent.strip()])
    for pattern in body.include_patterns_text.splitlines():
        pattern=pattern.strip()
        if pattern: cmd.extend(['--include', pattern])
    for pattern in body.exclude_patterns_text.splitlines():
        pattern=pattern.strip()
        if pattern: cmd.extend(['--exclude', pattern])
    if body.url_import_text.strip():
        import_file=folder/'manual_urls.txt'
        import_file.write_text(body.url_import_text,encoding='utf-8')
        cmd.extend(['--url-file',str(import_file)])
    if body.url_import_file_base64.strip():
        import base64, binascii
        suffix=Path(body.url_import_filename or 'url-inventory.txt').suffix.lower()
        if suffix not in {'.txt','.csv','.xlsx','.xlsm'}:
            raise HTTPException(400,'URL inventory files must be .txt, .csv, .xlsx or .xlsm.')
        try:
            file_bytes=base64.b64decode(body.url_import_file_base64,validate=True)
        except (binascii.Error, ValueError):
            raise HTTPException(400,'The URL inventory file could not be decoded. Please reselect the file and try again.')
        if len(file_bytes)>4_000_000:
            raise HTTPException(413,'URL inventory files must be 4 MB or smaller.')
        import_file=folder/('manual_inventory'+suffix)
        import_file.write_bytes(file_bytes)
        cmd.extend(['--url-file',str(import_file)])
    for feed in body.feed_urls_text.splitlines():
        feed=feed.strip()
        if feed: cmd.extend(['--feed',feed])
    if body.use_previous_crawl:
        with connect() as c:
            previous=c.execute("SELECT output_path FROM crawls WHERE project_id=? AND status IN ('finished','complete') ORDER BY started_at DESC LIMIT 1",(project_id,)).fetchone()
        if previous and Path(previous['output_path']).exists(): cmd.extend(['--previous-report',previous['output_path']])
    try:
        with open(log,'w',encoding='utf-8') as logf:
            proc=subprocess.Popen(cmd,cwd=str(BASE),stdout=logf,stderr=subprocess.STDOUT,creationflags=(subprocess.CREATE_NEW_PROCESS_GROUP if os.name=='nt' else 0))
    except Exception as e: raise HTTPException(500,f'Could not start crawler: {e}')
    RUNS[cid]=proc
    with connect() as c: c.execute('INSERT INTO crawls VALUES(?,?,?,?,?,?,?,?,?)',(cid,project_id,now(),'running',body.max_pages,body.delay,str(output),str(progress),proc.pid))
    return {"id":cid,"project_id":project_id,"status":"running","started_at":now(),"max_pages":body.max_pages,"delay":body.delay}
@app.get('/api/crawls/{crawl_id}/progress')
def crawl_progress(crawl_id: str):
    with connect() as c: r=c.execute('SELECT * FROM crawls WHERE id=?',(crawl_id,)).fetchone()
    if not r: raise HTTPException(404,'Crawl not found')
    path=Path(r['progress_path']); p={}
    if path.exists():
        try: p=json.loads(path.read_text(encoding='utf-8'))
        except Exception: p={}
    proc=RUNS.get(crawl_id)
    state=p.get('status',r['status'])
    if proc and proc.poll() is not None and state=='running':
        state='failed'; p['message']='Crawler process exited unexpectedly. Check crawler.log.'
    with connect() as c: c.execute('UPDATE crawls SET status=? WHERE id=?',(state,crawl_id))
    crawled=int(p.get('crawled',0)); max_pages=int(p.get('max_pages',r['max_pages']))
    try: elapsed=max(1.0,(datetime.now(timezone.utc)-datetime.fromisoformat(r['started_at'])).total_seconds())
    except Exception: elapsed=1.0
    speed=round(crawled/elapsed,2); remaining=max(0,max_pages-crawled); eta=round(remaining/speed) if speed>0 and state=='running' else 0
    return {"crawlId":crawl_id,"state":state,"crawled":crawled,"queued":int(p.get('queued',0)) if state=='running' else 0,"maxPages":max_pages,"currentUrl":p.get('current_url',''),"message":p.get('message','Waiting for crawler output'),"errors":int(p.get('errors',0)),"speedPerSecond":speed,"etaSeconds":eta,"updatedAt":p.get('updated_at'),"reportReady":Path(r['output_path']).exists()}
@app.post('/api/crawls/{crawl_id}/stop')
def stop_crawl(crawl_id: str):
    with connect() as c: r=c.execute('SELECT * FROM crawls WHERE id=?',(crawl_id,)).fetchone()
    if not r: raise HTTPException(404,'Crawl not found')
    proc=RUNS.get(crawl_id)
    if proc and proc.poll() is None:
        proc.terminate()
        try: proc.wait(timeout=4)
        except subprocess.TimeoutExpired: proc.kill()
    with connect() as c: c.execute("UPDATE crawls SET status='stopped' WHERE id=?",(crawl_id,))
    pp=Path(r['progress_path'])
    try: pp.write_text(json.dumps({"status":"stopped","crawled":0,"max_pages":r['max_pages'],"message":"Stopped by user","updated_at":now()}),encoding='utf-8')
    except OSError: pass
    return {"ok":True,"status":"stopped"}
@app.get('/api/crawls/{crawl_id}/pages')
def crawl_pages(crawl_id: str, limit: int=100000, offset: int=0):
    with connect() as c: r=c.execute('SELECT * FROM crawls WHERE id=?',(crawl_id,)).fetchone()
    if not r: raise HTTPException(404,'Crawl not found')
    path=Path(r['output_path'])
    if not path.exists(): return {"items":[],"total":0,"ready":False}
    try:
        wb=load_workbook(path,read_only=True,data_only=True); ws=wb['1 - All'] if '1 - All' in wb.sheetnames else wb[wb.sheetnames[0]]
        diagnostics={}
        if 'Audit Diagnostics' in wb.sheetnames:
            dws=wb['Audit Diagnostics']; dit=dws.iter_rows(values_only=True); dheaders=[str(x) if x is not None else '' for x in next(dit,[])]
            for values in dit:
                d={h:(v.isoformat() if hasattr(v,'isoformat') else v) for h,v in zip(dheaders,values) if h}
                if d.get('Address'): diagnostics[str(d['Address'])]=d
        it=ws.iter_rows(values_only=True); headers=[str(x) if x is not None else '' for x in next(it,[])]; rows=[]
        # Keep the dashboard response focused: large visible-text/raw JSON fields remain
        # available in the paginated Data Explorer and full Excel workbook, not duplicated
        # into every dashboard page payload.
        from seo_crawler import HEADERS as BASE_PAGE_HEADERS
        dashboard_page_fields={
            'Page Data: Image data JSON','Page Data: Links data JSON','Page Data: Structured data JSON',
            'Page Data: Resources JSON','Page Data: Forms JSON','Page Data: Hreflang links JSON',
            'Page Data: Accessibility signals JSON','Page Data: Local SEO presence signals JSON',
            'Page Data: Open Graph JSON','Page Data: Twitter card JSON','Page Data: Headings JSON',
            'Page Data: Viewport','Page Data: Extraction notes'
        }
        keep_headers=set(BASE_PAGE_HEADERS)|dashboard_page_fields
        for idx,values in enumerate(it):
            if idx < offset: continue
            if len(rows)>=max(1,min(limit,100000)): break
            row={h:(v.isoformat() if hasattr(v,'isoformat') else v) for h,v in zip(headers,values) if h and h in keep_headers}
            if row.get('Address'):
                row['__diagnostics']=diagnostics.get(str(row['Address']),{})
                rows.append(row)
        total=max(0,ws.max_row-1); wb.close()
        return {"items":rows,"total":total,"ready":True,"sheet":"1 - All"}
    except Exception as e: raise HTTPException(500,f'Could not read crawl report: {e}')
@app.get('/api/crawls/{crawl_id}/download')
def crawl_download(crawl_id: str):
    from fastapi.responses import FileResponse
    with connect() as c: r=c.execute('SELECT * FROM crawls WHERE id=?',(crawl_id,)).fetchone()
    if not r: raise HTTPException(404,'Crawl not found')
    p=Path(r['output_path'])
    if not p.exists(): raise HTTPException(404,'Report is not ready yet')
    return FileResponse(p,filename=f"seo-audit-{crawl_id[:8]}.xlsx",media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.get('/api/crawls/{crawl_id}/sheets')
def crawl_sheets(crawl_id: str):
    """List every worksheet in a crawl workbook, including row/column counts."""
    with connect() as c: r=c.execute('SELECT * FROM crawls WHERE id=?',(crawl_id,)).fetchone()
    if not r: raise HTTPException(404,'Crawl not found')
    path=Path(r['output_path'])
    if not path.exists(): return {"ready":False,"sheets":[]}
    try:
        wb=load_workbook(path,read_only=True,data_only=True)
        sheets=[]
        for ws in wb.worksheets:
            headers=[str(v) if v is not None else '' for v in next(ws.iter_rows(min_row=1,max_row=1,values_only=True),())]
            sheets.append({"name":ws.title,"rows":max(0,(ws.max_row or 1)-1),"columns":len([h for h in headers if h]),"headers":headers})
        wb.close()
        return {"ready":True,"sheets":sheets}
    except Exception as e: raise HTTPException(500,f'Could not inspect crawl workbook: {e}')

@app.get('/api/crawls/{crawl_id}/sheets/{sheet_name}')
def crawl_sheet_data(crawl_id: str, sheet_name: str, limit: int=100, offset: int=0, search: str=''):
    """Read any worksheet dynamically with bounded pagination and optional row search."""
    with connect() as c: r=c.execute('SELECT * FROM crawls WHERE id=?',(crawl_id,)).fetchone()
    if not r: raise HTTPException(404,'Crawl not found')
    path=Path(r['output_path'])
    if not path.exists(): return {"ready":False,"sheet":sheet_name,"headers":[],"items":[],"total":0}
    try:
        wb=load_workbook(path,read_only=True,data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close(); raise HTTPException(404,'Worksheet not found')
        ws=wb[sheet_name]
        it=ws.iter_rows(values_only=True)
        raw_headers=next(it,())
        headers=[str(v) if v is not None else '' for v in raw_headers]
        items=[]; total=0; offset=max(0,offset); limit=max(1,min(limit,1000)); needle=search.casefold().strip()
        for values in it:
            row={h:(v.isoformat() if hasattr(v,'isoformat') else v) for h,v in zip(headers,values) if h}
            if needle and needle not in json.dumps(row,ensure_ascii=False,default=str).casefold(): continue
            if total >= offset and len(items)<limit: items.append(row)
            total += 1
        wb.close()
        return {"ready":True,"sheet":sheet_name,"headers":headers,"items":items,"total":total,"offset":offset,"limit":limit}
    except HTTPException: raise
    except Exception as e: raise HTTPException(500,f'Could not read worksheet: {e}')

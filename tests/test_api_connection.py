from fastapi.testclient import TestClient
import api
import seo_crawler


def test_health_and_root_endpoints_are_json():
    client = TestClient(api.app)
    health = client.get("/api/health")
    assert health.status_code == 200
    assert health.json()["ok"] is True
    assert health.json()["service"] == "rankpilot-crawler-api"
    root = client.get("/")
    assert root.status_code == 200
    assert root.json()["health"] == "/api/health"


def test_vercel_deployment_origin_passes_cors_preflight():
    client = TestClient(api.app)
    response = client.options(
        "/api/projects",
        headers={
            "Origin": "https://rankpilot-ctbocoj6z-game-over2.vercel.app",
            "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "content-type",
        },
    )
    assert response.status_code == 200
    assert response.headers["access-control-allow-origin"] == "https://rankpilot-ctbocoj6z-game-over2.vercel.app"


def test_create_project_returns_json_and_project_is_listed(monkeypatch):
    monkeypatch.setattr(seo_crawler, "public_host", lambda host: True)
    client = TestClient(api.app)
    response = client.post("/api/projects", json={"name": "test connection", "url": "https://example.com/"})
    assert response.status_code == 200
    payload = response.json()
    assert payload["name"] == "test connection"
    assert payload["url"] == "https://example.com"
    project_id = payload["id"]
    try:
        listed = client.get("/api/projects")
        assert listed.status_code == 200
        assert any(project["id"] == project_id for project in listed.json())
    finally:
        with api.connect() as conn:
            conn.execute("DELETE FROM crawls WHERE project_id=?", (project_id,))
            conn.execute("DELETE FROM projects WHERE id=?", (project_id,))


def test_start_crawl_passes_dashboard_settings_and_inventory_files(monkeypatch):
    import base64
    from pathlib import Path

    monkeypatch.setattr(seo_crawler, "public_host", lambda host: True)
    captured = {}

    class FakeProcess:
        pid = 987654
        def poll(self):
            return None

    def fake_popen(cmd, **kwargs):
        captured["cmd"] = cmd
        return FakeProcess()

    monkeypatch.setattr(api.subprocess, "Popen", fake_popen)
    client = TestClient(api.app)
    project = client.post("/api/projects", json={"name": "settings test", "url": "https://example.com"})
    assert project.status_code == 200
    project_id = project.json()["id"]
    inventory = base64.b64encode(b"https://example.com/one\nhttps://example.com/two\n").decode("ascii")
    crawl = client.post(f"/api/projects/{project_id}/crawls", json={
        "max_pages": 20,
        "delay": 0.5,
        "use_sitemap": False,
        "respect_robots": False,
        "render_js": False,
        "user_agent": "RankPilotAuthorisedTest/1.0",
        "include_patterns_text": "/blog/*",
        "exclude_patterns_text": "/cart/*",
        "url_import_text": "https://example.com/manual-page",
        "url_import_file_base64": inventory,
        "url_import_filename": "inventory.txt",
        "feed_urls_text": "/feed.xml",
        "use_previous_crawl": False,
    })
    assert crawl.status_code == 200
    cmd = captured["cmd"]
    assert "--no-sitemap" in cmd
    assert "--ignore-robots" in cmd
    assert "--user-agent" in cmd and "RankPilotAuthorisedTest/1.0" in cmd
    assert cmd.count("--include") == 1 and "/blog/*" in cmd
    assert cmd.count("--exclude") == 1 and "/cart/*" in cmd
    assert cmd.count("--url-file") == 2
    assert cmd.count("--feed") == 1 and "/feed.xml" in cmd
    crawl_id = crawl.json()["id"]
    api.RUNS.pop(crawl_id, None)
    with api.connect() as conn:
        conn.execute("DELETE FROM crawls WHERE project_id=?", (project_id,))
        conn.execute("DELETE FROM projects WHERE id=?", (project_id,))

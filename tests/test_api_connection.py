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

"""Integration tests for dashboard access to every crawler workbook worksheet."""
import tempfile
import uuid
from pathlib import Path

from fastapi.testclient import TestClient
from openpyxl import Workbook

import api


def test_dashboard_can_list_and_read_all_workbook_sheets():
    project_id = uuid.uuid4().hex
    crawl_id = uuid.uuid4().hex
    tmp = Path(tempfile.mkdtemp())
    workbook_path = tmp / "crawl.xlsx"

    wb = Workbook()
    main = wb.active
    main.title = "1 - All"
    main.append(["Address", "Status Code", "Page Data: Image data JSON", "A future dynamic field"])
    main.append(["https://example.com/", 200, "[]", "kept"])
    for name, headers, values in [
        ("Page Data", ["Address", "Visible text", "H1 count"], ["https://example.com/", "Example page", 1]),
        ("URL Discovery", ["URL", "Discovery source", "Discovery status"], ["https://example.com/", "seed_url", "crawled"]),
        ("Audit Diagnostics", ["Address", "Extraction status"], ["https://example.com/", "checked"]),
    ]:
        sheet = wb.create_sheet(name)
        sheet.append(headers)
        sheet.append(values)
    wb.save(workbook_path)

    with api.connect() as conn:
        conn.execute("INSERT INTO projects VALUES(?,?,?,?)", (project_id, "test dashboard", "https://example.com", "2026-10-16T00:00:00+00:00"))
        conn.execute("INSERT INTO crawls VALUES(?,?,?,?,?,?,?,?,?)", (crawl_id, project_id, "2026-10-16T00:00:00+00:00", "finished", 10, 0.5, str(workbook_path), str(tmp / "progress.json"), None))

    try:
        client = TestClient(api.app)
        listing = client.get(f"/api/crawls/{crawl_id}/sheets")
        assert listing.status_code == 200
        names = {s["name"] for s in listing.json()["sheets"]}
        assert {"1 - All", "Page Data", "URL Discovery", "Audit Diagnostics"}.issubset(names)
        all_sheet = client.get(f"/api/crawls/{crawl_id}/sheets/1%20-%20All?limit=10")
        assert all_sheet.status_code == 200
        assert "A future dynamic field" in all_sheet.json()["headers"]
        assert all_sheet.json()["items"][0]["A future dynamic field"] == "kept"
        discovery = client.get(f"/api/crawls/{crawl_id}/sheets/URL%20Discovery?limit=10")
        assert discovery.status_code == 200
        assert discovery.json()["items"][0]["Discovery source"] == "seed_url"
    finally:
        with api.connect() as conn:
            conn.execute("DELETE FROM crawls WHERE id=?", (crawl_id,))
            conn.execute("DELETE FROM projects WHERE id=?", (project_id,))

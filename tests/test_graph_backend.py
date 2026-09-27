"""Testa o GraphBackend sem rede (MSAL e HTTP simulados)."""

from __future__ import annotations

import pytest

from finance import storage
from finance.config import GraphSettings


class FakeResp:
    def __init__(self, status=200, json_data=None, content=b""):
        self.status_code, self._json, self.content = status, json_data or {}, content

    @property
    def ok(self):
        return 200 <= self.status_code < 300

    def json(self):
        return self._json


@pytest.fixture
def backend(monkeypatch):
    class FakeApp:
        def __init__(self, *a, **k):
            pass

        def acquire_token_by_refresh_token(self, rt, scopes):
            return {"access_token": "tok", "refresh_token": "rt2", "expires_in": 3600}

    monkeypatch.setattr(storage.msal, "PublicClientApplication", FakeApp)
    s = GraphSettings(auth="refresh_token", client_id="cid", tenant_id="consumers",
                      file_path="Financas/controle financeiro.xlsx", refresh_token="rt1")
    return storage.GraphBackend(s)


def test_urls(backend):
    assert backend.meta_url.endswith("/me/drive/root:/Financas/controle%20financeiro.xlsx")
    assert backend.content_url.endswith(":/content")


def test_read_write_and_conflict(backend, monkeypatch):
    calls = []

    def fake_request(method, url, headers=None, timeout=None, **kw):
        calls.append((method, url, dict(headers or {})))
        if method == "GET":
            return FakeResp(json_data={"eTag": "e1", "@microsoft.graph.downloadUrl": "https://dl"})
        if headers.get("If-Match") == "velho":
            return FakeResp(412, {"error": {"message": "precondition failed"}})
        return FakeResp(200, {"eTag": "e2"})

    monkeypatch.setattr(storage.requests, "request", fake_request)
    monkeypatch.setattr(storage.requests, "get", lambda url, timeout: FakeResp(content=b"xlsx"))

    assert backend.read() == (b"xlsx", "e1")
    assert backend.write(b"novo", "e1") == "e2"
    assert calls[-1][2]["If-Match"] == "e1" and calls[-1][2]["Authorization"] == "Bearer tok"
    with pytest.raises(storage.ConflictError):
        backend.write(b"novo", "velho")


def test_missing_file_and_config_errors(backend, monkeypatch):
    monkeypatch.setattr(storage.requests, "request", lambda *a, **k: FakeResp(404))
    assert backend.read() == (None, None)
    with pytest.raises(storage.StorageError, match="refresh_token"):
        storage.GraphBackend(GraphSettings(auth="refresh_token", client_id="x",
                                           tenant_id="consumers", file_path="a.xlsx"))

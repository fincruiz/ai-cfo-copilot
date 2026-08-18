import pytest
from app.services.auth_service import _auth_redirect_url
def test_production_rejects_localhost(monkeypatch):
 from app.services import auth_service; monkeypatch.setattr(auth_service.settings,"environment","production"); monkeypatch.setattr(auth_service.settings,"auth_frontend_url","http://localhost:3000")
 with pytest.raises(RuntimeError): _auth_redirect_url("/auth/callback")
def test_production_https(monkeypatch):
 from app.services import auth_service; monkeypatch.setattr(auth_service.settings,"environment","production"); monkeypatch.setattr(auth_service.settings,"auth_frontend_url","https://app.example.com")
 assert _auth_redirect_url("/auth/callback")=="https://app.example.com/auth/callback"

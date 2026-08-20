"""Judoscale request-queue-time reporting is wired and actually activates.

The failure mode this guards is silent: judoscale installs cleanly, logs
"Not activated - no API URL provided", and reports nothing — autoscaling sees no
metrics while everything looks healthy.
"""

from fastapi.testclient import TestClient

from app.main import app, judoscale_api_url

MIDDLEWARE = "FastAPIRequestQueueTimeMiddleware"


def _entry():
    return next(m for m in app.user_middleware if m.cls.__name__ == MIDDLEWARE)


def _options(entry):
    return getattr(entry, "kwargs", None) or getattr(entry, "options", {})


def test_middleware_is_installed():
    assert MIDDLEWARE in [m.cls.__name__ for m in app.user_middleware]


def test_requests_still_succeed_through_the_middleware():
    with TestClient(app) as client:
        assert client.get("/health").status_code == 200


def test_prefers_judoscales_own_variable():
    env = {"JUDOSCALE_URL": "https://judo", "RAILS_AUTOSCALE_URL": "https://rails"}
    assert judoscale_api_url(env) == "https://judo"


def test_falls_back_to_the_heroku_addon_variable():
    """The add-on exports RAILS_AUTOSCALE_URL; judoscale reads only JUDOSCALE_URL,
    so without this fallback the integration is a silent no-op."""
    assert judoscale_api_url({"RAILS_AUTOSCALE_URL": "https://rails"}) == "https://rails"


def test_returns_none_when_neither_is_set():
    assert judoscale_api_url({}) is None


def test_url_is_passed_as_extra_config_not_via_os_environ():
    """Regression: judoscale builds its config singleton at IMPORT time, so
    assigning os.environ["JUDOSCALE_URL"] below the import in app.main is already
    too late — is_enabled stays False and nothing is ever reported. The middleware
    calls judoconfig.update(extra_config) before checking is_enabled, which is the
    supported way in."""
    options = _options(_entry())
    assert "extra_config" in options
    if JUDOSCALE_CONFIGURED := judoscale_api_url():
        assert options["extra_config"]["API_BASE_URL"] == JUDOSCALE_CONFIGURED
    else:
        assert options["extra_config"] == {}

import httpx
import pytest

from weather_app import client as client_module
from weather_app.client import WeatherFetchError, fetch_seoul_current_weather


_SAMPLE_PAYLOAD = {
    "current": {
        "time": "2026-05-23T14:00",
        "temperature_2m": 18.3,
        "relative_humidity_2m": 65,
        "wind_speed_10m": 2.4,
        "weather_code": 1,
    }
}


def _install_mock(monkeypatch, handler):
    transport = httpx.MockTransport(handler)
    original_get = httpx.get

    def patched_get(url, **kwargs):
        kwargs.pop("timeout", None)
        with httpx.Client(transport=transport) as c:
            return c.get(url, **kwargs)

    monkeypatch.setattr(client_module.httpx, "get", patched_get)
    return original_get


def test_fetch_parses_payload(monkeypatch):
    def handler(request: httpx.Request) -> httpx.Response:
        assert "open-meteo.com" in request.url.host
        assert request.url.params["latitude"] == "37.5665"
        return httpx.Response(200, json=_SAMPLE_PAYLOAD)

    _install_mock(monkeypatch, handler)
    result = fetch_seoul_current_weather()
    assert result.temperature_c == 18.3
    assert result.humidity_pct == 65
    assert result.wind_ms == 2.4
    assert result.weather_code == 1


def test_fetch_raises_on_http_error(monkeypatch):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, text="boom")

    _install_mock(monkeypatch, handler)
    with pytest.raises(WeatherFetchError):
        fetch_seoul_current_weather()


def test_fetch_raises_on_malformed_payload(monkeypatch):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json={"current": {}})

    _install_mock(monkeypatch, handler)
    with pytest.raises(WeatherFetchError):
        fetch_seoul_current_weather()

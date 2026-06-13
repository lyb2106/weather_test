from datetime import datetime

from weather_app.models import CurrentWeather
from weather_app.renderer import render_html, render_to_file


def _sample() -> CurrentWeather:
    return CurrentWeather(
        temperature_c=18.3,
        humidity_pct=65,
        wind_ms=2.4,
        weather_code=1,
        observed_at=datetime(2026, 5, 23, 14, 0),
    )


def test_render_html_contains_core_values():
    html = render_html(_sample())
    assert "18.3" in html
    assert "65%" in html
    assert "2.4 m/s" in html
    assert "대체로 맑음" in html
    assert "서울" in html
    assert "2026-05-23 14:00 KST" in html


def test_render_html_has_html_structure():
    html = render_html(_sample())
    assert html.lstrip().startswith("<!DOCTYPE html>")
    assert "</html>" in html
    assert 'lang="ko"' in html


def test_render_to_file_writes_utf8(tmp_path):
    out = tmp_path / "weather.html"
    result = render_to_file(_sample(), out)
    assert result == out
    content = out.read_text(encoding="utf-8")
    assert "서울" in content
    assert "18.3" in content

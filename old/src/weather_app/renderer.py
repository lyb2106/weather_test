"""CurrentWeather → 시각화된 HTML 페이지 렌더링."""

from __future__ import annotations

from html import escape
from pathlib import Path

from .formatter import (
    weather_code_to_emoji,
    weather_code_to_korean,
    weather_code_to_theme,
)
from .models import CurrentWeather

_THEME_GRADIENTS: dict[str, str] = {
    "sunny": "linear-gradient(160deg, #87ceeb 0%, #ffd89b 100%)",
    "cloudy": "linear-gradient(160deg, #8e9eab 0%, #d7dde8 100%)",
    "rain": "linear-gradient(160deg, #4b6cb7 0%, #8ea4c8 100%)",
    "snow": "linear-gradient(160deg, #c9d6ff 0%, #f6f8ff 100%)",
    "storm": "linear-gradient(160deg, #232526 0%, #5c6373 100%)",
}

_THEME_TEXT_COLOR: dict[str, str] = {
    "sunny": "#1f2937",
    "cloudy": "#1f2937",
    "rain": "#f8fafc",
    "snow": "#1f2937",
    "storm": "#f8fafc",
}


def _format_observed_at(weather: CurrentWeather) -> str:
    return weather.observed_at.strftime("%Y-%m-%d %H:%M KST")


def render_html(weather: CurrentWeather) -> str:
    theme = weather_code_to_theme(weather.weather_code)
    background = _THEME_GRADIENTS[theme]
    text_color = _THEME_TEXT_COLOR[theme]
    emoji = weather_code_to_emoji(weather.weather_code)
    status = escape(weather_code_to_korean(weather.weather_code))
    observed = escape(_format_observed_at(weather))
    temperature = f"{weather.temperature_c:.1f}"
    humidity = f"{weather.humidity_pct}"
    wind = f"{weather.wind_ms:.1f}"

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>서울 현재 날씨</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; }}
  html, body {{ margin: 0; padding: 0; height: 100%; }}
  body {{
    font-family: "Pretendard", "Noto Sans KR", -apple-system, BlinkMacSystemFont,
      "Segoe UI", system-ui, sans-serif;
    background: {background};
    color: {text_color};
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 24px;
  }}
  .card {{
    width: 100%;
    max-width: 480px;
    background: rgba(255, 255, 255, 0.88);
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
    border-radius: 24px;
    box-shadow: 0 20px 50px rgba(0, 0, 0, 0.18);
    padding: 36px 32px 28px;
    color: #1f2937;
    text-align: center;
  }}
  .header {{
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 4px;
  }}
  .city {{
    font-size: 22px;
    font-weight: 700;
    letter-spacing: -0.02em;
  }}
  .observed-at {{
    font-size: 13px;
    color: #6b7280;
  }}
  .emoji {{
    font-size: 88px;
    line-height: 1;
    margin: 18px 0 6px;
  }}
  .status {{
    font-size: 18px;
    font-weight: 600;
    color: #374151;
    margin-bottom: 12px;
  }}
  .temperature {{
    font-size: 72px;
    font-weight: 700;
    letter-spacing: -0.04em;
    line-height: 1;
    margin: 8px 0 24px;
  }}
  .temperature .unit {{
    font-size: 36px;
    font-weight: 500;
    color: #6b7280;
    margin-left: 2px;
  }}
  .metrics {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    margin-top: 4px;
  }}
  .metric {{
    background: rgba(243, 244, 246, 0.85);
    border-radius: 16px;
    padding: 14px 12px;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 4px;
  }}
  .metric .icon {{ font-size: 24px; }}
  .metric .label {{
    font-size: 12px;
    color: #6b7280;
    font-weight: 500;
  }}
  .metric .value {{
    font-size: 20px;
    font-weight: 700;
    color: #111827;
  }}
  .footer {{
    margin-top: 22px;
    font-size: 11px;
    color: #9ca3af;
    letter-spacing: 0.04em;
  }}
  .footer a {{ color: inherit; text-decoration: none; }}
</style>
</head>
<body>
  <main class="card" role="main">
    <div class="header">
      <div class="city">서울</div>
      <div class="observed-at">{observed}</div>
    </div>
    <div class="emoji" aria-hidden="true">{emoji}</div>
    <div class="status">{status}</div>
    <div class="temperature">{temperature}<span class="unit">°C</span></div>
    <div class="metrics">
      <div class="metric">
        <div class="icon" aria-hidden="true">💧</div>
        <div class="label">습도</div>
        <div class="value">{humidity}%</div>
      </div>
      <div class="metric">
        <div class="icon" aria-hidden="true">💨</div>
        <div class="label">풍속</div>
        <div class="value">{wind} m/s</div>
      </div>
    </div>
    <div class="footer">출처: <a href="https://open-meteo.com" target="_blank" rel="noopener">Open-Meteo</a></div>
  </main>
</body>
</html>
"""


def render_to_file(weather: CurrentWeather, path: Path) -> Path:
    html = render_html(weather)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
    return path

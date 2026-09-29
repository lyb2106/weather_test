"""Open-Meteo API 클라이언트."""

from __future__ import annotations

import httpx

from .models import CurrentWeather

SEOUL_LAT = 37.5665
SEOUL_LON = 126.9780
ENDPOINT = "https://api.open-meteo.com/v1/forecast"
TIMEOUT_SECONDS = 5.0


class WeatherFetchError(RuntimeError):
    """외부 API 호출 또는 응답 파싱 실패."""


def fetch_seoul_current_weather() -> CurrentWeather:
    params = {
        "latitude": SEOUL_LAT,
        "longitude": SEOUL_LON,
        "current": "temperature_2m,relative_humidity_2m,wind_speed_10m,weather_code",
        "timezone": "Asia/Seoul",
        "wind_speed_unit": "ms",
    }
    try:
        response = httpx.get(ENDPOINT, params=params, timeout=TIMEOUT_SECONDS)
        response.raise_for_status()
        payload = response.json()
    except httpx.HTTPError as exc:
        raise WeatherFetchError(f"Open-Meteo 요청 실패: {exc}") from exc
    except ValueError as exc:
        raise WeatherFetchError(f"응답 JSON 파싱 실패: {exc}") from exc

    try:
        return CurrentWeather.from_open_meteo(payload)
    except (KeyError, TypeError, ValueError) as exc:
        raise WeatherFetchError(f"응답 형식이 예상과 다릅니다: {exc}") from exc

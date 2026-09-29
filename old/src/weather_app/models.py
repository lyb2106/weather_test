"""Open-Meteo 응답을 표현하는 도메인 모델."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class CurrentWeather:
    temperature_c: float
    humidity_pct: int
    wind_ms: float
    weather_code: int
    observed_at: datetime

    @classmethod
    def from_open_meteo(cls, payload: dict) -> "CurrentWeather":
        current = payload["current"]
        return cls(
            temperature_c=float(current["temperature_2m"]),
            humidity_pct=int(current["relative_humidity_2m"]),
            wind_ms=float(current["wind_speed_10m"]),
            weather_code=int(current["weather_code"]),
            observed_at=datetime.fromisoformat(current["time"]),
        )

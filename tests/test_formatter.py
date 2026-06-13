from weather_app.formatter import (
    weather_code_to_emoji,
    weather_code_to_korean,
    weather_code_to_theme,
)


def test_korean_known_codes():
    assert weather_code_to_korean(0) == "맑음"
    assert weather_code_to_korean(3) == "흐림"
    assert weather_code_to_korean(63) == "비"
    assert weather_code_to_korean(95) == "뇌우"


def test_korean_unknown_code_falls_back():
    assert weather_code_to_korean(999) == "알 수 없음"


def test_emoji_known_codes():
    assert weather_code_to_emoji(0) == "☀️"
    assert weather_code_to_emoji(95) == "⛈️"


def test_emoji_unknown_code_falls_back():
    assert weather_code_to_emoji(999) == "🌡️"


def test_theme_buckets():
    assert weather_code_to_theme(0) == "sunny"
    assert weather_code_to_theme(3) == "cloudy"
    assert weather_code_to_theme(63) == "rain"
    assert weather_code_to_theme(73) == "snow"
    assert weather_code_to_theme(95) == "storm"

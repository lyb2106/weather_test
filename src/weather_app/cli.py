"""`uv run weather` CLI 진입점."""

from __future__ import annotations

import argparse
import sys
import webbrowser
from pathlib import Path

from .client import WeatherFetchError, fetch_seoul_current_weather
from .renderer import render_to_file

_DEFAULT_OUTPUT = Path("output") / "seoul-weather.html"


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="weather",
        description="서울 현재 날씨를 시각화된 HTML 파일로 저장합니다.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=_DEFAULT_OUTPUT,
        help=f"HTML 출력 경로 (기본: {_DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--open",
        dest="open_in_browser",
        action="store_true",
        help="생성 후 기본 브라우저로 자동 열기",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")

    args = _build_parser().parse_args(argv)

    try:
        weather = fetch_seoul_current_weather()
    except WeatherFetchError as exc:
        print(f"날씨 정보를 가져오지 못했습니다: {exc}", file=sys.stderr)
        return 1

    output_path = render_to_file(weather, args.output)
    resolved = output_path.resolve()
    print(f"서울 현재 날씨를 생성했습니다 → {resolved}")

    if args.open_in_browser:
        webbrowser.open(resolved.as_uri())

    return 0


if __name__ == "__main__":
    sys.exit(main())

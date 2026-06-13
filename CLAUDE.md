# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

서울의 현재 날씨를 Open-Meteo API에서 받아와 자기완결형(self-contained) HTML 파일로 시각화하는 CLI 도구. 외부 폰트/CDN 의존 없이 HTML 한 파일에 CSS 인라인으로 카드 UI 를 생성한다.

상세 기획은 `PROJECT_PLAN.md` 에 있다 - 새 기능을 추가하기 전 참고할 것.

## 주요 명령어

| 작업 | 명령 |
|------|------|
| 날씨 HTML 생성 | `uv run weather` |
| 생성 + 브라우저로 자동 열기 | `uv run weather --open` |
| 출력 경로 지정 | `uv run weather --output <path>` |
| 전체 테스트 | `uv run pytest` |
| 단일 테스트 파일 | `uv run pytest tests/test_renderer.py` |
| 단일 테스트 함수 | `uv run pytest tests/test_renderer.py::test_render_html_contains_core_values` |
| 린트 | `uv run ruff check .` |
| 의존성 추가 | `uv add <pkg>` / `uv add --dev <pkg>` |

기본 출력 경로는 `output/seoul-weather.html` (gitignore 됨).

## 아키텍처

데이터 흐름은 단방향 파이프라인이다:

```
client.fetch_seoul_current_weather()   # httpx GET (5s 타임아웃)
  → CurrentWeather.from_open_meteo()   # dict → frozen dataclass
  → renderer.render_html(weather)      # 단일 HTML 문자열
  → renderer.render_to_file(weather, path)
```

각 단계는 독립적으로 테스트 가능하도록 분리되어 있다 (`tests/test_client.py` 는 `httpx.MockTransport` 로 네트워크 차단, `tests/test_renderer.py` 는 고정 `CurrentWeather` 인스턴스 사용).

### WMO weather_code 매핑 (`formatter.py`)

Open-Meteo 가 반환하는 WMO 코드(0~99)는 세 가지 매핑을 통과한다:
- `weather_code_to_korean(code)` - 한국어 상태어 ("맑음", "대체로 맑음", "비" 등)
- `weather_code_to_emoji(code)` - 카드 메인 이모지 (☀️, ⛅, 🌧️ 등)
- `weather_code_to_theme(code)` - 배경 그라데이션 버킷 `sunny/cloudy/rain/snow/storm` (5단계)

새 코드를 처리하려면 세 dict 모두 갱신할 것. `_to_theme` 의 분기 순서가 fall-through 로직이라 변경 시 주의.

### HTML 렌더링 (`renderer.py`)

- 단일 페이지, 외부 리소스 의존 없음 - CDN/이미지/JS 절대 추가 금지 (자기완결성이 핵심 요건)
- 사용자 입력은 `html.escape()` 로 escape - 새 필드 추가 시 동일하게 처리
- 테마별 그라데이션과 텍스트 색상은 `_THEME_GRADIENTS`, `_THEME_TEXT_COLOR` dict 에서 관리
- 폰트 스택: `"Pretendard", "Noto Sans KR", system-ui` - 로컬 설치 안 되어 있어도 적절히 폴백

### CLI (`cli.py`)

- Windows cp949 콘솔에서 한글이 깨지는 문제 때문에 `main()` 진입 시 `sys.stdout/stderr.reconfigure(encoding="utf-8")` 호출. 출력 메시지 추가 시 이 동작에 의존해도 됨.
- 네트워크/파싱 오류는 모두 `WeatherFetchError` 로 통일되어 stderr + exit code 1 로 흘러간다.

## 코드 컨벤션

- 모든 새 모듈 상단에 `from __future__ import annotations` 사용 (PEP 604 union 문법을 런타임 평가 없이 쓰기 위함)
- `requires-python = ">=3.13"` - 3.13+ 문법 사용 가능
- ruff `line-length = 100`

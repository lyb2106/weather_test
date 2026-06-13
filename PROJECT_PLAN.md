# 서울 날씨 조회 프로젝트 기획

> 이 문서는 후속 에이전트가 읽고 그대로 구현에 착수할 수 있도록 작성된 산출물이다.

## Context

`weather-app` 은 서울의 현재 날씨(기온, 습도, 풍속, 상태)를 외부 API에서 받아와 **시각적으로 보기 좋은 HTML 파일**로 출력하는 도구다. CLI 로 실행하면 HTML 파일을 생성하고 결과 경로를 안내하며, 옵션으로 브라우저에서 바로 열어볼 수 있다.

## 확정 사항

- **외부 API**: Open-Meteo (https://api.open-meteo.com) - API 키 불필요, 무료
- **사용 형태**: CLI 스크립트 (`uv run weather`)
- **조회 범위**: 현재 날씨만 (기온, 습도, 풍속, 상태)
- **출력 형식**: HTML 파일 (시각화된 카드 UI). 콘솔에는 생성 경로만 안내.

## 기술 스택

| 항목 | 선택 |
|------|------|
| 패키지 매니저 | uv |
| Python | 3.11+ |
| HTTP 클라이언트 | `httpx` |
| CLI 인자 처리 | 표준 라이브러리 `argparse` |
| HTML 렌더링 | f-string + `html.escape` (인라인 CSS) |
| 테스트 | `pytest` (dev) |
| 린터 | `ruff` (dev) |

## 프로젝트 구조

```
weather-app/
├─ pyproject.toml
├─ .python-version
├─ README.md
├─ PROJECT_PLAN.md
├─ src/
│  └─ weather_app/
│     ├─ __init__.py
│     ├─ __main__.py
│     ├─ cli.py
│     ├─ client.py
│     ├─ models.py
│     ├─ formatter.py
│     └─ renderer.py
├─ output/
│  └─ seoul-weather.html
└─ tests/
   ├─ test_formatter.py
   ├─ test_renderer.py
   └─ test_client.py
```

`pyproject.toml` 의 `[project.scripts]` 에 `weather = "weather_app.cli:main"` 등록 → `uv run weather` 로 호출.

## Open-Meteo API

- 엔드포인트: `GET https://api.open-meteo.com/v1/forecast`
- 서울 좌표: `latitude=37.5665`, `longitude=126.9780`
- 파라미터:
  - `current=temperature_2m,relative_humidity_2m,wind_speed_10m,weather_code`
  - `timezone=Asia/Seoul`
  - `wind_speed_unit=ms`
- 응답 `current`: `temperature_2m`, `relative_humidity_2m`, `wind_speed_10m`, `weather_code`, `time`

`weather_code` 는 WMO 표준 (0~99). 0(맑음), 1~3(대체로 맑음/구름조금/흐림), 45/48(안개), 51~67(이슬비/비), 71~77(눈), 80~82(소나기), 95~99(뇌우).

## 모듈별 책임

- **`client.py`**: `fetch_seoul_current_weather() -> dict` - httpx 동기 GET, 5초 타임아웃, 비정상 응답 시 `WeatherFetchError`.
- **`models.py`**: `CurrentWeather` dataclass (`temperature_c`, `humidity_pct`, `wind_ms`, `weather_code`, `observed_at`). dict → dataclass 변환 함수 포함.
- **`formatter.py`**: `weather_code_to_korean(code: int) -> str` 와 `weather_code_to_emoji(code: int) -> str`. WMO 0~99 커버.
- **`renderer.py`**: `render_html(w: CurrentWeather) -> str` + `render_to_file(w, path)`. CSS 인라인.
- **`cli.py`**: argparse - `--output PATH` (기본 `output/seoul-weather.html`), `--open` (브라우저 자동 열기). 네트워크 오류 시 한국어 메시지 + exit 1.

## HTML 시각화 설계

**레이아웃**: 가운데 정렬된 단일 카드(약 480px 폭). 상하 그라데이션 배경. 한 화면 수렴.

**구성**:
1. 상단 헤더: "서울" + 관측 시각 (`2026-05-23 14:00 KST`)
2. 메인: 큰 날씨 이모지(72~96px) + 상태어
3. 기온 강조: 큰 숫자 `18.3°C`
4. 하위 카드 2개(가로): 💧 습도 `65%` / 💨 풍속 `2.4 m/s`
5. 푸터: "출처: Open-Meteo"

**스타일**:
- 폰트: `"Pretendard", "Noto Sans KR", system-ui, sans-serif`
- 카드: `rgba(255,255,255,0.85)`, `border-radius: 24px`, 부드러운 그림자
- 날씨 상태별 배경 그라데이션 5단계 (맑음/흐림/비/눈/뇌우)
- 외부 리소스 의존 없음 (자기완결 HTML)
- 반응형: `max-width: 92vw`

**기본 산출물 경로**: `weather-app/output/seoul-weather.html`

## 구현 순서

1. `uv init --package` 로 src 레이아웃 초기화
2. `uv add httpx`, `uv add --dev pytest ruff`
3. `pyproject.toml` 에 `[project.scripts]` 등록
4. `models.py` → `formatter.py` → `client.py` → `renderer.py` → `cli.py` 순서로 구현
5. `tests/` 작성
6. `output/` 을 `.gitignore` 에 추가 (선택)

## 검증

- `uv run weather` → HTML 생성, 콘솔에 경로 출력
- 브라우저로 열어 카드 UI 확인
- `uv run weather --open` → 자동 브라우저 오픈
- 네트워크 차단 시 친절한 에러 + exit 1
- `uv run pytest` 통과
- `uv run ruff check .` 통과

## 비범위

- 시간별/일별 예보, 미세먼지, 다른 도시
- 텍스트/JSON 출력 모드
- 다크 모드 토글, 인터랙션(JS)
- 캐싱/스케줄링/배포

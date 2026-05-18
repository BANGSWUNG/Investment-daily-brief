# 📋 투자 데일리 브리프 (Investment Daily Brief)

**매크로 촉매 모멘텀 트레이딩**을 위한 자동화된 일일 보고서.
매일 아침 이메일 하나로 "어디가 뜨거운지"와 "왜 뜨거운지"를 동시에 확인합니다.

## 보고서 구성

```
📋 투자 데일리 브리프
├── PART 1: 🔥 시장 열기 (Market Heat)
│   ├── 카테고리별 최열 종목 (국가/섹터/크립토/원자재)
│   └── 카테고리별 전체 순위 (열기점수, RVOL, ATR%, OBV 5일%)
│
├── PART 2: 📰 뉴스 키워드 (GDELT)
│   ├── 한국/신흥국 관련 글로벌 뉴스 키워드
│   ├── 매크로 키워드 보도량 + 감성 점수
│   └── 전일 대비 급등 키워드
│
└── PART 3: 🔗 오늘의 체크리스트
    └── 열기 확인 → 촉매 확인 → 차트 진입 → 진입 판단
```

**이메일 본문**: 통합 마크다운 (빠른 확인용)
**첨부 파일**: Excel (상세 표 + 차트, 추이 데이터)

## 실행 시간

| 구분 | UTC | KST | 미장 마감 후 |
|:---:|:---:|:---:|:---:|
| **서머타임** (3~11월) | 00:00 | 09:00 | +4시간 |
| **겨울** (11~3월) | 00:00 | 09:00 | +3시간 |

기상(08:30) 후 30분, 업무 시작(09:30) 전 확인 가능.

## 비용

**$0** — GDELT API 무료, yfinance 무료, GitHub Actions 무료 범위 내.

---

## 셋업 가이드

### 사전 준비

1. **Gmail 앱 비밀번호 발급** (2분)
   - https://myaccount.google.com/apppasswords 접속
   - 앱 이름: `Daily Brief` 입력
   - 생성된 16자리 비밀번호 복사 (이 비밀번호는 한 번만 표시됨)

### 1단계: 저장소 생성

1. GitHub.com → 우측 상단 **+** → **New repository**
2. 이름: `investment-daily-brief` (비공개 추천)
3. **Create repository** 클릭

### 2단계: 파일 업로드

이 폴더의 파일들을 전부 업로드합니다:

```
investment-daily-brief/
├── market_heat_report.py     ← 시장 열기 트래커
├── gdelt_daily_report.py     ← GDELT 뉴스 (기존 파이프라인에서 복사)
├── unified_report.py         ← 통합 래퍼
├── send_email.py             ← 이메일 발송
└── .github/
    └── workflows/
        └── daily-brief.yml   ← GitHub Actions 자동화
```

> ⚠️ `.github` 폴더는 드래그 앤 드롭이 안 될 수 있습니다.
> **Add file** → **Create new file** → 파일 이름에
> `.github/workflows/daily-brief.yml` 입력 → 내용 붙여넣기

> 💡 기존 GDELT 저장소가 별도로 있다면, `gdelt_daily_report.py`를
> 이 저장소로 복사해 넣으세요. 두 저장소를 하나로 합치는 거예요.

### 3단계: Actions 권한 설정

1. 저장소 → **Settings** → 좌측 **Actions** → **General**
2. "Workflow permissions" → **Read and write permissions** 선택
3. **Save** 클릭

### 4단계: 이메일 설정

1. 저장소 → **Settings** → **Secrets and variables** → **Actions**
2. **New repository secret** 으로 3개 추가:

| Secret 이름 | 값 |
|:---|:---|
| `EMAIL_USERNAME` | Gmail 주소 (예: myname@gmail.com) |
| `EMAIL_PASSWORD` | 1단계에서 발급한 앱 비밀번호 (16자리) |
| `EMAIL_TO` | 받을 이메일 주소 (같은 Gmail이어도 OK) |

### 5단계: 수동 테스트

1. 저장소 → **Actions** 탭
2. 좌측에 **📋 투자 데일리 브리프** 클릭
3. 우측 **Run workflow** → **Run workflow** 클릭
4. 2~3분 후 이메일 수신 확인

성공하면 매일 아침 KST 09:00에 자동으로 이메일이 옵니다.

---

## 커스터마이징

### 추적 종목 추가/변경

`market_heat_report.py` 상단의 딕셔너리를 수정:

```python
# 한국 개별 종목 추가 예시
KOREA_STOCKS = {
    "삼성전자": "005930.KS",
    "SK하이닉스": "000660.KS",
}
```

### 실행 시간 변경

`.github/workflows/daily-brief.yml`의 cron 수정:
```yaml
# 현재: UTC 00:00 (KST 09:00)
- cron: "0 0 * * 1-5"

# 예: UTC 22:00 (KST 07:00) — 더 일찍 받고 싶을 때
- cron: "0 22 * * 0-4"
```

### GDELT 키워드 변경

`gdelt_daily_report.py`에서 매크로 키워드 목록을 수정하면 됩니다.

---

## 파일 설명

| 파일 | 역할 |
|:---|:---|
| `market_heat_report.py` | yfinance → 28개 ETF/크립토 데이터 수집 → ATR, RVOL, OBV, 열기점수 계산 → Excel + 마크다운 |
| `gdelt_daily_report.py` | GDELT API → 글로벌 뉴스 키워드 집계 + 감성 분석 → 마크다운 |
| `unified_report.py` | 위 두 스크립트를 순차 실행 → 통합 마크다운 생성 |
| `send_email.py` | 마크다운 → HTML 변환 → Gmail SMTP로 발송 (Excel 첨부) |

## FAQ

**Q: GDELT 파이프라인이 아직 없는데?**
A: `gdelt_daily_report.py` 없이도 실행됩니다. 시장 열기 파트만 포함된 보고서가 생성됩니다.

**Q: 이메일 대신 텔레그램/슬랙으로 받을 수 있나요?**
A: `send_email.py`를 텔레그램/슬랙 봇으로 대체하면 됩니다. 마크다운 포맷이라 대부분의 메신저에서 잘 렌더링됩니다.

**Q: GitHub Actions 무료 한도는?**
A: 월 2,000분. 이 파이프라인은 1회 3~5분이므로, 매일 실행해도 월 ~100분만 소모됩니다.

#!/usr/bin/env python3
"""
GDELT 일일 키워드 집계 파이프라인 v5
- 매크로 키워드 우선 실행 (rate limit 회피)
- timelinevolraw 모드로 실제 기사 수 확보 (250 상한 해제)
- 신흥국: 따옴표 제거 + sourcecountry 필터 병용
- 한국: sourcelang:korean OR english 유지

사용법:
  python gdelt_daily_report.py                    # 오늘 리포트
  python gdelt_daily_report.py --date 2026-05-12  # 특정 날짜
  python gdelt_daily_report.py --output ./reports  # 출력 폴더
  python gdelt_daily_report.py --no-macro          # 매크로 스캔 생략
"""

import requests
import json
import re
import argparse
from datetime import datetime, timedelta
from collections import Counter
from pathlib import Path
import time
import urllib.parse

# ============================================================
# 설정
# ============================================================

GDELT_API_BASE = "https://api.gdeltproject.org/api/v2/doc/doc"

MACRO_KEYWORDS = [
    "Federal Reserve", "interest rate", "inflation", "tariff",
    "oil price", "GDP", "unemployment", "bond yield",
    "recession", "trade war", "sanctions", "currency",
    "central bank", "rate cut", "rate hike", "CPI",
    "semiconductor", "AI chip", "export control"
]

# 한국 검색어 (고유명사 → 따옴표 구문 검색)
KOREA_SEARCH_TERMS = [
    "South Korea", "Korea economy", "Samsung", "KOSPI",
    "Hyundai", "SK Hynix", "Korean won", "Bank of Korea",
    "Korea trade", "Korea semiconductor", "Korea export",
    "LG Energy", "Korea policy", "Korea technology"
]

# 글로벌 검색어 (따옴표 구문 검색)
BROAD_SEARCH_TERMS = [
    "stock market", "central bank", "interest rate",
    "oil price", "trade deal", "bond market",
    "tech stocks", "earnings report", "GDP growth",
    "Federal Reserve", "European Central Bank",
    "currency market", "commodity prices", "supply chain"
]

# 신흥국 검색어 — 키워드+국가코드 조합 (따옴표 없이 AND 검색)
# (keyword, sourcecountry FIPS code)
EMERGING_QUERIES = [
    ("economy", "IN"),   # India
    ("economy", "BR"),   # Brazil
    ("economy", "ID"),   # Indonesia
    ("economy", "VM"),   # Vietnam
    ("economy", "MX"),   # Mexico
    ("economy", "TU"),   # Turkey
    ("economy", "TH"),   # Thailand
    ("economy", "SF"),   # South Africa
    ("economy", "PL"),   # Poland
    ("trade", "IN"),
    ("trade", "BR"),
    ("trade", "ID"),
    ("semiconductor", "IN"),
    ("oil", "BR"),
    ("manufacturing", "VM"),
    ("manufacturing", "MX"),
]

STOPWORDS = {
    "the","a","an","and","or","but","in","on","at","to","for","of","with",
    "by","from","is","are","was","were","be","been","has","have","had","do",
    "does","did","will","would","could","should","may","might","can","shall",
    "not","no","nor","so","if","then","than","that","this","these","those",
    "it","its","as","up","out","about","into","over","after","before",
    "between","under","again","further","once","here","there","when","where",
    "why","how","all","each","every","both","few","more","most","other",
    "some","such","only","own","same","also","just","very","new","said",
    "says","say","news","report","reports","according","year","years","day",
    "days","time","first","last","week","month","per","get","set","got",
    "one","two","three","many","much","make","made","like","well","back",
    "even","still","way","take","come","good","know","think","see","look",
    "want","give","use","find","tell","ask","work","seem","feel","try",
    "leave","call","need","become","keep","let","begin","show","hear","play",
    "run","move","live","believe","bring","happen","write","provide","sit",
    "stand","lose","pay","meet","include","continue","going","being","having",
    "what","which","who","whom","whose","during","while","through",
    "s","t","re","ve","ll","d","m",
    "mr","ms","mrs","dr","st","vs","etc","us",
    "told","added","he","she","they","we","his","her",
    "monday","tuesday","wednesday","thursday","friday",
    "saturday","sunday","today","yesterday","tomorrow",
    "jan","feb","mar","apr","may","jun","jul","aug",
    "sep","oct","nov","dec","january","february","march",
    "april","june","july","august","september","october",
    "november","december","2024","2025","2026","2027",
    "update","updated","breaking","exclusive","watch",
    "video","photo","photos","image","images","click",
    "read","full","story","article","opinion","editorial",
    "analysis","world","global","international","national",
    "local","people","government","country","state","city",
    "million","billion","trillion","percent","number",
    "part","group","company","market","stock","share",
    "price","high","low","data","plan","amid","among",
    "off","gets","Inc","Corp","Ltd","Co",
    # 스페인어
    "de","la","el","en","los","las","del","al","es","se","un","una","que",
    "por","con","para","su","mas","como","pero","sus","le","ya","ha","fue",
    "son","ser","entre","desde","sobre","sin","hasta","hay","donde","muy",
    # 프랑스어
    "le","les","des","du","au","aux","ce","ces","est","et","ou","ne","pas",
    "une","dans","sur","qui","sont","avec","pour","plus","par","tout","fait",
    # 포르투갈어
    "da","do","dos","das","na","no","nas","ao","aos","uma","com","nao",
    "mais","foi","tem","pode","seu","sua","seus","suas","isso","pelo","pela",
    # 독일어
    "der","die","den","dem","und","ist","von","zu","mit","sich","auf",
    "nicht","ein","eine","auch","als","nach","bei","aus","wird","oder",
    # 러시아어
    "на","за","из","по","от","до","об","не","но","да","что","как","это",
    "все","его","при","так","уже","для","бы","же","ли","он","она","они",
    # 터키어/폴란드어/인도네시아어
    "bir","bu","ve","ile","olan","dan","olarak","gibi","daha","var",
    "nie","ze","jest","jak","ale","tak","czy",
    "dan","yang","ini","itu","di","ke","dari","ada","untuk","dengan",
    # 이탈리아어
    "il","lo","gli","che","si","non","sono","ma","anche","nel","nella",
    # 중국어 뉴스사이트
    "中华网","新闻","新浪财经","新浪网","东方财富网","网易","搜狐",
    "凤凰网","腾讯","百度","参考消息","环球网","央视","人民网",
    "的","了","在","是","和","有",
    # 뉴스 소스
    "reuters","associated","press","bloomberg","cnbc","bbc","cnn",
    "fox","nyt","wsj","ap","afp",
    # 검색어 오염 방지
    "economy","trade","investment","export","import","energy",
    "technology","election","policy","growth","crisis","reform",
    "regulation","korea","south","india","brazil","indonesia",
    "vietnam","mexico","turkey","thailand","africa","poland",
    "emerging","frontier","manufacturing","nearshoring","mining",
}


# ============================================================
# GDELT API
# ============================================================

def query_gdelt(query_str, mode="artlist", timespan="24h",
                max_records=250, retry=2):
    """
    GDELT Doc API 호출. query_str은 완성된 쿼리 문자열.
    """
    params = {
        "query": query_str,
        "mode": mode,
        "format": "json",
        "timespan": timespan,
        "maxrecords": max_records,
        "sort": "datedesc"
    }
    url = f"{GDELT_API_BASE}?{urllib.parse.urlencode(params)}"

    for attempt in range(retry + 1):
        try:
            resp = requests.get(url, timeout=12)
            resp.raise_for_status()
            data = resp.json()
            return data
        except requests.exceptions.Timeout:
            if attempt < retry:
                time.sleep(1)
                continue
            return None
        except (requests.exceptions.RequestException, json.JSONDecodeError):
            return None
    return None


def search_articles(keyword, sourcelang=None, sourcecountry=None,
                    phrase=True, timespan="24h", max_records=250):
    """
    기사 검색 래퍼.
    phrase=True: 따옴표 구문 검색 ("Federal Reserve")
    phrase=False: AND 검색 (Federal Reserve → 두 단어 모두 포함)
    """
    parts = []
    if keyword:
        if phrase:
            parts.append(f'"{keyword}"')
        else:
            parts.append(keyword)
    if sourcelang:
        if " OR " in sourcelang:
            lang_parts = [f"sourcelang:{l.strip()}" for l in sourcelang.split(" OR ")]
            parts.append(f"({' OR '.join(lang_parts)})")
        else:
            parts.append(f"sourcelang:{sourcelang}")
    if sourcecountry:
        parts.append(f"sourcecountry:{sourcecountry}")

    query_str = " ".join(parts)
    data = query_gdelt(query_str, mode="artlist", timespan=timespan,
                       max_records=max_records)
    if data:
        return data.get("articles", [])
    return []


def get_article_volume(keyword, timespan="24h"):
    """timelinevolraw 모드로 실제 기사 수 조회 (250 상한 없음)"""
    query_str = f'"{keyword}"'
    data = query_gdelt(query_str, mode="timelinevolraw", timespan=timespan)
    if data and "timeline" in data:
        timeline = data["timeline"]
        if timeline and timeline[0].get("data"):
            total = sum(d.get("value", 0) for d in timeline[0]["data"])
            return int(total)
    return 0


def get_tone(keyword, timespan="24h"):
    """톤(감성) 점수 조회"""
    query_str = f'"{keyword}"'
    data = query_gdelt(query_str, mode="timelinetone", timespan=timespan)
    if data and "timeline" in data:
        timeline = data["timeline"]
        if timeline and timeline[0].get("data"):
            tones = [d["value"] for d in timeline[0]["data"] if d.get("value")]
            if tones:
                return round(sum(tones) / len(tones), 2)
    return None


# ============================================================
# 키워드 추출
# ============================================================

def extract_keywords_from_titles(articles):
    """기사 제목에서 키워드 빈도 집계"""
    word_counter = Counter()
    bigram_counter = Counter()

    for article in articles:
        title = article.get("title", "")
        if not title:
            continue
        cleaned = re.sub(r'[^\w\s\'-]', ' ', title)
        words = cleaned.split()

        filtered = []
        for w in words:
            w_lower = w.lower().strip("'-")
            if (len(w_lower) > 1
                and not w_lower.isdigit()
                and w_lower not in STOPWORDS):
                filtered.append(w)

        for w in filtered:
            if w[0].isupper() and len(w) > 1:
                word_counter[w] += 1
            else:
                word_counter[w.lower()] += 1

        for i in range(len(filtered) - 1):
            bigram = f"{filtered[i]} {filtered[i+1]}"
            bigram_counter[bigram] += 1

    return word_counter, bigram_counter


def compute_keyword_changes(today_counts, yesterday_counts):
    changes = {}
    all_kw = set(list(today_counts.keys())[:50]) | set(list(yesterday_counts.keys())[:50])
    for kw in all_kw:
        tv = today_counts.get(kw, 0)
        yv = yesterday_counts.get(kw, 0)
        if yv == 0 and tv > 0:
            changes[kw] = {"today": tv, "yesterday": 0, "change": "NEW"}
        elif yv > 0:
            pct = round((tv - yv) / yv * 100, 1)
            changes[kw] = {"today": tv, "yesterday": yv, "change": pct}
    return changes


# ============================================================
# 데이터 수집
# ============================================================

def collect_articles_phrase(search_terms, sourcelang=None, timespan="24h",
                           max_records=250, label=""):
    """구문 검색 (따옴표) 기반 수집"""
    all_articles = []
    for term in search_terms:
        batch = search_articles(term, sourcelang=sourcelang,
                               phrase=True, timespan=timespan,
                               max_records=max_records)
        all_articles.extend(batch)
        if batch:
            print(f"  '{term}' -> {len(batch)}건")
        time.sleep(0.2)
    return _deduplicate(all_articles, label)


def collect_articles_country(queries, sourcelang=None, timespan="24h",
                             max_records=250, label=""):
    """키워드 + sourcecountry 필터 기반 수집 (따옴표 없음)"""
    all_articles = []
    for keyword, country_code in queries:
        batch = search_articles(keyword, sourcelang=sourcelang,
                               sourcecountry=country_code,
                               phrase=False, timespan=timespan,
                               max_records=max_records)
        all_articles.extend(batch)
        if batch:
            print(f"  '{keyword}' ({country_code}) -> {len(batch)}건")
        time.sleep(0.2)
    return _deduplicate(all_articles, label)


def _deduplicate(articles, label):
    seen = set()
    unique = []
    for a in articles:
        url = a.get("url", "")
        if url and url not in seen:
            seen.add(url)
            unique.append(a)
    print(f"  => {label} 총 {len(unique)}건 (중복 제거)")
    return unique


def check_macro_keywords(timespan="24h"):
    """매크로 키워드별 보도량(timelinevolraw) + 톤"""
    results = []
    for kw in MACRO_KEYWORDS:
        # timelinevolraw로 실제 기사 수 조회 (250 상한 없음)
        volume = get_article_volume(kw, timespan=timespan)
        if volume > 0:
            tone = get_tone(kw, timespan=timespan)
            results.append({"keyword": kw, "count": volume, "tone": tone})
            print(f"  '{kw}' -> {volume}건, 톤={tone}")
        time.sleep(0.2)
    results.sort(key=lambda x: x["count"], reverse=True)
    return results


# ============================================================
# 마크다운 리포트
# ============================================================

def generate_markdown_report(date_str, korea_data, emerging_data,
                              global_data, macro_data, changes=None):
    lines = []

    lines.append("---")
    lines.append(f"date: {date_str}")
    lines.append("type: gdelt-daily")
    lines.append(f"created: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("tags:")
    lines.append("  - gdelt")
    lines.append("  - daily-report")
    lines.append("  - catalyst-screening")
    lines.append("---")
    lines.append("")
    lines.append(f"# GDELT 일일 키워드 리포트 — {date_str}")
    lines.append("")

    # 1. 매크로
    lines.append("## 1. 매크로 키워드 보도량")
    lines.append("")
    if macro_data:
        lines.append("| 키워드 | 기사 수 | 톤 (감성) | 신호 |")
        lines.append("|--------|---------|-----------|------|")
        for item in macro_data[:15]:
            tone = item["tone"] if item["tone"] is not None else "N/A"
            count = f"{item['count']:,}"
            sig = ""
            if isinstance(tone, (int, float)):
                if tone < -3: sig = "🔴 강한 부정"
                elif tone < -1: sig = "🟠 부정"
                elif tone > 3: sig = "🟢 강한 긍정"
                elif tone > 1: sig = "🟡 긍정"
                else: sig = "⚪ 중립"
            lines.append(f"| {item['keyword']} | {count} | {tone} | {sig} |")
    else:
        lines.append("*매크로 데이터 수집 실패*")
    lines.append("")

    # 2. 한국
    lines.append("## 2. 한국 관련 뉴스 키워드 Top 20")
    lines.append("")
    kr_words, kr_bigrams = korea_data
    if kr_words:
        lines.append("### 단일 키워드")
        lines.append("| 순위 | 키워드 | 빈도 |")
        lines.append("|------|--------|------|")
        for i, (w, c) in enumerate(kr_words.most_common(20), 1):
            lines.append(f"| {i} | {w} | {c} |")
        lines.append("")
        top_bi = [(b, c) for b, c in kr_bigrams.most_common(10) if c >= 2]
        if top_bi:
            lines.append("### 2단어 조합 (바이그램)")
            lines.append("| 순위 | 키워드 조합 | 빈도 |")
            lines.append("|------|------------|------|")
            for i, (b, c) in enumerate(top_bi, 1):
                lines.append(f"| {i} | {b} | {c} |")
    else:
        lines.append("*한국 뉴스 데이터 없음*")
    lines.append("")

    # 3. 신흥국
    lines.append("## 3. 신흥국 뉴스 키워드 Top 20")
    lines.append("")
    em_words, _ = emerging_data
    if em_words:
        lines.append("| 순위 | 키워드 | 빈도 |")
        lines.append("|------|--------|------|")
        for i, (w, c) in enumerate(em_words.most_common(20), 1):
            lines.append(f"| {i} | {w} | {c} |")
    else:
        lines.append("*신흥국 뉴스 데이터 없음*")
    lines.append("")

    # 4. 글로벌
    lines.append("## 4. 글로벌 뉴스 키워드 Top 20")
    lines.append("")
    gl_words, _ = global_data
    if gl_words:
        lines.append("| 순위 | 키워드 | 빈도 |")
        lines.append("|------|--------|------|")
        for i, (w, c) in enumerate(gl_words.most_common(20), 1):
            lines.append(f"| {i} | {w} | {c} |")
    else:
        lines.append("*글로벌 뉴스 데이터 없음*")
    lines.append("")

    # 5. 전일 대비
    if changes:
        lines.append("## 5. 전일 대비 급등 키워드")
        lines.append("")
        new_kws = [(k,v) for k,v in changes.items()
                   if v["change"] == "NEW" and v["today"] >= 3]
        rising = [(k,v) for k,v in changes.items()
                  if isinstance(v["change"], (int,float))
                  and v["change"] > 50 and v["today"] >= 3]
        rising.sort(key=lambda x: x[1]["change"], reverse=True)
        if new_kws or rising:
            lines.append("| 키워드 | 오늘 | 어제 | 변화 |")
            lines.append("|--------|------|------|------|")
            for kw, v in new_kws[:10]:
                lines.append(f"| **{kw}** | {v['today']} | 0 | 🆕 신규 |")
            for kw, v in rising[:10]:
                lines.append(f"| **{kw}** | {v['today']} | {v['yesterday']} | ↑ {v['change']}% |")
        else:
            lines.append("*전일 대비 유의미한 급등 키워드 없음*")
        lines.append("")

    # 6. Claude 프롬프트
    lines.append("## 6. Claude 토론용 프롬프트")
    lines.append("")
    top_kw = []
    if kr_words:
        top_kw.extend([w for w, _ in kr_words.most_common(5)])
    if macro_data:
        top_kw.extend([m["keyword"] for m in macro_data[:3]])
    kw_str = ", ".join(top_kw[:8]) if top_kw else "(키워드 없음)"

    lines.append("```")
    lines.append(f"오늘({date_str}) GDELT 리포트에서 고빈도 키워드는: {kw_str}")
    lines.append("")
    lines.append("1. 이 키워드들 사이의 연관성과, 시장이 주목하는 내러티브가 뭔지 분석해줘.")
    lines.append("2. 이 중 스윙 트레이딩(3일~2주) 촉매로 작용할 수 있는 건 뭐야?")
    lines.append("3. 코스톨라니 달걀 모형 기준으로 관련 섹터가 어디쯤인지 의견을 줘.")
    lines.append("```")
    lines.append("")
    lines.append("---")
    lines.append(f"*Generated by GDELT Daily Pipeline v5 | {datetime.now().strftime('%H:%M KST')}*")

    return "\n".join(lines)


# ============================================================
# 메인
# ============================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", type=str, default=None)
    parser.add_argument("--output", type=str, default="./reports")
    parser.add_argument("--no-macro", action="store_true")
    args = parser.parse_args()

    date_str = args.date if args.date else datetime.now().strftime("%Y-%m-%d")
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"{'='*60}")
    print(f"  GDELT 일일 키워드 리포트 v5 — {date_str}")
    print(f"{'='*60}")

    # ── 1. 매크로 (가장 먼저! rate limit 전에) ──
    macro_data = []
    if not args.no_macro:
        print("\n[1/5] 매크로 키워드 보도량 스캔 중...")
        macro_data = check_macro_keywords(timespan="24h")
        print(f"  => {len(macro_data)}개 키워드 감지")
    else:
        print("\n[1/5] 매크로 스캔 건너뜀")

    # ── 2. 한국 (한국어 + 영어, 구문 검색) ──
    print("\n[2/5] 한국 관련 뉴스 수집 중...")
    kr_articles = collect_articles_phrase(
        KOREA_SEARCH_TERMS,
        sourcelang="korean OR english",
        label="한국"
    )
    kr_keywords = extract_keywords_from_titles(kr_articles)

    # ── 3. 신흥국 (영어, 키워드+국가코드 AND 검색) ──
    print("\n[3/5] 신흥국 뉴스 수집 중...")
    em_articles = collect_articles_country(
        EMERGING_QUERIES,
        sourcelang="english",
        label="신흥국"
    )
    em_keywords = extract_keywords_from_titles(em_articles)

    # ── 4. 글로벌 (영어, 구문 검색) ──
    print("\n[4/5] 글로벌 뉴스 수집 중...")
    gl_articles = collect_articles_phrase(
        BROAD_SEARCH_TERMS,
        sourcelang="english",
        label="글로벌"
    )
    gl_keywords = extract_keywords_from_titles(gl_articles)

    # ── 5. 전일 대비 ──
    print("\n[5/5] 전일 대비 키워드 변화 계산 중...")
    changes = None
    yd_articles = collect_articles_phrase(
        BROAD_SEARCH_TERMS[:6],
        sourcelang="english",
        timespan="48h", max_records=100,
        label="비교용"
    )
    if yd_articles:
        yd_words, _ = extract_keywords_from_titles(yd_articles)
        gl_words_today, _ = gl_keywords
        changes = compute_keyword_changes(gl_words_today, yd_words)

    # ── 리포트 생성 ──
    print("\n[생성] 마크다운 리포트 작성 중...")
    report = generate_markdown_report(
        date_str=date_str,
        korea_data=kr_keywords,
        emerging_data=em_keywords,
        global_data=gl_keywords,
        macro_data=macro_data,
        changes=changes
    )

    report_path = output_dir / f"{date_str}.md"
    report_path.write_text(report, encoding="utf-8")
    print(f"\n✅ 리포트 저장 완료: {report_path}")

    # 콘솔 요약
    print(f"\n{'='*60}")
    print("  빠른 요약")
    print(f"{'='*60}")
    kr_words, _ = kr_keywords
    gl_words, _ = gl_keywords
    if kr_words:
        top5 = ", ".join([f"{w}({c})" for w, c in kr_words.most_common(5)])
        print(f"  🇰🇷 한국 Top 5: {top5}")
    if gl_words:
        top5 = ", ".join([f"{w}({c})" for w, c in gl_words.most_common(5)])
        print(f"  🌍 글로벌 Top 5: {top5}")
    if macro_data:
        top3 = ", ".join([f"{m['keyword']}({m['count']:,}건)" for m in macro_data[:3]])
        print(f"  📊 매크로 Top 3: {top3}")
    print(f"\n📄 전체 리포트: {report_path}")

    return str(report_path)


if __name__ == "__main__":
    main()

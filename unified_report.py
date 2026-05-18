#!/usr/bin/env python3
"""
투자 데일리 브리프 — 통합 보고서 생성기
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GDELT 뉴스 키워드 + 시장 열기 트래커 → 통합 마크다운 + Excel 보고서

실행 순서:
  1. market_heat_report.py → 시장 열기 데이터 수집 + Excel 생성
  2. gdelt_daily_report.py → GDELT 뉴스 키워드 수집 + 마크다운 생성
  3. 두 결과를 통합 마크다운으로 합침
  4. send_email.py → 이메일 발송 (본문: 통합 마크다운, 첨부: Excel)
"""

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path


def run_script(script_name: str, output_dir: str) -> bool:
    """개별 스크립트 실행"""
    script_path = Path(__file__).parent / script_name
    if not script_path.exists():
        print(f"  ⚠ {script_name} 파일 없음 — 건너뜀")
        return False

    print(f"\n{'='*50}")
    print(f"▶ {script_name} 실행 중...")
    print(f"{'='*50}")

    result = subprocess.run(
        [sys.executable, str(script_path), "--output", output_dir],
        capture_output=False,
    )
    if result.returncode != 0:
        print(f"  ✗ {script_name} 실행 실패 (exit code: {result.returncode})")
        return False

    print(f"  ✓ {script_name} 완료")
    return True


def find_report_file(output_dir: Path, prefix: str, ext: str, date_str: str) -> Path | None:
    """보고서 파일 탐색"""
    # 날짜 포함 파일명 우선
    candidates = [
        output_dir / f"{prefix}{date_str}{ext}",
        output_dir / f"{prefix}_{date_str}{ext}",
        output_dir / f"{date_str}{ext}",
    ]
    for c in candidates:
        if c.exists():
            return c

    # 패턴 매칭 폴백
    matches = sorted(output_dir.glob(f"{prefix}*{ext}"), key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0] if matches else None


def merge_reports(output_dir: Path, date_str: str) -> Path:
    """GDELT 마크다운 + 시장 열기 마크다운 → 통합 마크다운"""

    # ── 시장 열기 마크다운 읽기 ──
    heat_md_path = find_report_file(output_dir, "market_heat_", ".md", date_str)
    heat_content = ""
    if heat_md_path and heat_md_path.exists():
        heat_content = heat_md_path.read_text(encoding="utf-8")
        # 기존 제목 제거 (통합 제목 사용)
        lines = heat_content.split("\n")
        if lines and lines[0].startswith("# "):
            lines = lines[1:]
        heat_content = "\n".join(lines).strip()
        print(f"  ✓ 시장 열기 마크다운 로드: {heat_md_path.name}")
    else:
        heat_content = "> ⚠ 시장 열기 데이터를 가져오지 못했습니다."
        print(f"  ⚠ 시장 열기 마크다운 없음")

    # ── GDELT 마크다운 읽기 ──
    gdelt_md_path = find_report_file(output_dir, "", ".md", date_str)
    # market_heat 파일은 제외
    if gdelt_md_path and "market_heat" in gdelt_md_path.name:
        # 다른 .md 파일 찾기
        all_mds = sorted(output_dir.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)
        gdelt_md_path = next(
            (p for p in all_mds if "market_heat" not in p.name and "unified" not in p.name),
            None
        )

    gdelt_content = ""
    if gdelt_md_path and gdelt_md_path.exists():
        gdelt_content = gdelt_md_path.read_text(encoding="utf-8")
        lines = gdelt_content.split("\n")
        if lines and lines[0].startswith("# "):
            lines = lines[1:]
        gdelt_content = "\n".join(lines).strip()
        print(f"  ✓ GDELT 마크다운 로드: {gdelt_md_path.name}")
    else:
        gdelt_content = "> ⚠ GDELT 뉴스 데이터를 가져오지 못했습니다."
        print(f"  ⚠ GDELT 마크다운 없음")

    # ── 통합 마크다운 조립 ──
    now = datetime.now().strftime("%Y-%m-%d %H:%M KST")

    unified = f"""# 📋 투자 데일리 브리프 — {date_str}

> 자동 생성: {now} | 매크로 촉매 모멘텀 트레이딩 의사결정 지원

---

# 🔥 PART 1 — 시장 열기 (Market Heat)

{heat_content}

---

# 📰 PART 2 — 뉴스 키워드 (GDELT)

{gdelt_content}

---

# 🔗 PART 3 — 오늘의 체크리스트

위 데이터를 기반으로 아래 순서로 진행하세요:

1. **열기 확인**: 카테고리별 최열 종목을 확인하고, 매크로 방향성과 일치하는지 판단
2. **촉매 확인**: GDELT 급등 키워드가 어떤 시장/섹터와 연결되는지 교차 확인
3. **차트 진입**: 열기 + 촉매가 겹치는 시장의 개별 종목 차트를 TradingView에서 확인
4. **진입 판단**: 볼린저밴드 + EMA 20/50 + RSI 14 + OBV 인디케이터 스택으로 최종 진입점 결정

> ※ 상세 데이터는 첨부된 Excel 파일의 각 시트에서 확인하세요.

---
*투자 데일리 브리프 | GDELT(무료) + yfinance(무료) | GitHub Actions 자동 생성*
"""

    # ── 저장 ──
    unified_path = output_dir / f"daily_brief_{date_str}.md"
    unified_path.write_text(unified.strip(), encoding="utf-8")
    print(f"\n✓ 통합 보고서 생성: {unified_path}")

    return unified_path


def main():
    parser = argparse.ArgumentParser(description="투자 데일리 브리프 통합 생성기")
    parser.add_argument("--output", default="./reports", help="보고서 저장 디렉토리")
    args = parser.parse_args()

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y-%m-%d")

    print("╔══════════════════════════════════════════════════╗")
    print(f"║  📋 투자 데일리 브리프 — {date_str}             ║")
    print("╚══════════════════════════════════════════════════╝")

    # 1. 시장 열기 트래커 실행
    heat_ok = run_script("market_heat_report.py", args.output)

    # 2. GDELT 뉴스 리포트 실행
    gdelt_ok = run_script("gdelt_daily_report.py", args.output)

    if not heat_ok and not gdelt_ok:
        print("\n✗ 두 스크립트 모두 실패했습니다.")
        sys.exit(1)

    # 3. 통합 마크다운 생성
    print(f"\n{'='*50}")
    print("▶ 통합 보고서 병합 중...")
    print(f"{'='*50}")
    unified_path = merge_reports(output_dir, date_str)

    # 4. 결과 파일 목록 출력
    print(f"\n{'='*50}")
    print("📁 생성된 파일:")
    print(f"{'='*50}")
    for f in sorted(output_dir.glob(f"*{date_str}*")):
        size = f.stat().st_size
        unit = "KB" if size < 1e6 else "MB"
        size_display = size / 1024 if size < 1e6 else size / 1e6
        print(f"  • {f.name} ({size_display:.1f} {unit})")

    print(f"\n✓ 데일리 브리프 생성 완료!")


if __name__ == "__main__":
    main()

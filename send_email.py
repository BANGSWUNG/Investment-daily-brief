#!/usr/bin/env python3
"""
이메일 발송 스크립트
━━━━━━━━━━━━━━━━━━
통합 마크다운 → HTML 이메일 본문 + Excel 첨부 발송

사용법:
  python send_email.py \
    --md reports/daily_brief_2026-05-14.md \
    --attachments reports/market_heat_2026-05-14.xlsx \
    --to your@email.com \
    --from sender@gmail.com \
    --password "앱비밀번호16자리"

환경변수로도 전달 가능:
  EMAIL_TO, EMAIL_USERNAME, EMAIL_PASSWORD
"""

import argparse
import os
import re
import smtplib
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path


def md_to_html(md_text: str) -> str:
    """마크다운 → 간이 HTML 변환 (외부 라이브러리 불필요)"""
    html = md_text

    # 코드 블록 보호 (변환에서 제외)
    code_blocks = []
    def save_code(m):
        code_blocks.append(m.group(0))
        return f"__CODE_BLOCK_{len(code_blocks)-1}__"
    html = re.sub(r"```.*?```", save_code, html, flags=re.DOTALL)

    # 헤더
    html = re.sub(r"^### (.+)$", r"<h3>\1</h3>", html, flags=re.MULTILINE)
    html = re.sub(r"^## (.+)$", r"<h2>\1</h2>", html, flags=re.MULTILINE)
    html = re.sub(r"^# (.+)$", r"<h1>\1</h1>", html, flags=re.MULTILINE)

    # 굵게, 기울임
    html = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", html)
    html = re.sub(r"\*(.+?)\*", r"<em>\1</em>", html)

    # 수평선
    html = re.sub(r"^---+$", "<hr>", html, flags=re.MULTILINE)

    # 인용
    html = re.sub(r"^> (.+)$", r'<blockquote style="border-left:3px solid #ccc;padding-left:12px;color:#666;">\1</blockquote>', html, flags=re.MULTILINE)

    # 테이블 변환
    lines = html.split("\n")
    result = []
    in_table = False
    for line in lines:
        stripped = line.strip()
        if "|" in stripped and stripped.startswith("|"):
            cells = [c.strip() for c in stripped.split("|")[1:-1]]
            # 구분선 행 스킵
            if all(set(c) <= set("-: ") for c in cells):
                continue
            if not in_table:
                result.append('<table style="border-collapse:collapse;width:100%;margin:8px 0;">')
                # 첫 행은 헤더
                row_html = "".join(
                    f'<th style="border:1px solid #ddd;padding:6px 10px;background:#2F5496;color:white;font-size:13px;text-align:center;">{c}</th>'
                    for c in cells
                )
                result.append(f"<tr>{row_html}</tr>")
                in_table = True
            else:
                row_html = "".join(
                    f'<td style="border:1px solid #ddd;padding:5px 10px;font-size:13px;text-align:center;">{c}</td>'
                    for c in cells
                )
                result.append(f"<tr>{row_html}</tr>")
        else:
            if in_table:
                result.append("</table>")
                in_table = False
            result.append(line)
    if in_table:
        result.append("</table>")
    html = "\n".join(result)

    # 리스트
    html = re.sub(r"^(\d+)\. (.+)$", r"<li>\2</li>", html, flags=re.MULTILINE)
    html = re.sub(r"^- (.+)$", r"<li>\1</li>", html, flags=re.MULTILINE)

    # 코드 블록 복원
    for i, block in enumerate(code_blocks):
        html = html.replace(f"__CODE_BLOCK_{i}__", f"<pre>{block}</pre>")

    # 줄바꿈
    html = re.sub(r"\n\n+", "<br><br>", html)
    html = html.replace("\n", "\n")

    # 전체 감싸기
    return f"""
    <div style="font-family:'Apple SD Gothic Neo','Malgun Gothic',Arial,sans-serif;
                max-width:800px;margin:0 auto;padding:20px;color:#333;
                line-height:1.6;font-size:14px;">
        {html}
    </div>
    """


def send_email(
    to_addr: str,
    from_addr: str,
    password: str,
    subject: str,
    html_body: str,
    attachments: list[Path] = None,
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 465,
):
    """HTML 이메일 + 첨부 파일 발송"""
    msg = MIMEMultipart("mixed")
    msg["From"] = f"투자 데일리 브리프 <{from_addr}>"
    msg["To"] = to_addr
    msg["Subject"] = subject

    # HTML 본문
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # 첨부 파일
    if attachments:
        for filepath in attachments:
            if not filepath.exists():
                print(f"  ⚠ 첨부 파일 없음: {filepath}")
                continue
            with open(filepath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{filepath.name}"',
            )
            msg.attach(part)
            print(f"  📎 첨부: {filepath.name} ({filepath.stat().st_size / 1024:.0f} KB)")

    # 발송
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(from_addr, password)
        server.send_message(msg)

    print(f"  ✉ 발송 완료 → {to_addr}")


def main():
    parser = argparse.ArgumentParser(description="데일리 브리프 이메일 발송")
    parser.add_argument("--md", required=True, help="통합 마크다운 파일 경로")
    parser.add_argument("--attachments", nargs="*", default=[], help="첨부 파일 경로들")
    parser.add_argument("--to", default=os.environ.get("EMAIL_TO", ""), help="수신자 이메일")
    parser.add_argument("--from", dest="from_addr", default=os.environ.get("EMAIL_USERNAME", ""), help="발신자 이메일")
    parser.add_argument("--password", default=os.environ.get("EMAIL_PASSWORD", ""), help="앱 비밀번호")
    parser.add_argument("--smtp-server", default="smtp.gmail.com")
    parser.add_argument("--smtp-port", type=int, default=465)
    args = parser.parse_args()

    # 유효성 검사
    md_path = Path(args.md)
    if not md_path.exists():
        print(f"✗ 마크다운 파일 없음: {md_path}")
        sys.exit(1)

    if not all([args.to, args.from_addr, args.password]):
        print("✗ 이메일 설정이 부족합니다.")
        print("  필요: --to, --from, --password (또는 환경변수 EMAIL_TO, EMAIL_USERNAME, EMAIL_PASSWORD)")
        sys.exit(1)

    # 마크다운 → HTML
    md_text = md_path.read_text(encoding="utf-8")
    html_body = md_to_html(md_text)

    # 제목
    date_str = datetime.now().strftime("%Y-%m-%d")
    subject = f"📋 투자 데일리 브리프 — {date_str}"

    # 첨부 파일
    attachment_paths = [Path(a) for a in args.attachments]

    print(f"\n📧 이메일 발송 중...")
    send_email(
        to_addr=args.to,
        from_addr=args.from_addr,
        password=args.password,
        subject=subject,
        html_body=html_body,
        attachments=attachment_paths,
        smtp_server=args.smtp_server,
        smtp_port=args.smtp_port,
    )


if __name__ == "__main__":
    import sys
    main()

#!/usr/bin/env python3
"""
한마당 카드뉴스 자동 생성기
Google Sheets에서 당일 뉴스를 크롤링해 카드뉴스 HTML을 자동 생성합니다.

사용법:
  python3 generate_cardnews.py          # 오늘 날짜 뉴스
  python3 generate_cardnews.py 2026-05-07  # 특정 날짜 뉴스
"""

import csv
import io
import re
import os
import sys
import glob
import urllib.request
import urllib.error
import urllib.parse
from datetime import date, datetime, timezone, timedelta

# 크롤러가 수집시각을 MYT(말레이시아)로 저장하므로 대상 날짜도 MYT 기준으로 계산.
# (GitHub Actions 러너는 UTC라 date.today()를 쓰면 cron MYT 06:48 실행 시 전날로 잡힘)
MYT = timezone(timedelta(hours=8))
from html import escape
try:
    import anthropic
except ImportError:
    anthropic = None

# .env 파일에서 환경 변수 로드
def _load_env():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

_load_env()

# ── 설정 ──────────────────────────────────────────────
SHEET_ID   = "18BUYAw1ruBDUEbvxg8AUpm9WOCsvB6iZy1amAzCpgKg"
SHEET_URL  = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
OUTPUT     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "카드뉴스.html")
MAX_CARDS  = 7   # 최대 기사 카드 수 (커버·아웃트로 제외)
TIMEOUT    = 12

FETCH_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}

# 카테고리별 폴백 이미지 (기사 이미지 없을 때)
FALLBACK_IMAGES = {
    "정치": "https://images.unsplash.com/photo-1529107386315-e1a2ed48a620?w=800&q=80",
    "경제": "https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=800&q=80",
    "사회": "https://images.unsplash.com/photo-1529156069898-49953e39b3ac?w=800&q=80",
    "보건": "https://images.unsplash.com/photo-1584036561566-baf8f5f1b144?w=800&q=80",
    "교육": "https://images.unsplash.com/photo-1503676260728-1c00da094a0b?w=800&q=80",
    "범죄": "https://images.unsplash.com/photo-1589829545856-d10d557cf95f?w=800&q=80",
    "외교": "https://images.unsplash.com/photo-1529107386315-e1a2ed48a620?w=800&q=80",
    "환경": "https://images.unsplash.com/photo-1441974231531-c6227db76b6e?w=800&q=80",
    "교통": "https://images.unsplash.com/photo-1519003722824-194d4455a60c?w=800&q=80",
    "기타": "https://images.unsplash.com/photo-1504711434969-e33886168f5c?w=800&q=80",
}
DEFAULT_FALLBACK = "https://images.unsplash.com/photo-1504711434969-e33886168f5c?w=800&q=80"


# ── HTTP 유틸 ─────────────────────────────────────────
def http_get(url, timeout=TIMEOUT):
    req = urllib.request.Request(url, headers=FETCH_HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        final = r.url
        if final != url:
            req2 = urllib.request.Request(final, headers=FETCH_HEADERS)
            with urllib.request.urlopen(req2, timeout=timeout) as r2:
                return r2.read().decode("utf-8", errors="ignore")
        return r.read().decode("utf-8", errors="ignore")


# ── 스프레드시트 ──────────────────────────────────────
def fetch_sheet():
    print("📥 Google Sheets 데이터 가져오는 중...")
    raw = http_get(SHEET_URL)
    reader = csv.reader(io.StringIO(raw))
    rows = list(reader)
    if rows:
        rows = rows[1:]   # 헤더 제거
    print(f"   총 {len(rows)}개 기사 로드됨")
    return rows


def filter_by_date(rows, target_date: str):
    """target_date: 'YYYY-MM-DD' 형식 — 수집시각(G열) 기준으로 필터"""
    result = [r for r in rows if len(r) >= 7 and r[6].startswith(target_date)]
    print(f"   수집시각 기준 {target_date} 기사: {len(result)}개")
    return result


# ── 뉴스 출처 추출 ───────────────────────────────────
SOURCE_MAP = {
    "nst.com.my":        "NST",
    "malaymail.com":     "Malay Mail",
    "bernama.com":       "Bernama",
    "thestar.com.my":    "The Star",
    "freemalaysiatoday": "FMT",
    "malaysiakini.com":  "Malaysiakini",
    "sinchew":           "Sin Chew",
    "hmetro.com.my":     "Harian Metro",
    "utusan.com.my":     "Utusan",
    "astroawani.com":    "Astro Awani",
}

def get_source_name(url):
    for key, name in SOURCE_MAP.items():
        if key in url:
            return name
    m = re.search(r'https?://(?:www\.)?([^/]+)', url)
    if m:
        domain = m.group(1).split('.')[0].upper()
        return domain
    return "뉴스"


# ── OG 이미지 추출 ────────────────────────────────────
OG_PATTERNS = [
    r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\'](https?://[^"\'>\s]+)',
    r'<meta[^>]+content=["\'](https?://[^"\'>\s]+)["\'][^>]+property=["\']og:image["\']',
    r'<meta[^>]+name=["\']twitter:image["\'][^>]+content=["\'](https?://[^"\'>\s]+)',
    r'<meta[^>]+content=["\'](https?://[^"\'>\s]+)["\'][^>]+name=["\']twitter:image["\']',
]

def get_og_image(url, fallback=""):
    if not url or not url.startswith("http"):
        return fallback
    try:
        html = http_get(url, timeout=10)
        for pat in OG_PATTERNS:
            m = re.search(pat, html, re.IGNORECASE)
            if m:
                img = m.group(1).strip()
                if img and not img.endswith(".svg"):
                    return img
    except Exception as e:
        print(f"   ⚠️  이미지 추출 실패 ({url[:60]}...): {e}")
    return fallback


# ── 요약 ─────────────────────────────────────────────
def _ai_summarize(title, summary):
    msg = anthropic.Anthropic().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=200,
        messages=[{
            "role": "user",
            "content": (
                f"다음 말레이시아 뉴스를 카드뉴스용으로 2~3문장(4~5줄)으로 "
                f"자연스럽게 한국어 요약해줘. 핵심 정보만 담고 간결하게. 요약문만 출력해.\n\n"
                f"제목: {title}\n내용: {summary}"
            ),
        }],
    )
    return msg.content[0].text.strip()


def _sentence_summarize(text, max_sentences=3):
    sentences = re.split(r'(?<=[.!?。]) +|(?<=다\.) +|(?<=다) (?=[가-힣A-Z])', text)
    sentences = [s.strip() for s in sentences if s.strip()]
    return ' '.join(sentences[:max_sentences])


def summarize(title, summary):
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if api_key and anthropic:
        try:
            return _ai_summarize(title, summary)
        except Exception as e:
            print(f"   ⚠️  AI 요약 실패: {e}")
    return _sentence_summarize(summary)


# ── HTML 생성 ─────────────────────────────────────────
CSS = """
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif;
    background: #e8e8e8;
    display: flex; flex-direction: column; align-items: center; padding: 40px 20px;
  }
  .slideshow-wrapper { position: relative; width: 540px; }
  .slides { overflow: hidden; width: 540px; }
  .slide-track { display: flex; transition: transform 0.38s cubic-bezier(0.4,0,0.2,1); will-change: transform; }
  .nav { display: flex; align-items: center; justify-content: center; gap: 14px; margin-top: 22px; }
  .nav button {
    width: 40px; height: 40px; border: none; border-radius: 50%;
    background: rgba(0,0,0,0.15); color: #333; font-size: 18px; cursor: pointer;
    transition: background 0.2s; display: flex; align-items: center; justify-content: center;
  }
  .nav button:hover { background: rgba(0,0,0,0.28); }
  .nav button:disabled { opacity: 0.25; cursor: default; }
  .dots { display: flex; gap: 7px; }
  .dot { width: 8px; height: 8px; border-radius: 50%; background: rgba(0,0,0,0.2); cursor: pointer; transition: all 0.2s; }
  .dot.active { background: #F97316; width: 22px; border-radius: 4px; }
  .slide-counter { text-align: center; font-size: 13px; color: #999; margin-top: 10px; }
  .card { width: 540px; height: 675px; position: relative; overflow: hidden; flex-shrink: 0; }

  /* 배지 - 상단 중앙 pill */
  .badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: #F97316; color: #fff;
    font-size: 24px; font-weight: 700; padding: 16px 40px; border-radius: 60px; letter-spacing: 0.5px;
  }
  .badge-sep { opacity: 0.65; font-weight: 400; }
  .badge-wrap { position: absolute; top: 36px; left: 50%; transform: translateX(-50%); z-index: 10; white-space: nowrap; }

  /* 워터마크 - 하단 중앙 */
  .wm { position: absolute; z-index: 30; font-size: 32px; font-weight: 800; letter-spacing: 2px; white-space: nowrap; }
  .wm-bottom {
    bottom: 24px; left: 50%; transform: translateX(-50%);
    color: #F97316; text-shadow: 0 1px 6px rgba(0,0,0,0.55);
  }

  /* COVER */
  .card-cover { background: #111; }
  .cover-bg { position: absolute; inset: 0; background-size: cover; background-position: center; opacity: 0.55; }
  .cover-grad { position: absolute; inset: 0; background: linear-gradient(to top, rgba(10,10,10,0.95) 0%, rgba(10,10,10,0.45) 55%, transparent 100%); }
  .cover-body { position: absolute; top: 54%; left: 0; right: 0; z-index: 5; padding: 0 40px; text-align: center; transform: translateY(-50%); }
  .cover-body h1 { font-size: 52px; font-weight: 900; color: #fff; line-height: 1.18; margin-bottom: 16px; word-break: keep-all; }
  .cover-body p { font-size: 28px; color: rgba(255,255,255,0.6); line-height: 1.7; word-break: keep-all; }

  /* ARTICLE - 전체 이미지 배경 */
  .card-article { background: #111; }
  .art-bg { position: absolute; inset: 0; background-size: cover; background-position: center; opacity: 0.5; }
  .art-grad { position: absolute; inset: 0; background: linear-gradient(to top, rgba(5,5,5,0.97) 0%, rgba(5,5,5,0.6) 42%, rgba(0,0,0,0.1) 100%); }
  .art-body { position: absolute; bottom: 0; left: 0; right: 0; z-index: 5; padding: 0 38px 72px; }
  .art-body h2 { font-size: 40px; font-weight: 900; color: #fff; line-height: 1.22; margin-bottom: 16px; word-break: keep-all; }
  .art-body p { font-size: 14px; color: rgba(255,255,255,0.78); line-height: 1.78; word-break: keep-all; }
  .art-body .src { font-size: 11px; color: rgba(255,255,255,0.3); margin-top: 14px; }

  /* OUTRO */
  .card-outro { background: #111; }
  .outro-bg { position: absolute; inset: 0; background-size: cover; background-position: center; opacity: 0.25; }
  .outro-grad { position: absolute; inset: 0; background: rgba(10,10,10,0.78); }
  .outro-body { position: relative; z-index: 5; height: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 48px; text-align: center; }
  .outro-logo { font-size: 42px; font-weight: 900; color: #F97316; letter-spacing: 4px; margin-bottom: 8px; }
  .outro-sub { font-size: 13px; color: rgba(255,255,255,0.4); margin-bottom: 28px; letter-spacing: 1px; }
  .outro-line { width: 44px; height: 3px; background: #F97316; border-radius: 2px; margin: 0 auto 32px; }
  .outro-body h2 { font-size: 24px; font-weight: 800; color: #fff; line-height: 1.5; word-break: keep-all; margin-bottom: 16px; }
  .outro-body p { font-size: 13px; color: rgba(255,255,255,0.45); line-height: 1.85; word-break: keep-all; margin-bottom: 40px; }
  .outro-tags { display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; }
  .outro-tag { border: 1.5px solid rgba(249,115,22,0.45); color: rgba(249,115,22,0.85); font-size: 12px; font-weight: 600; padding: 6px 14px; border-radius: 20px; }
  .gen-info { font-size: 12px; color: #aaa; text-align: center; margin-top: 10px; }
"""

JS = """
  const TOTAL = {total};
  let cur = 0;
  const track   = document.getElementById('slideTrack');
  const dotsEl  = document.getElementById('dots');
  const counter = document.getElementById('counter');
  const btnPrev = document.getElementById('btnPrev');
  const btnNext = document.getElementById('btnNext');
  for (let i = 0; i < TOTAL; i++) {
    const d = document.createElement('div');
    d.className = 'dot';
    d.onclick = () => go(i);
    dotsEl.appendChild(d);
  }
  function go(n) {
    if (n < 0 || n >= TOTAL) return;
    cur = n;
    track.style.transform = `translateX(${-540 * cur}px)`;
    dotsEl.querySelectorAll('.dot').forEach((d, i) => d.classList.toggle('active', i === cur));
    btnPrev.disabled = cur === 0;
    btnNext.disabled = cur === TOTAL - 1;
    counter.textContent = `${cur + 1} / ${TOTAL}`;
  }
  document.addEventListener('keydown', e => {
    if (e.key === 'ArrowRight') go(cur + 1);
    if (e.key === 'ArrowLeft')  go(cur - 1);
  });
  let tx = 0;
  track.addEventListener('touchstart', e => tx = e.touches[0].clientX);
  track.addEventListener('touchend',   e => { const d = tx - e.changedTouches[0].clientX; if (Math.abs(d) > 50) go(cur + (d > 0 ? 1 : -1)); });
  go(0);
"""


def card_cover(cover_img, target_date):
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    date_short = dt.strftime("%y.%m.%d")
    weekday = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
    date_full = dt.strftime("%Y년 %m월 %d일")
    return f"""
      <div class="card card-cover">
        <div class="cover-bg" style="background-image:url('{escape(cover_img)}')"></div>
        <div class="cover-grad"></div>
        <div class="badge-wrap">
          <span class="badge">말레이시아 뉴스 <span class="badge-sep">|</span> {date_short}</span>
        </div>
        <div class="cover-body">
          <h1>오늘의<br>말레이시아</h1>
          <p>{date_full} ({weekday})<br>스와이프해서 주요 뉴스를 확인하세요 👉</p>
        </div>
        <span class="wm wm-bottom">🍊한마당</span>
      </div>"""


def card_article(num, title, summary, img_url, url, date_short):
    bg = f"background-image:url('{escape(img_url)}')" if img_url else "background:#222"
    source = get_source_name(url)
    return f"""
      <div class="card card-article">
        <div class="art-bg" style="{bg}"></div>
        <div class="art-grad"></div>
        <div class="badge-wrap">
          <span class="badge">말레이시아 이슈 <span class="badge-sep">|</span> {date_short}</span>
        </div>
        <div class="art-body">
          <h2>{escape(title)}</h2>
          <p>{escape(summary)}</p>
          <div class="src">출처 : {escape(source)}</div>
        </div>
        <span class="wm wm-bottom">🍊한마당</span>
      </div>"""


def card_outro(outro_img, target_date):
    return f"""
      <div class="card card-outro">
        <div class="outro-bg" style="background-image:url('{escape(outro_img)}')"></div>
        <div class="outro-grad"></div>
        <div class="outro-body">
          <div class="outro-logo">🍊한마당</div>
          <div class="outro-sub">말레이시아 교민 뉴스 채널</div>
          <div class="outro-line"></div>
          <h2>더 많은 말레이시아 소식을<br>한마당에서 만나세요</h2>
          <p>매일 업데이트되는 말레이시아 실시간 뉴스<br>정치·경제·사회·문화를 한눈에</p>
          <div class="outro-tags">
            <span class="outro-tag">#말레이시아뉴스</span>
            <span class="outro-tag">#한마당</span>
            <span class="outro-tag">#실시간뉴스</span>
            <span class="outro-tag">#{target_date.replace('-','')}</span>
          </div>
        </div>
      </div>"""


def build_html(slides_html, total, gen_time):
    js = JS.replace("{total}", str(total))
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>한마당 실시간 카드뉴스</title>
<style>{CSS}</style>
</head>
<body>
<div class="slideshow-wrapper">
  <div class="slides">
    <div class="slide-track" id="slideTrack">
{slides_html}
    </div>
  </div>
  <div class="nav">
    <button id="btnPrev" onclick="go(cur-1)">&#8592;</button>
    <div class="dots" id="dots"></div>
    <button id="btnNext" onclick="go(cur+1)">&#8594;</button>
  </div>
  <div class="slide-counter" id="counter"></div>
</div>
<div class="gen-info">자동 생성: {gen_time}</div>
<script>{js}</script>
</body>
</html>"""


# ── 이미지 내보내기 ───────────────────────────────────
def export_images(html_path, total_cards, target_date):
    """playwright로 각 카드를 1080×1350 PNG로 저장"""
    from playwright.sync_api import sync_playwright

    out_dir = os.path.join(os.path.dirname(html_path), "images", target_date)
    os.makedirs(out_dir, exist_ok=True)

    print(f"\n📸 이미지 내보내는 중 → {out_dir}")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(
            viewport={"width": 540, "height": 675},
            device_scale_factor=2,   # 2× → 1080×1350px
        )
        page.goto(f"file://{html_path}", wait_until="networkidle")
        page.wait_for_timeout(1500)

        for i in range(total_cards):
            page.evaluate(f"go({i})")
            page.wait_for_timeout(500)

            card = page.query_selector(".slide-track .card:nth-child({})".format(i + 1))
            path = os.path.join(out_dir, f"템플릿{i+1}.png")
            card.screenshot(path=path)
            print(f"   ✅ 템플릿{i+1}.png")

        browser.close()

    print(f"\n🎉 {total_cards}장 저장 완료: {out_dir}\n")
    return out_dir


# ── 이미지 GitHub Release 업로드 ──────────────────────
def upload_to_github_release(image_dir, target_date):
    """이미지를 GitHub Release에 업로드하고 공개 URL 목록 반환"""
    import subprocess, shutil

    repo = os.environ.get("GITHUB_REPOSITORY", "minsnj/hanmadang-cardnews")
    tag  = f"cardnews-{target_date}"

    images = sorted(
        glob.glob(os.path.join(image_dir, "템플릿*.png")),
        key=lambda x: int(re.search(r'\d+', os.path.basename(x)).group())
    )
    if not images:
        raise RuntimeError("업로드할 이미지가 없습니다.")

    # URL-safe 파일명으로 복사 (card01.png, card02.png, ...)
    safe_paths = []
    for i, src in enumerate(images):
        dst = os.path.join(image_dir, f"card{i+1:02d}.png")
        shutil.copy2(src, dst)
        safe_paths.append(dst)

    # 기존 release 삭제 후 재생성 (재실행 대비)
    subprocess.run(["gh", "release", "delete", tag, "--yes", "--repo", repo],
                   capture_output=True)

    subprocess.run([
        "gh", "release", "create", tag, *safe_paths,
        "--title", f"카드뉴스 {target_date}",
        "--notes", "자동 생성",
        "--repo", repo,
    ], check=True)

    base = f"https://github.com/{repo}/releases/download/{tag}"
    urls = [f"{base}/card{i+1:02d}.png" for i in range(len(safe_paths))]
    print(f"   {len(urls)}개 이미지 업로드 완료")
    return urls


# ── Instagram Graph API 포스팅 ─────────────────────────
def refresh_instagram_token(access_token):
    """토큰 갱신 후 GitHub Secret 업데이트. 실패해도 기존 토큰으로 계속 진행."""
    import json, subprocess
    try:
        url = f"https://graph.instagram.com/refresh_access_token?grant_type=ig_refresh_token&access_token={access_token}"
        with urllib.request.urlopen(url) as r:
            data = json.loads(r.read())
        new_token = data.get("access_token", "")
        if not new_token:
            print("⚠️  토큰 갱신 응답에 access_token이 없습니다.")
            return access_token
        repo = os.environ.get("GITHUB_REPOSITORY", "")
        if repo:
            result = subprocess.run(
                ["gh", "secret", "set", "INSTAGRAM_ACCESS_TOKEN", "--body", new_token, "--repo", repo],
                capture_output=True, text=True,
                env={**os.environ, "GH_TOKEN": os.environ.get("GH_TOKEN", os.environ.get("GH_PAT", ""))}
            )
            if result.returncode == 0:
                print("   ✅ 토큰 갱신 및 Secret 업데이트 완료")
            else:
                print(f"   ⚠️  Secret 업데이트 실패 (포스팅은 계속): {result.stderr.strip()}")
        return new_token
    except Exception as e:
        print(f"⚠️  토큰 갱신 실패 (포스팅은 계속): {e}")
        return access_token


def post_via_graph_api(image_dir, target_date):
    """Instagram Graph API로 캐러셀 포스팅"""
    import json, time as _time

    access_token = os.environ.get("INSTAGRAM_ACCESS_TOKEN", "")
    user_id      = os.environ.get("INSTAGRAM_USER_ID", "")

    if not access_token or not user_id:
        print("⚠️  INSTAGRAM_ACCESS_TOKEN / INSTAGRAM_USER_ID가 없습니다. 포스팅을 건너뜁니다.")
        return

    print("\n🔄 토큰 갱신 중...")
    access_token = refresh_instagram_token(access_token)

    dt       = datetime.strptime(target_date, "%Y-%m-%d")
    date_str = dt.strftime("%Y년 %m월 %d일")
    weekday  = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
    caption  = (
        f"📰 {date_str} ({weekday}) 말레이시아 실시간 뉴스\n\n"
        "오늘의 말레이시아 주요 소식을 한마당이 전합니다.\n"
        "스와이프해서 더 많은 뉴스를 확인하세요 👉\n\n"
        "#말레이시아뉴스 #한마당 #실시간뉴스 #말레이시아 #Malaysia "
        f"#{target_date.replace('-', '')} #해외뉴스 #카드뉴스 #교민뉴스 #말레이시아한인 #한인생활"
    )

    def ig_post(path, data):
        url  = f"https://graph.instagram.com/v21.0{path}"
        data["access_token"] = access_token
        body = urllib.parse.urlencode(data).encode()
        req  = urllib.request.Request(url, data=body, method="POST")
        try:
            with urllib.request.urlopen(req) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            raise RuntimeError(f"Graph API 오류 ({e.code}): {e.read().decode()}")

    # 1. 이미지 GitHub Release에 업로드
    print("\n📤 이미지 업로드 중...")
    image_urls = upload_to_github_release(image_dir, target_date)
    _time.sleep(5)  # Release 전파 대기

    # 2. 이미지 컨테이너 생성
    print(f"\n📸 이미지 컨테이너 생성 중 ({len(image_urls)}개)...")
    container_ids = []
    for i, img_url in enumerate(image_urls):
        res = ig_post(f"/{user_id}/media", {
            "image_url":        img_url,
            "is_carousel_item": "true",
        })
        container_ids.append(res["id"])
        print(f"   [{i+1}/{len(image_urls)}] {res['id']}")
        _time.sleep(1)

    # 3. 캐러셀 컨테이너 생성
    print("\n🎠 캐러셀 컨테이너 생성 중...")
    carousel = ig_post(f"/{user_id}/media", {
        "media_type": "CAROUSEL",
        "children":   ",".join(container_ids),
        "caption":    caption,
    })
    print(f"   캐러셀 ID: {carousel['id']}")

    # 4. 게시
    _time.sleep(5)
    print("\n🚀 게시 중...")
    result = ig_post(f"/{user_id}/media_publish", {"creation_id": carousel["id"]})
    media_id = result["id"]
    print(f"   ✅ 게시 완료! 게시물 ID: {media_id}")

    # 5. permalink 조회 (스토리는 로컬 auto_story.py가 담당 — 아래 주석 참고)
    try:
        perm_url = (f"https://graph.instagram.com/v21.0/{media_id}"
                    f"?fields=permalink&access_token={access_token}")
        with urllib.request.urlopen(perm_url) as r:
            permalink = json.loads(r.read()).get("permalink", "")
        print(f"   🔗 게시물 링크: {permalink}")
    except Exception as e:
        print(f"⚠️  permalink 조회 실패: {e}")

    # ⚠️ 스토리 포스팅을 여기서 하지 않는다.
    # GitHub Actions IP는 인스타그램이 스토리 '링크스티커'를 차단하므로
    # (링크 없는 스토리만 올라감) 스토리는 로컬 Mac launchd(auto_story.py)가
    # 링크스티커 포함해서 단독으로 담당한다. 여기서 post_story()를 호출하면
    # 링크 없는 스토리가 중복으로 올라가므로 절대 호출하지 말 것.
    # post_story(image_dir, target_date, permalink, access_token, user_id, ig_post)  # 사용 금지


def generate_story_png(first_card_path, output_dir):
    """story.html을 Playwright로 렌더링하여 story.png 생성"""
    import shutil, tempfile
    from playwright.sync_api import sync_playwright

    story_html = os.path.join(os.path.dirname(os.path.abspath(__file__)), "story.html")
    if not os.path.exists(story_html):
        raise FileNotFoundError("story.html이 없습니다.")

    # card_01.png를 절대 경로로 바꿔 임시 HTML 생성
    with open(story_html, encoding="utf-8") as f:
        html = f.read()
    html = html.replace("exports/card_01.png", first_card_path.replace("\\", "/"))

    with tempfile.NamedTemporaryFile(suffix=".html", mode="w", delete=False, encoding="utf-8") as tmp:
        tmp.write(html)
        tmp_path = tmp.name

    story_png = os.path.join(output_dir, "story.png")
    import time as _t
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 540, "height": 960}, device_scale_factor=2)
            page.goto(f"file://{tmp_path}", wait_until="networkidle")
            _t.sleep(1.5)
            page.screenshot(path=story_png, clip={"x": 0, "y": 0, "width": 540, "height": 960})
            browser.close()
        print(f"   ✅ story.png 생성됨")
    finally:
        os.unlink(tmp_path)

    return story_png


def post_story(image_dir, target_date, permalink, access_token, user_id, ig_post):
    """story.png 생성 후 instagrapi(패치)로 링크스티커 포함 스토리 게시"""
    import json, re, types, time as _time

    print("\n📖 스토리 생성 중...")

    # 첫 번째 카드 이미지 경로
    first_card = os.path.join(image_dir, "card01.png")
    if not os.path.exists(first_card):
        cards = sorted(glob.glob(os.path.join(image_dir, "card*.png")))
        if not cards:
            cards = sorted(glob.glob(os.path.join(image_dir, "템플릿*.png")))
        first_card = cards[0] if cards else None
    if not first_card:
        print("⚠️  첫 번째 카드 이미지를 찾을 수 없습니다.")
        return

    story_png = generate_story_png(first_card, image_dir)

    # ── instagrapi 클라이언트 (패치 적용) ──────────────────
    from instagrapi import Client
    from instagrapi.extractors import extract_media_v1
    from instagrapi.types import StoryLink

    session_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ig_session.json")
    if not os.path.exists(session_path):
        print("⚠️  .ig_session.json 없음 — 스토리 건너뜀")
        return

    with open(session_path) as f:
        settings = json.load(f)
    cookies = settings.get("cookies", {})
    if not cookies.get("sessionid"):
        print("⚠️  세션 만료 — 스토리 건너뜀")
        return

    cl = Client()
    cl.set_locale("ko_KR")
    cl.set_timezone_offset(32400)
    cl.delay_range = [1, 3]
    for k, v in cookies.items():
        cl.private.cookies.set(k, v)
    user_id_str = cookies.get("ds_user_id", "")
    if not user_id_str:
        m = re.match(r"^\d+", cookies.get("sessionid", ""))
        user_id_str = m.group() if m else ""
    cl.authorization_data = {
        "ds_user_id": user_id_str,
        "sessionid":  cookies["sessionid"],
        "should_use_header_over_cookies": True,
    }
    cl.username = "hanmadang_my"
    if "mid" in cookies:
        cl.mid = cookies["mid"]

    # 패치 1: validate_reel_url bypass
    _saved: dict = {}
    _orig = cl.private_request.__func__
    def _private_request(self, endpoint, *args, **kwargs):
        if endpoint == "media/validate_reel_url/":
            return {"status": "ok"}
        result = _orig(self, endpoint, *args, **kwargs)
        if endpoint == "media/configure_to_story/" and isinstance(result, dict):
            _saved.clear(); _saved.update(result)
        return result
    cl.private_request = types.MethodType(_private_request, cl)

    # 패치 2: media 키 없을 때 캐시에서 복구
    def _extract(self, configured, exception_cls, context: str):
        media = configured.get("media") if isinstance(configured, dict) else None
        if media is None and _saved:
            media = _saved.get("media")
        if media is None:
            raise exception_cls(
                f"{context} configure succeeded without media payload",
                response=self.last_response,
                **(self.last_json if isinstance(self.last_json, dict) else {}),
            )
        return extract_media_v1(media)
    cl._extract_configured_media_or_raise = types.MethodType(_extract, cl)

    cl.load_settings(session_path)

    # ── 스토리 업로드 ──────────────────────────────────────
    print("\n📤 스토리 업로드 중 (링크스티커)...")
    story = cl.photo_upload_to_story(
        str(story_png),
        links=[StoryLink(webUri=permalink)],
    )
    print(f"   ✅ 스토리 게시 완료! ID: {story.pk}")
    print(f"   🔗 링크스티커: {permalink}")


# ── 메인 ──────────────────────────────────────────────
def main():
    target_date = sys.argv[1] if len(sys.argv) > 1 else datetime.now(MYT).strftime("%Y-%m-%d")
    print(f"\n🗓  대상 날짜: {target_date}")

    # 1. 시트 데이터 가져오기
    rows = fetch_sheet()

    # 2. 당일 뉴스 필터
    today_rows = filter_by_date(rows, target_date)
    if not today_rows:
        print(f"❌ {target_date} 날짜의 기사가 없습니다.")
        sys.exit(1)

    # 3. 한국-말레이시아 연관 기사 우선 정렬 후 MAX_CARDS 개 선택
    KR_KEYWORDS = [
        "한국", "한인", "교민", "코리아", "한류", "K-pop", "K-drama", "케이팝", "케이드라마",
        "삼성", "LG", "현대", "기아", "롯데", "CJ", "SK", "포스코",
        "Korea", "Korean", "Koreans", "Seoul", "Busan",
        "Samsung", "Hyundai", "Kia", "Lotte",
    ]

    def korea_score(row):
        text = " ".join([row[3] if len(row) > 3 else "", row[4] if len(row) > 4 else ""])
        return sum(1 for kw in KR_KEYWORDS if kw.lower() in text.lower())

    today_rows_sorted = sorted(today_rows, key=korea_score, reverse=True)
    selected = today_rows_sorted[:MAX_CARDS]

    kr_related = sum(1 for r in selected if korea_score(r) > 0)
    print(f"   한국-말레이시아 연관 기사 {kr_related}개 우선 선정")

    # 4. 각 기사 OG 이미지 크롤링
    print(f"\n🖼  기사 이미지 크롤링 중 ({len(selected)}개)...")
    articles = []
    for i, row in enumerate(selected):
        date_str = row[0][:10]
        cat      = row[1].strip()
        title_kr = row[3].strip()
        summary  = row[4].strip()
        url      = row[5].strip()

        print(f"   [{i+1}/{len(selected)}] {title_kr[:30]}...")
        fallback = FALLBACK_IMAGES.get(cat, DEFAULT_FALLBACK)
        img = get_og_image(url, fallback)
        short = summarize(title_kr, summary)

        articles.append({
            "cat": cat, "title": title_kr, "summary": short,
            "img": img, "date": date_str, "url": url,
        })

    # 5. HTML 슬라이드 조립
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    date_short = dt.strftime("%y.%m.%d")

    cover_img = articles[0]["img"] if articles else DEFAULT_FALLBACK
    outro_img = articles[-1]["img"] if articles else DEFAULT_FALLBACK

    slides = [card_cover(cover_img, target_date)]

    for i, art in enumerate(articles):
        slides.append(card_article(i + 1, art["title"], art["summary"], art["img"], art["url"], date_short))

    slides.append(card_outro(outro_img, target_date))

    total    = len(slides)
    gen_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html     = build_html("\n".join(slides), total, gen_time)

    # 6. 파일 저장 (임시 파일로 쓴 뒤 이동 - macOS EDEADLK 우회)
    import tempfile, shutil
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=os.path.dirname(OUTPUT),
                                     suffix=".html.tmp", delete=False) as tmp:
        tmp.write(html)
        tmp_path = tmp.name
    shutil.move(tmp_path, OUTPUT)

    print(f"\n✅ 완료! 총 {total}장 ({len(articles)}개 기사)")
    print(f"   저장 위치: {OUTPUT}")
    print(f"   생성 시각: {gen_time}")

    # PNG 이미지 내보내기
    image_dir = export_images(OUTPUT, total, target_date)

    # 인스타그램 자동 포스팅
    post_via_graph_api(image_dir, target_date)


if __name__ == "__main__":
    main()

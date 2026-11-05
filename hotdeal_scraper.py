"""
퀘이사존 / 아카라이브 핫딜 목록 수집기 (Scrapling 0.5 기반)

사용 예:
    python hotdeal_scraper.py                         # 두 사이트, 최대 4페이지(3분 간격) → deals.json + deals.xlsx
    python hotdeal_scraper.py --pages 1               # 1페이지만
    python hotdeal_scraper.py -o 핫딜.json            # → 핫딜.json + 핫딜.xlsx (이전 결과는 before.json)
    python hotdeal_scraper.py --excel 오늘핫딜.xlsx   # 엑셀 파일 이름만 바꾸기
    python hotdeal_scraper.py --no-excel              # JSON만 저장
    python hotdeal_scraper.py --site quasarzone --keyword SSD --keyword 모니터
    python hotdeal_scraper.py --dump-html             # 1페이지 원본 HTML 저장(셀렉터 점검용)
    python hotdeal_scraper.py --browser               # HTTP 대신 브라우저(StealthySession)로 렌더링

설치:
    pip install "scrapling[fetchers]"
    scrapling install        # --browser 옵션을 쓸 때만 필요

파일 처리:
- 시작할 때 결과 파일(기본 deals.json)이 있으면 before.json으로 이름을 바꿉니다. before.json이 이미 있으면 덮어씁니다.
- 수집이 끝나면 deals.json을 새로 만들어 저장하고, 같은 이름의 엑셀(deals.xlsx)로 변환합니다.

페이지 이동 규칙:
- 사이트마다 1페이지부터 수집하고, 한 페이지를 처리한 뒤 --interval(기본 3분)만큼 기다렸다가 다음 페이지로 갑니다.
- 어떤 페이지에 before.json에 있는 글(같은 글 번호 또는 같은 제목)이 나오면 그 사이트만 거기서 멈춥니다.
  before.json에 있는 글은 다시 저장하지 않습니다.
- 제목은 같지만 before.json의 글보다 나중에 작성된 글(매주 올라오는 쿠폰 글 등)은 새 글로 저장하고 멈추지 않습니다.
- 작성된 지 --max-age-hours(기본 24시간)가 지난 글이 나오면 그 사이트만 거기서 멈추고, 그 글은 저장하지 않습니다.
- 같은 실행 안에서 다음 페이지로 밀려 내려온 글은 중복 저장하지 않으며, 멈추는 조건으로 보지 않습니다.

주의:
- 퀘이사존 셀렉터는 2026-11-04 실제 페이지로 검증했고, 아카라이브는 실제 수집 결과로만 확인했습니다.
  사이트 개편 시 깨질 수 있으므로 결과가 0건이면 --dump-html로 저장한 HTML을 보고 SITES 항목만 고치면 됩니다.
- robots.txt를 준수하고(robots_txt_obey=True), 요청 간 지연을 둡니다.
- 봇 차단(Cloudflare 챌린지 등) 자동 해제 기능은 의도적으로 넣지 않았습니다.
  차단 페이지가 감지되면 로그로 알리고 해당 요청은 건너뜁니다.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import anyio
from scrapling.spiders import Spider, Request, Response

from export_excel import posted_time, write_xlsx


@dataclass
class Seen:
    """이전 결과(before.json)에 있던 글. 사이트별로 글 번호와, 제목별 (글 번호, 작성 시각)을 모아 둔다."""
    ids: set[str] = field(default_factory=set)
    titles: dict[str, list[tuple[str, datetime | None]]] = field(default_factory=dict)

    def match(self, item: dict[str, Any]) -> str | None:
        """'same': 이전에 수집한 글. 'repost': 제목은 같지만 그 뒤에 새로 올라온 글(매주 올라오는 쿠폰 글 등).
        None: 처음 보는 글."""
        if item["post_id"] in self.ids:
            return "same"
        olds = self.titles.get(item["title"]) if item["title"] else None
        if not olds:
            return None
        new_dt = posted_at(item)
        return "repost" if all(is_newer(item["post_id"], new_dt, *old) for old in olds) else "same"


def is_newer(new_id: str, new_dt: datetime | None, old_id: str, old_dt: datetime | None) -> bool:
    """작성 시각으로 새 글이 더 나중인지 판단. 시각을 모르거나 같으면 글 번호로 판단(번호가 클수록 나중 글)."""
    if new_dt is not None and old_dt is not None and new_dt != old_dt:
        return new_dt > old_dt
    return new_id.isdigit() and old_id.isdigit() and int(new_id) > int(old_id)


def load_before(path: Path) -> dict[str, Seen]:
    """이전 결과 파일(.json/.jsonl/.csv)을 읽어 사이트 이름(퀘이사존/아카라이브)별로 정리. 없으면 빈 값."""
    if not path.exists():
        return {}
    text = path.read_text(encoding="utf-8-sig")
    ext = path.suffix.lower()
    if ext == ".csv":
        items = list(csv.DictReader(text.splitlines()))
    elif ext == ".jsonl":
        items = [json.loads(line) for line in text.splitlines() if line.strip()]
    else:
        items = json.loads(text) if text.strip() else []

    before: dict[str, Seen] = {}
    for it in items:
        seen = before.setdefault(str(it.get("site", "")), Seen())
        post_id = str(it.get("post_id") or "")
        if post_id:
            seen.ids.add(post_id)
        comments = it.get("comments")
        comments = int(comments) if str(comments).isdigit() else None
        title = strip_comment_suffix(str(it.get("title") or "").strip(), comments)
        if title:
            try:
                dt = posted_at(it)
            except ValueError:              # 수집 시각 형식이 잘못된 경우
                dt = None
            seen.titles.setdefault(title, []).append((post_id, dt))
    return before


DATE_ONLY = re.compile(r"(?:(\d{4})[.\-/])?(\d{1,2})[.\-/](\d{1,2})\.?")


def posted_at(item: dict[str, Any]) -> datetime | None:
    """작성 시각(한국 시각) 추정. 알 수 없으면 None.

    - ISO 시각(아카라이브), 'N분/시간/일 전'(퀘이사존)은 export_excel.posted_time으로 변환.
    - '11.04' 같은 날짜만 있는 표기: 퀘이사존은 24시간 안의 글을 'N시간 전'으로 보여 주고
      그보다 오래된 글만 날짜로 보여 주므로, 수집 시각에서 24시간 전과 그날 끝 중 이른 쪽으로 본다.
    """
    posted = str(item.get("posted") or "").strip()
    collected = str(item.get("collected_at") or "")
    dt = posted_time(posted, collected)
    if dt is not None or not collected:
        return dt
    m = DATE_ONLY.fullmatch(posted)
    if not m:
        return None
    base = datetime.fromisoformat(collected)
    month, day = int(m.group(2)), int(m.group(3))
    year = int(m.group(1)) if m.group(1) else base.year
    try:
        day_end = datetime(year, month, day, 23, 59, 59)
        if not m.group(1) and day_end.date() > base.date():     # 연도가 없고 미래 날짜면 작년 글
            day_end = datetime(year - 1, month, day, 23, 59, 59)
    except ValueError:
        return None
    return min(day_end, base - timedelta(hours=24))


def is_too_old(item: dict[str, Any], max_age: timedelta | None) -> bool:
    """수집 시각 기준으로 max_age 이상 지난 글이면 True. 작성 시각을 알 수 없으면 False."""
    if max_age is None:
        return False
    dt = posted_at(item)
    return dt is not None and datetime.fromisoformat(item["collected_at"]) - dt >= max_age

# ---------------------------------------------------------------------------
# 사이트별 설정 — 셀렉터는 리스트 앞에서부터 차례로 시도하고, 처음 값이 나온 것을 씁니다.
# ---------------------------------------------------------------------------
SITES: dict[str, dict[str, Any]] = {
    "quasarzone": {
        "label": "퀘이사존",
        "list_url": "https://quasarzone.com/bbs/qb_saleinfo?page={page}",
        # 2026-11-04 실제 페이지(v2 목록) 기준으로 검증. 뒤쪽 후보는 이전 구조용.
        "row": ["div.v2-list div.v2-list-row", "div.market-type-list table tbody tr"],
        "skip_row_class": ["v2-partner-row"],             # 파트너(광고) 핫딜
        "link": ["a.subject-link", "p.tit a"],
        "title": ["a.subject-link", "span.ellipsis-with-reply-cnt"],
        "status": ["span.v2-thumb-done", "span.label"],   # 종료된 딜에만 '종료' 표시
        "category": ["span.v2-badge", "span.category"],
        "price": ["span.v2-list-row__price", "span.text-orange"],
        "store": ["span.v2-list-row__ship img::attr(alt)"],   # 판매처는 로고 이미지의 alt
        "delivery": ["span.v2-list-row__ship:last-child", "span.delivery"],
        # 댓글·조회·추천 수는 페이지 로드 후 qz-count-hydrate.js가 최신 값으로 다시 채운다.
        # HTTP 모드로 받은 HTML의 값은 그 이전 값이라 실제와 다를 수 있다(--browser 모드는 미검증).
        "comments": ["span.ctn-count"],
        "views": ["span.qc-count-hit", "span.count"],
        "date": ["span.v2-list-row__time", "span.date"],    # '11분 전' 같은 상대 표기 그대로
        "votes": ["span.qc-count-good", "span.num.num"],
        "id_regex": r"/qb_saleinfo/views/(\d+)",
    },
    "arca": {
        "label": "아카라이브",
        "list_url": "https://arca.live/b/hotdeal?p={page}",
        "row": ["div.list-table div.vrow.hybrid", "div.list-table a.vrow:not(.notice)"],
        # 2026-11-04 저장된 실제 페이지로 검증.
        "skip_row_class": ["notice", "head"],
        "skip_if_text": [("span.category-badge", "광고")],  # 광고 글
        "link": ["a.title.hybrid-title", "a.title", "a[href*='/b/hotdeal/']"],
        "title": ["a.title.hybrid-title", "span.title", "a.title"],
        "done_marker": [".deal-close"],                    # 종료된 딜은 줄 전체에 deal-close class
        "category": ["a.badge", "span.badge"],
        "price": ["span.deal-price"],
        "store": ["span.deal-store"],
        "delivery": ["span.deal-delivery"],
        "comments": ["span.comment-count", "span.info"],
        "views": ["span.col-view", "span.vcol.col-view"],
        "date": ["time"],
        "votes": ["span.col-rate", "span.vcol.col-rate"],
        "id_regex": r"/b/hotdeal/(\d+)",
    },
}

BLOCK_MARKERS = ("cf-chl", "challenge-platform", "Just a moment", "Attention Required")


def first_text(el, selectors: list[str]) -> str:
    """셀렉터 후보를 차례로 시도해 처음 나온 텍스트(공백 정리)를 반환."""
    for sel in selectors or []:
        found = el.css(sel)
        if found:
            text = found[0].get_all_text(separator=" ", strip=True)
            text = re.sub(r"\s+", " ", str(text)).strip()
            if text:
                return text
    return ""


def first_node(el, selectors: list[str]):
    for sel in selectors or []:
        found = el.css(sel)
        if found:
            return found[0]
    return None


def to_int(text: str) -> int | None:
    """'1.2k', '2,394', '(3)' 같은 표기를 정수로 변환. 실패하면 None."""
    if not text:
        return None
    t = text.lower().replace(",", "").strip("[]() ")
    m = re.search(r"(\d+(?:\.\d+)?)\s*([km만]?)", t)
    if not m:
        return None
    num = float(m.group(1))
    mult = {"k": 1_000, "m": 1_000_000, "만": 10_000}.get(m.group(2), 1)
    return int(num * mult)


def strip_comment_suffix(title: str, comments: int | None) -> str:
    """제목 끝에 붙은 댓글 수 표기('... [5]')가 댓글 수와 같으면 제거."""
    m = re.search(r"\s*\[(\d+)\]\s*$", title or "")
    if m and comments is not None and int(m.group(1)) == comments:
        return title[:m.start()].rstrip()
    return title


def parse_rows(page, site_key: str) -> list[dict[str, Any]]:
    """목록 페이지(Scrapling Selector/Response)에서 핫딜 행을 추출. 네트워크와 무관한 순수 함수."""
    cfg = SITES[site_key]
    rows = []
    for sel in cfg["row"]:
        rows = page.css(sel)
        if rows:
            break

    items: list[dict[str, Any]] = []
    skip = set(cfg.get("skip_row_class", []))
    for row in rows:
        classes = set(str(row.attrib.get("class", "")).split())
        if classes & skip:
            continue
        if any(first_text(row, [sel]) == text for sel, text in cfg.get("skip_if_text", [])):
            continue

        link_el = row if row.tag == "a" and row.attrib.get("href") else first_node(row, cfg["link"])
        if link_el is None:
            continue
        href = str(link_el.attrib.get("href", ""))
        m = re.search(cfg["id_regex"], href)
        if not m:
            continue
        url = page.urljoin(href.split("?")[0])

        title = first_text(row, cfg["title"])
        status = first_text(row, cfg.get("status", []))
        if not status and first_node(row, cfg.get("done_marker", [])) is not None:
            status = "종료"
        # 퀘이사존은 제목 링크 안에 상태 라벨이 함께 들어 있으므로 제목에서 제거
        if status and title.startswith(status):
            title = title[len(status):].strip()

        date_text = ""
        date_el = first_node(row, cfg["date"])
        if date_el is not None:
            date_text = str(date_el.attrib.get("datetime", "")) or first_text(row, cfg["date"])

        comments = to_int(first_text(row, cfg.get("comments", [])))
        title = strip_comment_suffix(title, comments)

        items.append({
            "site": cfg["label"],
            "post_id": m.group(1),
            "title": title,
            "url": url,
            "status": status,
            "category": first_text(row, cfg.get("category", [])),
            "price": first_text(row, cfg.get("price", [])),
            "store": first_text(row, cfg.get("store", [])),
            "delivery": first_text(row, cfg.get("delivery", [])),
            "comments": comments,
            "views": to_int(first_text(row, cfg.get("views", []))),
            "votes": to_int(first_text(row, cfg.get("votes", []))),
            "posted": date_text,
            "collected_at": datetime.now().isoformat(timespec="seconds"),
        })
    return items


class HotDealSpider(Spider):
    name = "hotdeal"
    robots_txt_obey = True          # robots.txt Disallow / Crawl-delay 준수
    concurrent_requests = 2
    concurrent_requests_per_domain = 1
    download_delay = 3.0            # 같은 도메인 요청 간 최소 3초
    autothrottle_enabled = True
    max_blocked_retries = 0         # 차단 응답은 재시도하지 않고 건너뜀
    logging_level = logging.INFO

    def __init__(self, sites: list[str], pages: int, keywords: list[str],
                 use_browser: bool, dump_dir: Path | None,
                 interval: float = 0.0, before: dict[str, Seen] | None = None,
                 max_age: timedelta | None = None):
        self.sites = sites
        self.pages = pages
        self.interval = interval            # 다음 페이지 요청 전 대기 시간(초)
        self.keywords = [k.lower() for k in keywords]
        self.use_browser = use_browser
        self.dump_dir = dump_dir
        self.before = before or {}          # 이전 결과(before.json)의 글, 사이트 이름별
        self.max_age = max_age              # 이보다 오래된 글이 나오면 그 사이트 중단 (None이면 제한 없음)
        self.seen: set[tuple[str, str]] = set()
        super().__init__()

    def configure_sessions(self, manager):
        if self.use_browser:
            from scrapling.fetchers import AsyncStealthySession
            manager.add("main", AsyncStealthySession(headless=True, network_idle=True))
        else:
            from scrapling.fetchers import FetcherSession
            manager.add("main", FetcherSession(impersonate="chrome"))

    def page_request(self, site: str, page: int) -> Request:
        return Request(SITES[site]["list_url"].format(page=page), sid="main",
                       meta={"site": site, "page": page})

    async def start_requests(self):
        # 1페이지만 먼저 요청하고, 다음 페이지는 parse()에서 조건을 보고 요청한다.
        for site in self.sites:
            yield self.page_request(site, 1)

    async def is_blocked(self, response: Response) -> bool:
        head = response.body[:20000].decode("utf-8", "ignore") if hasattr(response, "body") else ""
        blocked = await super().is_blocked(response) or any(mk in head for mk in BLOCK_MARKERS)
        if blocked:
            site = response.meta.get("site", "?")
            self.logger.warning(
                f"[{site}] 봇 차단 응답 감지 (HTTP {response.status}): {response.url} — 이 요청은 건너뜁니다."
            )
            if self.dump_dir and response.meta.get("page") == 1:
                self.dump_dir.mkdir(parents=True, exist_ok=True)
                out = self.dump_dir / f"{site}_page1_blocked.html"
                out.write_bytes(response.body)
                self.logger.info(f"차단 화면 HTML 저장: {out}")
        return blocked

    async def on_error(self, request: Request, error: Exception) -> None:
        self.logger.warning(f"요청 실패: {request.url} ({error})")

    async def parse(self, response: Response):
        site = response.meta.get("site")
        page_no = response.meta.get("page")

        if self.dump_dir and page_no == 1:
            self.dump_dir.mkdir(parents=True, exist_ok=True)
            out = self.dump_dir / f"{site}_page1.html"
            out.write_bytes(response.body)
            self.logger.info(f"HTML 저장: {out}")

        items = parse_rows(response, site)
        if not items:
            self.logger.warning(
                f"[{site} p{page_no}] 추출된 글이 0건입니다. 셀렉터가 맞지 않거나 차단됐을 수 있습니다. "
                f"--dump-html로 원본을 확인하세요."
            )
        label = SITES[site]["label"]
        before = self.before.get(label, Seen())

        new_count = already = too_old = reposts = 0
        for item in items:
            key = (item["site"], item["post_id"])
            if key in self.seen:            # 같은 실행 안의 중복 (새 글 때문에 다음 페이지로 밀린 글)
                continue
            self.seen.add(key)
            match = before.match(item)
            if match == "same":             # before.json에 있는 글: 저장하지 않고, 이 사이트는 이 페이지에서 멈춤
                already += 1
                continue
            if match == "repost":           # 같은 제목이지만 새로 올라온 글: 멈추지 않고 새 글로 저장
                reposts += 1
            if is_too_old(item, self.max_age):
                too_old += 1
                continue
            item["is_new"] = True
            new_count += 1
            if self.keywords and not any(k in item["title"].lower() for k in self.keywords):
                continue
            yield item

        self.logger.info(
            f"[{label} {page_no}페이지] {len(items)}건 중 새 글 {new_count}건"
            f" (이전 결과에 있음 {already}건, 기간 지남 {too_old}건, 같은 제목 새 글 {reposts}건)"
        )

        # 다음 페이지로 넘어갈지 결정
        if already:
            self.logger.info(f"[{label} {page_no}페이지] 이전 결과에 있는 글이 나와 {label} 수집을 마칩니다.")
            return
        if too_old:
            hours = self.max_age.total_seconds() / 3600
            self.logger.info(f"[{label} {page_no}페이지] {hours:g}시간이 지난 글이 나와 {label} 수집을 마칩니다.")
            return
        if page_no >= self.pages or not items:
            return
        if self.interval > 0:
            self.logger.info(f"[{label}] {self.interval:g}초 후 {page_no + 1}페이지를 수집합니다.")
            if not await self._wait(self.interval):
                self.logger.info(f"[{label}] 중단 요청으로 {page_no + 1}페이지는 수집하지 않습니다.")
                return
        yield self.page_request(site, page_no + 1)

    async def _wait(self, seconds: float) -> bool:
        """seconds만큼 기다린다. 도중에 Ctrl+C(중단 요청)가 들어오면 바로 False를 반환.

        Scrapling 0.5 엔진의 내부 속성(_pause_requested)을 읽으므로, 버전을 올리면 확인이 필요하다.
        """
        end = anyio.current_time() + seconds
        while (left := end - anyio.current_time()) > 0:
            engine = getattr(self, "_engine", None)
            if engine is not None and getattr(engine, "_pause_requested", False):
                return False
            await anyio.sleep(min(1.0, left))
        return True


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="퀘이사존/아카라이브 핫딜 목록 수집기")
    ap.add_argument("--site", choices=list(SITES), action="append",
                    help="수집할 사이트 (여러 번 지정 가능, 기본: 전부)")
    ap.add_argument("--pages", type=int, default=4, help="사이트별 최대 목록 페이지 수 (기본 4)")
    ap.add_argument("--interval", type=float, default=180,
                    help="다음 페이지를 요청하기 전 대기 시간, 초 (기본 180 = 3분)")
    ap.add_argument("--max-age-hours", type=float, default=24,
                    help="작성된 지 이 시간이 지난 글이 나오면 그 사이트 수집 중단 (기본 24, 0이면 제한 없음)")
    ap.add_argument("--keyword", action="append", default=[], help="제목 키워드 필터 (OR 조건)")
    ap.add_argument("-o", "--output", default="deals.json",
                    help="결과 파일 (.json/.jsonl/.csv, 기본 deals.json). .xlsx를 주면 같은 이름의 .json과 .xlsx로 저장")
    ap.add_argument("--before", help="이전 결과를 옮겨 둘 파일 (기본 before + 결과 파일 확장자, 예: before.json)")
    ap.add_argument("--no-before", action="store_true",
                    help="이전 결과를 옮기지도 비교하지도 않음 (결과 파일은 덮어씀)")
    ap.add_argument("--excel", help="엑셀 파일 (기본: 결과 파일과 같은 이름의 .xlsx)")
    ap.add_argument("--no-excel", action="store_true", help="엑셀로 변환하지 않음")
    ap.add_argument("--browser", action="store_true", help="브라우저 렌더링 사용 (scrapling install 필요)")
    ap.add_argument("--dump-html", action="store_true", help="각 사이트 1페이지 HTML을 ./html_dump 에 저장")
    args = ap.parse_args(argv)

    out = Path(args.output)
    excel = Path(args.excel) if args.excel else None
    if out.suffix.lower() == ".xlsx":       # 이전 사용법(-o deals.xlsx) 호환
        excel = excel or out
        out = out.with_suffix(".json")
    if excel is None:
        excel = out.with_suffix(".xlsx")
    before_path = Path(args.before) if args.before else out.with_name("before" + out.suffix)

    before: dict[str, Seen] = {}
    if not args.no_before:
        if out.resolve() == before_path.resolve():
            ap.error("--before 파일은 결과 파일과 이름이 달라야 합니다.")
        if out.exists():
            out.replace(before_path)        # before 파일이 이미 있으면 덮어씀
            print(f"이전 결과 {out} → {before_path}")
        before = load_before(before_path)
        if before:
            counts = ", ".join(f"{site} {len(s.ids)}건" for site, s in before.items())
            print(f"{before_path}와 비교합니다: {counts}")

    spider = HotDealSpider(
        sites=args.site or list(SITES),
        pages=max(1, args.pages),
        keywords=args.keyword,
        use_browser=args.browser,
        dump_dir=Path("html_dump") if args.dump_html else None,
        interval=max(0.0, args.interval),
        before=before,
        max_age=timedelta(hours=args.max_age_hours) if args.max_age_hours > 0 else None,
    )
    result = spider.start()

    ext = out.suffix.lower()
    if ext == ".csv":
        # 엑셀에서 한글이 깨지지 않도록 저장 후 BOM 추가
        result.items.to_csv(out)
        out.write_bytes(b"\xef\xbb\xbf" + out.read_bytes())
    elif ext == ".jsonl":
        result.items.to_jsonl(out)
    else:
        result.items.to_json(out, indent=True)
    print(f"{len(result.items)}건 저장 → {out.resolve()}")

    if not args.no_excel:
        try:
            write_xlsx(list(result.items), excel)
        except PermissionError:
            print(f"엑셀 저장 실패: {excel}이(가) 열려 있습니다. 닫은 뒤 "
                  f"'python export_excel.py {out} -o {excel}'로 변환하세요.")
            return 1
        print(f"엑셀 변환 → {excel.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

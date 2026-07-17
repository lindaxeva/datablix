from __future__ import annotations

import gzip
import ipaddress
import json
import re
import socket
import time
import xml.etree.ElementTree as ET
from collections import deque
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Callable, Iterable, Iterator, Literal
from urllib.parse import parse_qsl, urlencode, urldefrag, urljoin, urlparse, urlunparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import tldextract
except ImportError:  # pragma: no cover - handled with a clear runtime error
    tldextract = None

try:
    from playwright.sync_api import Browser, Page, Playwright, sync_playwright
except ImportError:  # Playwright is optional unless JavaScript rendering is requested.
    Browser = Page = Playwright = None
    sync_playwright = None


RenderMode = Literal["html", "auto", "javascript"]
ProgressCallback = Callable[[dict], None]

USER_AGENT = "DatablixResearchScanner/3.0 (+human-reviewed research tool)"
DEFAULT_TIMEOUT = 20
MAX_HTML_BYTES = 8_000_000
MAX_SITEMAP_BYTES = 12_000_000

TRACKING_QUERY_KEYS = {
    "fbclid",
    "gclid",
    "mc_cid",
    "mc_eid",
    "ref",
    "ref_src",
    "source",
    "utm_campaign",
    "utm_content",
    "utm_medium",
    "utm_source",
    "utm_term",
}

SKIP_EXTENSIONS = {
    ".7z", ".avi", ".bmp", ".css", ".csv", ".doc", ".docx", ".eot",
    ".epub", ".gif", ".gz", ".ico", ".jpeg", ".jpg", ".js", ".json",
    ".m4a", ".m4v", ".mov", ".mp3", ".mp4", ".mpeg", ".mpg", ".ods",
    ".odt", ".otf", ".pdf", ".png", ".ppt", ".pptx", ".rar", ".rss",
    ".svg", ".tar", ".tif", ".tiff", ".ttf", ".txt", ".wav", ".webm",
    ".webp", ".woff", ".woff2", ".xls", ".xlsx", ".xml", ".zip",
}

LOW_PRIORITY_PATH_WORDS = {
    "accessibility", "author", "blog", "careers", "category", "cookies",
    "events", "feed", "legal", "login", "news", "privacy", "register",
    "search", "signin", "signup", "tag", "terms", "wp-admin", "wp-login",
}

HIGH_PRIORITY_WORDS = {
    "apartment", "apartments", "building", "buildings", "communities",
    "community", "contact", "location", "locations", "portfolio", "properties",
    "property", "rental", "rentals", "residence", "residences", "suite", "suites",
}

STREET_SUFFIX_PATTERN = (
    r"(?:Avenue|Ave\.?|Boulevard|Blvd\.?|Circle|Court|Ct\.?|Crescent|Cres\.?|"
    r"Drive|Dr\.?|Highway|Hwy\.?|Lane|Ln\.?|Parkway|Pkwy\.?|Place|Pl\.?|"
    r"Road|Rd\.?|Route|Street|St\.?|Terrace|Ter\.?|Trail|Way)"
)
CANADIAN_POSTAL_PATTERN = r"[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ -]?\d[ABCEGHJ-NPRSTV-Z]\d"
ADDRESS_LINE_RE = re.compile(
    rf"\b\d{{1,6}}(?:[-–]\d{{1,6}})?\s+[A-Za-z0-9À-ÿ'’.,#&\- ]{{2,90}}?\s+{STREET_SUFFIX_PATTERN}\b"
    rf"(?:[^\n]{{0,100}}?\b{CANADIAN_POSTAL_PATTERN}\b)?",
    re.IGNORECASE,
)
POSTAL_RE = re.compile(rf"\b{CANADIAN_POSTAL_PATTERN}\b", re.IGNORECASE)
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE)
PHONE_RE = re.compile(
    r"(?<!\d)(?:\+?1[ .\-]?)?(?:\(?\d{3}\)?[ .\-]?)\d{3}[ .\-]?\d{4}(?:\s*(?:x|ext\.?|extension)\s*\d{1,6})?(?!\d)",
    re.IGNORECASE,
)
UNIT_COUNT_RE = re.compile(
    r"\b(\d{1,5})\s+(?:rental\s+)?(?:apartments?|residential\s+units?|units?|suites?)\b",
    re.IGNORECASE,
)
STOREY_RE = re.compile(
    r"\b(\d{1,3})\s*(?:[-–]\s*)?(?:storey|storeys|story|stories|floor|floors)\b",
    re.IGNORECASE,
)
CLASSIFICATION_PATTERNS = [
    ("High Rise", re.compile(r"\b(?:high[-\s]?rise|apartment\s+tower)\b", re.IGNORECASE)),
    ("Mid Rise", re.compile(r"\bmid[-\s]?rise\b", re.IGNORECASE)),
    ("Low Rise", re.compile(r"\blow[-\s]?rise\b", re.IGNORECASE)),
    ("Townhome", re.compile(r"\b(?:townhome|townhouse)s?\b", re.IGNORECASE)),
    ("Duplex", re.compile(r"\bduplex(?:es)?\b", re.IGNORECASE)),
    ("Garden Home", re.compile(r"\bgarden\s+(?:home|apartment)s?\b", re.IGNORECASE)),
    ("Luxury", re.compile(r"\bluxury\b", re.IGNORECASE)),
    ("Adult-oriented", re.compile(r"\badult[-\s]?oriented\b", re.IGNORECASE)),
]

SCHEMA_PROPERTY_TYPES = {
    "Accommodation", "Apartment", "ApartmentComplex", "House", "Place",
    "Residence", "SingleFamilyResidence", "RealEstateListing",
}
SCHEMA_ORG_TYPES = {
    "Corporation", "LocalBusiness", "Organization", "RealEstateAgent",
}


class WebsiteScanError(RuntimeError):
    """Raised when a website cannot be scanned safely or correctly."""


@dataclass(slots=True)
class ScanOptions:
    max_pages: int = 100
    max_depth: int = 5
    request_delay_seconds: float = 0.75
    request_timeout_seconds: int = DEFAULT_TIMEOUT
    render_mode: RenderMode = "auto"
    use_sitemaps: bool = True
    include_subdomains: bool = True
    maximum_sitemap_urls: int = 5_000
    maximum_queue_urls: int = 10_000
    follow_query_strings: bool = False
    obey_robots_txt: bool = True
    stop_after_consecutive_failures: int = 25

    def validate(self) -> None:
        if not 1 <= self.max_pages <= 2_000:
            raise WebsiteScanError("Maximum pages must be between 1 and 2,000.")
        if not 0 <= self.max_depth <= 20:
            raise WebsiteScanError("Maximum depth must be between 0 and 20.")
        if not 0.1 <= self.request_delay_seconds <= 30:
            raise WebsiteScanError("Request delay must be between 0.1 and 30 seconds.")
        if not 5 <= self.request_timeout_seconds <= 120:
            raise WebsiteScanError("Request timeout must be between 5 and 120 seconds.")
        if self.render_mode not in {"html", "auto", "javascript"}:
            raise WebsiteScanError("Unknown rendering mode.")


@dataclass(slots=True)
class PageResult:
    url: str
    final_url: str
    depth: int
    status_code: int | None
    content_type: str
    page_title: str
    heading: str
    rendered_with_javascript: bool
    discovered_links: int
    extracted_records: int
    scanned_at_utc: str
    outcome: str
    error: str = ""


@dataclass(slots=True)
class RecordCandidate:
    approved: bool = False
    building_name: str = ""
    management_owner: str = ""
    street_address: str = ""
    address_line_2: str = ""
    city: str = ""
    province: str = ""
    postal_code: str = ""
    country: str = ""
    phone: str = ""
    primary_email: str = ""
    website: str = ""
    number_of_apartments: str = ""
    building_classification: str = ""
    source_url: str = ""
    source_page_title: str = ""
    extraction_method: str = ""
    confidence: float = 0.0
    review_status: str = "Review required"
    evidence: str = ""


@dataclass(slots=True)
class ScanReport:
    start_url: str
    pages: list[PageResult] = field(default_factory=list)
    records: list[RecordCandidate] = field(default_factory=list)
    blocked_urls: list[str] = field(default_factory=list)
    skipped_urls: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    started_at_utc: str = ""
    completed_at_utc: str = ""

    def as_dict(self) -> dict:
        return {
            "start_url": self.start_url,
            "pages": [asdict(item) for item in self.pages],
            "records": [asdict(item) for item in self.records],
            "blocked_urls": self.blocked_urls,
            "skipped_urls": self.skipped_urls,
            "errors": self.errors,
            "started_at_utc": self.started_at_utc,
            "completed_at_utc": self.completed_at_utc,
        }


@dataclass(slots=True)
class FetchResult:
    requested_url: str
    final_url: str
    status_code: int
    content_type: str
    html: str
    rendered_with_javascript: bool


@dataclass(slots=True)
class RobotsRules:
    parser: RobotFileParser | None
    sitemaps: list[str]
    crawl_delay: float | None


class FullSiteScanner:
    """
    Same-site crawler for public, permitted pages.

    It supports:
    - robots.txt checks
    - XML sitemap and sitemap-index discovery
    - breadth-first internal-link crawling
    - optional JavaScript rendering with Playwright
    - JSON-LD and HTML extraction
    - conservative request delays and bounded crawling
    - SSRF protection against private/local network targets
    """

    def __init__(
        self,
        start_url: str,
        options: ScanOptions | None = None,
        progress_callback: ProgressCallback | None = None,
    ) -> None:
        self.options = options or ScanOptions()
        self.options.validate()
        self.progress_callback = progress_callback
        self.start_url = self._normalize_start_url(start_url)
        self._ensure_public_url(self.start_url)

        parsed = urlparse(self.start_url)
        self.start_origin = f"{parsed.scheme}://{parsed.netloc}"
        self.start_host = parsed.hostname or ""
        self.registered_domain = self._registered_domain(self.start_host)

        self.session = self._build_session()
        self.robots_cache: dict[str, RobotsRules] = {}
        self.playwright_context = None
        self.browser: Browser | None = None
        self.page: Page | None = None
        self._last_request_monotonic = 0.0

    def scan(self) -> ScanReport:
        report = ScanReport(
            start_url=self.start_url,
            started_at_utc=self._utc_now(),
        )

        queue: deque[tuple[str, int, int]] = deque()
        queued: set[str] = set()
        visited: set[str] = set()

        self._enqueue(queue, queued, self.start_url, depth=0, priority=0)

        if self.options.use_sitemaps:
            for sitemap_url in self._discover_sitemap_urls(self.start_url, report):
                for page_url in self._read_sitemap_tree(sitemap_url, report):
                    priority = self._url_priority(page_url)
                    self._enqueue(queue, queued, page_url, depth=0, priority=priority)
                    if len(queued) >= self.options.maximum_queue_urls:
                        break
                if len(queued) >= self.options.maximum_queue_urls:
                    break

        # Deque insertion above is simple and stable. Sort once so likely property pages
        # discovered from sitemaps are visited early, while still crawling the whole site.
        queue = deque(sorted(queue, key=lambda item: (item[2], item[1], item[0])))
        consecutive_failures = 0

        self._start_browser_if_needed()

        try:
            while queue and len(report.pages) < self.options.max_pages:
                current_url, depth, _priority = queue.popleft()
                current_url = self._canonicalize_url(current_url)

                if current_url in visited:
                    continue
                visited.add(current_url)

                if not self._is_in_scope(current_url) or self._should_skip_url(current_url):
                    report.skipped_urls.append(current_url)
                    continue

                if self.options.obey_robots_txt and not self._robots_allows(current_url):
                    report.blocked_urls.append(current_url)
                    self._notify(report, current_url, "Blocked by robots.txt")
                    continue

                try:
                    fetch = self._fetch_page(current_url)
                    consecutive_failures = 0
                except Exception as exc:  # continue crawling after page-specific failures
                    consecutive_failures += 1
                    error_message = f"{current_url}: {exc}"
                    report.errors.append(error_message)
                    report.pages.append(
                        PageResult(
                            url=current_url,
                            final_url=current_url,
                            depth=depth,
                            status_code=None,
                            content_type="",
                            page_title="",
                            heading="",
                            rendered_with_javascript=False,
                            discovered_links=0,
                            extracted_records=0,
                            scanned_at_utc=self._utc_now(),
                            outcome="Error",
                            error=str(exc),
                        )
                    )
                    self._notify(report, current_url, "Error")
                    if consecutive_failures >= self.options.stop_after_consecutive_failures:
                        report.errors.append(
                            "Scan stopped after too many consecutive page failures."
                        )
                        break
                    continue

                soup = self._make_soup(fetch.html)
                page_title = self._page_title(soup)
                heading = self._page_heading(soup)

                records = self._extract_records(fetch.final_url, soup, page_title)
                report.records.extend(records)

                links = self._extract_internal_links(fetch.final_url, soup)
                if depth < self.options.max_depth:
                    for link_url in links:
                        self._enqueue(
                            queue,
                            queued,
                            link_url,
                            depth=depth + 1,
                            priority=self._url_priority(link_url),
                        )

                report.pages.append(
                    PageResult(
                        url=current_url,
                        final_url=fetch.final_url,
                        depth=depth,
                        status_code=fetch.status_code,
                        content_type=fetch.content_type,
                        page_title=page_title,
                        heading=heading,
                        rendered_with_javascript=fetch.rendered_with_javascript,
                        discovered_links=len(links),
                        extracted_records=len(records),
                        scanned_at_utc=self._utc_now(),
                        outcome="Scanned",
                    )
                )
                self._notify(report, fetch.final_url, "Scanned")

            report.records = self._deduplicate_records(report.records)
            report.completed_at_utc = self._utc_now()
            return report
        finally:
            self.close()

    def _close_browser(self) -> None:
        if self.page is not None:
            try:
                self.page.close()
            except Exception:
                pass
            self.page = None
        if self.browser is not None:
            try:
                self.browser.close()
            except Exception:
                pass
            self.browser = None
        if self.playwright_context is not None:
            try:
                self.playwright_context.stop()
            except Exception:
                pass
            self.playwright_context = None

    def close(self) -> None:
        self._close_browser()
        self.session.close()

    # ------------------------------------------------------------------
    # Fetching, rendering, robots, and sitemap discovery
    # ------------------------------------------------------------------

    @staticmethod
    def _build_session() -> requests.Session:
        session = requests.Session()
        retry = Retry(
            total=2,
            connect=2,
            read=2,
            status=2,
            backoff_factor=0.75,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset({"GET", "HEAD"}),
            respect_retry_after_header=True,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        session.headers.update(
            {
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.5",
                "Accept-Language": "en-CA,en;q=0.9,fr-CA;q=0.7,fr;q=0.6",
            }
        )
        return session

    def _fetch_page(self, url: str) -> FetchResult:
        self._ensure_public_url(url)
        self._respect_delay(url)

        response = self.session.get(
            url,
            timeout=self.options.request_timeout_seconds,
            allow_redirects=True,
            stream=True,
        )
        response.raise_for_status()

        final_url = self._canonicalize_url(response.url)
        self._ensure_public_url(final_url)
        if not self._is_in_scope(final_url):
            raise WebsiteScanError("Redirected outside the permitted website scope.")

        content_type = response.headers.get("Content-Type", "").split(";", 1)[0].lower()
        if content_type and content_type not in {"text/html", "application/xhtml+xml"}:
            raise WebsiteScanError(f"Unsupported content type: {content_type}")

        raw = self._read_limited_response(response, MAX_HTML_BYTES)
        encoding = response.encoding or response.apparent_encoding or "utf-8"
        html = raw.decode(encoding, errors="replace")
        rendered = False

        should_render = self.options.render_mode == "javascript" or (
            self.options.render_mode == "auto" and self._looks_javascript_dependent(html)
        )
        if should_render and self.page is not None:
            html, final_url = self._render_with_playwright(final_url)
            rendered = True
        elif should_render and self.options.render_mode == "javascript":
            raise WebsiteScanError("JavaScript browser rendering is unavailable.")

        return FetchResult(
            requested_url=url,
            final_url=self._canonicalize_url(final_url),
            status_code=response.status_code,
            content_type=content_type or "text/html",
            html=html,
            rendered_with_javascript=rendered,
        )

    @staticmethod
    def _read_limited_response(response: requests.Response, limit: int) -> bytes:
        chunks: list[bytes] = []
        total = 0
        for chunk in response.iter_content(chunk_size=64 * 1024):
            if not chunk:
                continue
            total += len(chunk)
            if total > limit:
                raise WebsiteScanError("Page exceeded the configured size limit.")
            chunks.append(chunk)
        return b"".join(chunks)

    def _looks_javascript_dependent(self, html: str) -> bool:
        soup = self._make_soup(html)
        visible_text = self._visible_text(soup)
        links = soup.find_all("a", href=True)
        script_count = len(soup.find_all("script"))
        app_markers = bool(
            soup.select_one("#__next, #root, #app, [data-reactroot], [ng-version]")
        )
        hydration_markers = any(
            marker in html
            for marker in ("__NEXT_DATA__", "__NUXT__", "webpackJsonp", "hydrateRoot")
        )
        return (
            (len(visible_text) < 350 and script_count >= 4)
            or (len(links) < 3 and script_count >= 6)
            or (app_markers and len(visible_text) < 1_000)
            or hydration_markers
        )

    def _start_browser_if_needed(self) -> None:
        if self.options.render_mode == "html":
            return
        if sync_playwright is None:
            if self.options.render_mode == "javascript":
                raise WebsiteScanError(
                    "JavaScript rendering was selected, but Playwright is not installed."
                )
            return

        try:
            self.playwright_context = sync_playwright().start()
            self.browser = self.playwright_context.chromium.launch(headless=True)
            self.page = self.browser.new_page(
                user_agent=USER_AGENT,
                locale="en-CA",
                viewport={"width": 1440, "height": 1000},
            )
            self.page.route(
                "**/*",
                lambda route: route.abort()
                if route.request.resource_type in {"image", "media", "font"}
                else route.continue_(),
            )
        except Exception as exc:
            self._close_browser()
            if self.options.render_mode == "javascript":
                raise WebsiteScanError(
                    "Playwright could not start. Install its Chromium browser binary."
                ) from exc

    def _render_with_playwright(self, url: str) -> tuple[str, str]:
        if self.page is None:
            raise WebsiteScanError("JavaScript browser rendering is unavailable.")

        self._ensure_public_url(url)
        response = self.page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=self.options.request_timeout_seconds * 1_000,
        )
        try:
            self.page.wait_for_load_state("networkidle", timeout=6_000)
        except Exception:
            pass

        final_url = self._canonicalize_url(self.page.url)
        self._ensure_public_url(final_url)
        if not self._is_in_scope(final_url):
            raise WebsiteScanError("Browser navigation left the permitted website scope.")

        if response is not None and response.status >= 400:
            raise WebsiteScanError(f"Browser returned HTTP {response.status}.")

        return self.page.content(), final_url

    def _robots_for_url(self, url: str) -> RobotsRules:
        parsed = urlparse(url)
        origin = f"{parsed.scheme}://{parsed.netloc}"
        if origin in self.robots_cache:
            return self.robots_cache[origin]

        robots_url = urljoin(origin, "/robots.txt")
        parser: RobotFileParser | None = None
        sitemaps: list[str] = []
        crawl_delay: float | None = None

        try:
            self._ensure_public_url(robots_url)
            response = self.session.get(
                robots_url,
                timeout=self.options.request_timeout_seconds,
                allow_redirects=True,
            )
            if response.ok and len(response.content) <= 1_000_000:
                text = response.text
                parser = RobotFileParser()
                parser.set_url(robots_url)
                parser.parse(text.splitlines())
                crawl_delay = parser.crawl_delay(USER_AGENT) or parser.crawl_delay("*")
                for line in text.splitlines():
                    if line.lower().startswith("sitemap:"):
                        candidate = line.split(":", 1)[1].strip()
                        if candidate:
                            sitemaps.append(urljoin(origin, candidate))
            elif response.status_code in {401, 403}:
                parser = RobotFileParser()
                parser.set_url(robots_url)
                parser.parse(["User-agent: *", "Disallow: /"])
        except requests.RequestException:
            parser = None

        rules = RobotsRules(parser=parser, sitemaps=list(dict.fromkeys(sitemaps)), crawl_delay=crawl_delay)
        self.robots_cache[origin] = rules
        return rules

    def _robots_allows(self, url: str) -> bool:
        rules = self._robots_for_url(url)
        if rules.parser is None:
            return True
        return rules.parser.can_fetch(USER_AGENT, url)

    def _discover_sitemap_urls(self, start_url: str, report: ScanReport) -> list[str]:
        parsed = urlparse(start_url)
        origin = f"{parsed.scheme}://{parsed.netloc}"
        rules = self._robots_for_url(start_url)
        candidates = list(rules.sitemaps)
        candidates.extend(
            [
                urljoin(origin, "/sitemap.xml"),
                urljoin(origin, "/sitemap_index.xml"),
                urljoin(origin, "/sitemap-index.xml"),
                urljoin(origin, "/wp-sitemap.xml"),
            ]
        )
        valid: list[str] = []
        for url in dict.fromkeys(candidates):
            if self._is_in_scope(url) and not self._should_skip_url(url, allow_xml=True):
                valid.append(url)
        return valid

    def _read_sitemap_tree(self, sitemap_url: str, report: ScanReport) -> Iterator[str]:
        pending = deque([sitemap_url])
        seen_sitemaps: set[str] = set()
        yielded_urls = 0

        while pending and yielded_urls < self.options.maximum_sitemap_urls:
            current = self._canonicalize_url(pending.popleft(), keep_query=True)
            if current in seen_sitemaps:
                continue
            seen_sitemaps.add(current)

            try:
                self._ensure_public_url(current)
                if self.options.obey_robots_txt and not self._robots_allows(current):
                    report.blocked_urls.append(current)
                    continue
                self._respect_delay(current)
                response = self.session.get(
                    current,
                    timeout=self.options.request_timeout_seconds,
                    allow_redirects=True,
                )
                response.raise_for_status()
                data = response.content
                if len(data) > MAX_SITEMAP_BYTES:
                    raise WebsiteScanError("Sitemap exceeded the size limit.")
                if current.lower().endswith(".gz") or data[:2] == b"\x1f\x8b":
                    data = gzip.decompress(data)
                root = ET.fromstring(data)
            except Exception as exc:
                # Common sitemap guesses often do not exist, so keep these errors quiet.
                if current in self._robots_for_url(self.start_url).sitemaps:
                    report.errors.append(f"Could not read declared sitemap {current}: {exc}")
                continue

            root_name = self._xml_local_name(root.tag)
            if root_name == "sitemapindex":
                for loc in self._xml_locations(root):
                    loc = urljoin(current, loc)
                    if self._is_in_scope(loc):
                        pending.append(loc)
            elif root_name == "urlset":
                for loc in self._xml_locations(root):
                    loc = self._canonicalize_url(urljoin(current, loc))
                    if not self._is_in_scope(loc) or self._should_skip_url(loc):
                        continue
                    yielded_urls += 1
                    yield loc
                    if yielded_urls >= self.options.maximum_sitemap_urls:
                        break

    @staticmethod
    def _xml_local_name(tag: str) -> str:
        return tag.split("}", 1)[-1].lower()

    def _xml_locations(self, root: ET.Element) -> Iterator[str]:
        for element in root.iter():
            if self._xml_local_name(element.tag) == "loc" and element.text:
                yield element.text.strip()

    def _respect_delay(self, url: str) -> None:
        robots_delay = self._robots_for_url(url).crawl_delay
        delay = max(self.options.request_delay_seconds, robots_delay or 0.0)
        elapsed = time.monotonic() - self._last_request_monotonic
        if elapsed < delay:
            time.sleep(delay - elapsed)
        self._last_request_monotonic = time.monotonic()

    # ------------------------------------------------------------------
    # URL normalization, scope, and SSRF protection
    # ------------------------------------------------------------------

    @staticmethod
    def _normalize_start_url(value: str) -> str:
        value = value.strip()
        if not value:
            raise WebsiteScanError("Enter a website address.")
        if not value.startswith(("http://", "https://")):
            value = f"https://{value}"
        parsed = urlparse(value)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            raise WebsiteScanError("Enter a valid HTTP or HTTPS website address.")
        if parsed.username or parsed.password:
            raise WebsiteScanError("Website addresses containing credentials are not supported.")
        return FullSiteScanner._canonicalize_url(value)

    @staticmethod
    def _canonicalize_url(url: str, keep_query: bool = False) -> str:
        url, _fragment = urldefrag(url.strip())
        parsed = urlparse(url)
        scheme = parsed.scheme.lower()
        hostname = (parsed.hostname or "").lower()
        if not hostname:
            return url

        port = parsed.port
        if port and not ((scheme == "http" and port == 80) or (scheme == "https" and port == 443)):
            netloc = f"{hostname}:{port}"
        else:
            netloc = hostname

        path = re.sub(r"/{2,}", "/", parsed.path or "/")
        if path != "/" and path.endswith("/"):
            path = path[:-1]

        query = ""
        if keep_query:
            query = parsed.query
        elif parsed.query:
            filtered = [
                (key, value)
                for key, value in parse_qsl(parsed.query, keep_blank_values=True)
                if key.lower() not in TRACKING_QUERY_KEYS
            ]
            query = urlencode(sorted(filtered))

        return urlunparse((scheme, netloc, path, "", query, ""))

    def _registered_domain(self, hostname: str) -> str:
        if tldextract is None:
            raise WebsiteScanError(
                "The tldextract package is required for safe same-site crawling."
            )
        extracted = tldextract.TLDExtract(suffix_list_urls=())(hostname)
        registered = extracted.top_domain_under_public_suffix
        return registered or hostname

    def _is_in_scope(self, url: str) -> bool:
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            return False
        hostname = parsed.hostname.lower()
        if self.options.include_subdomains:
            return self._registered_domain(hostname) == self.registered_domain
        return hostname == self.start_host

    @staticmethod
    def _ensure_public_url(url: str) -> None:
        parsed = urlparse(url)
        hostname = parsed.hostname
        if parsed.scheme not in {"http", "https"} or not hostname:
            raise WebsiteScanError("Only valid public HTTP and HTTPS URLs can be scanned.")
        if hostname.lower() in {"localhost", "localhost.localdomain"}:
            raise WebsiteScanError("Local network addresses cannot be scanned.")

        try:
            addresses = socket.getaddrinfo(hostname, parsed.port or 443, type=socket.SOCK_STREAM)
        except socket.gaierror as exc:
            raise WebsiteScanError("The website hostname could not be resolved.") from exc

        for address in addresses:
            ip_text = address[4][0].split("%", 1)[0]
            try:
                ip_value = ipaddress.ip_address(ip_text)
            except ValueError:
                continue
            if not ip_value.is_global:
                raise WebsiteScanError(
                    "Private, loopback, link-local, multicast, and reserved addresses cannot be scanned."
                )

    def _should_skip_url(self, url: str, allow_xml: bool = False) -> bool:
        parsed = urlparse(url)
        path = parsed.path.lower()
        extension = "." + path.rsplit(".", 1)[-1] if "." in path.rsplit("/", 1)[-1] else ""
        if extension in SKIP_EXTENSIONS and not (allow_xml and extension in {".xml", ".gz"}):
            return True
        if not self.options.follow_query_strings and parsed.query:
            return True
        if any(word in path for word in LOW_PRIORITY_PATH_WORDS):
            return True
        return False

    def _enqueue(
        self,
        queue: deque[tuple[str, int, int]],
        queued: set[str],
        url: str,
        depth: int,
        priority: int,
    ) -> None:
        if len(queued) >= self.options.maximum_queue_urls:
            return
        normalized = self._canonicalize_url(url)
        if normalized in queued or not self._is_in_scope(normalized):
            return
        if self._should_skip_url(normalized):
            return
        queued.add(normalized)
        # High-priority links go to the front; all others preserve BFS behavior.
        item = (normalized, depth, priority)
        if priority < 0:
            queue.appendleft(item)
        else:
            queue.append(item)

    @staticmethod
    def _url_priority(url: str) -> int:
        value = url.lower()
        if any(word in value for word in HIGH_PRIORITY_WORDS):
            return -1
        if any(word in value for word in LOW_PRIORITY_PATH_WORDS):
            return 2
        return 0

    # ------------------------------------------------------------------
    # HTML parsing and extraction
    # ------------------------------------------------------------------

    @staticmethod
    def _make_soup(html: str) -> BeautifulSoup:
        return BeautifulSoup(html, "lxml")

    @staticmethod
    def _clean_text(value: str | None) -> str:
        if not value:
            return ""
        return re.sub(r"\s+", " ", value).strip(" \t\r\n|–—-")

    def _visible_text(self, soup: BeautifulSoup) -> str:
        cloned = BeautifulSoup(str(soup), "lxml")
        for node in cloned(["script", "style", "noscript", "svg", "template"]):
            node.decompose()
        return self._clean_text(cloned.get_text(" ", strip=True))

    def _page_title(self, soup: BeautifulSoup) -> str:
        title = self._clean_text(soup.title.get_text(" ", strip=True) if soup.title else "")
        return title[:300]

    def _page_heading(self, soup: BeautifulSoup) -> str:
        heading = soup.find("h1") or soup.find("h2")
        return self._clean_text(heading.get_text(" ", strip=True) if heading else "")[:300]

    def _extract_internal_links(self, base_url: str, soup: BeautifulSoup) -> list[str]:
        links: set[str] = set()
        for tag in soup.find_all("a", href=True):
            href = self._clean_text(tag.get("href"))
            if not href or href.startswith(("#", "mailto:", "tel:", "javascript:", "data:")):
                continue
            candidate = self._canonicalize_url(urljoin(base_url, href))
            if self._is_in_scope(candidate) and not self._should_skip_url(candidate):
                links.add(candidate)
        return sorted(links, key=lambda value: (self._url_priority(value), value))

    def _extract_records(
        self,
        page_url: str,
        soup: BeautifulSoup,
        page_title: str,
    ) -> list[RecordCandidate]:
        records: list[RecordCandidate] = []
        records.extend(self._records_from_json_ld(page_url, soup, page_title))
        records.extend(self._records_from_microdata(page_url, soup, page_title))
        records.extend(self._records_from_address_elements(page_url, soup, page_title))

        if not any(record.street_address for record in records):
            records.extend(self._records_from_visible_text(page_url, soup, page_title))

        return self._deduplicate_records(records)

    def _records_from_json_ld(
        self,
        page_url: str,
        soup: BeautifulSoup,
        page_title: str,
    ) -> list[RecordCandidate]:
        output: list[RecordCandidate] = []
        organization_name = ""

        nodes: list[dict] = []
        for script in soup.find_all("script", attrs={"type": re.compile(r"ld\+json", re.I)}):
            raw = script.string or script.get_text(" ", strip=True)
            if not raw:
                continue
            raw = raw.strip().lstrip("\ufeff")
            try:
                parsed = json.loads(raw)
            except json.JSONDecodeError:
                # Some sites wrap otherwise-valid JSON-LD in HTML comments.
                cleaned = re.sub(r"^\s*<!--|-->\s*$", "", raw).strip()
                try:
                    parsed = json.loads(cleaned)
                except json.JSONDecodeError:
                    continue
            nodes.extend(self._flatten_json_ld(parsed))

        for node in nodes:
            node_types = self._schema_types(node.get("@type"))
            if node_types & SCHEMA_ORG_TYPES and not organization_name:
                organization_name = self._clean_text(node.get("name"))

        for node in nodes:
            node_types = self._schema_types(node.get("@type"))
            address = node.get("address") if isinstance(node.get("address"), dict) else {}
            has_property_type = bool(node_types & SCHEMA_PROPERTY_TYPES)
            has_address = bool(address or node.get("streetAddress"))
            if not has_property_type and not has_address:
                continue

            parent_org = node.get("provider") or node.get("seller") or node.get("brand") or {}
            if isinstance(parent_org, str):
                parent_org_name = parent_org
            elif isinstance(parent_org, dict):
                parent_org_name = parent_org.get("name", "")
            else:
                parent_org_name = ""

            record = RecordCandidate(
                building_name=self._clean_text(node.get("name")),
                management_owner=self._clean_text(parent_org_name or organization_name),
                street_address=self._clean_text(address.get("streetAddress") or node.get("streetAddress")),
                address_line_2=self._clean_text(address.get("postOfficeBoxNumber")),
                city=self._clean_text(address.get("addressLocality")),
                province=self._clean_text(address.get("addressRegion")),
                postal_code=self._normalize_postal(address.get("postalCode")),
                country=self._country_value(address.get("addressCountry")),
                phone=self._clean_text(node.get("telephone")),
                primary_email=self._clean_text(node.get("email")),
                website=self._clean_text(node.get("url") or page_url),
                number_of_apartments=self._unit_count_from_node(node),
                building_classification=self._classification_from_text(
                    " ".join(
                        self._clean_text(str(node.get(key, "")))
                        for key in ("@type", "additionalType", "description", "keywords")
                    )
                ),
                source_url=page_url,
                source_page_title=page_title,
                extraction_method="JSON-LD structured data",
                confidence=0.95 if has_address else 0.80,
                evidence=self._evidence_text(node),
            )
            output.append(record)
        return output

    def _flatten_json_ld(self, value) -> list[dict]:
        output: list[dict] = []
        if isinstance(value, list):
            for item in value:
                output.extend(self._flatten_json_ld(item))
        elif isinstance(value, dict):
            graph = value.get("@graph")
            if isinstance(graph, list):
                output.extend(self._flatten_json_ld(graph))
            output.append(value)
            for key in ("itemListElement", "mainEntity", "about"):
                child = value.get(key)
                if isinstance(child, (list, dict)):
                    output.extend(self._flatten_json_ld(child))
        return output

    @staticmethod
    def _schema_types(value) -> set[str]:
        if isinstance(value, str):
            return {value.rsplit("/", 1)[-1]}
        if isinstance(value, list):
            return {str(item).rsplit("/", 1)[-1] for item in value}
        return set()

    def _unit_count_from_node(self, node: dict) -> str:
        for key in ("numberOfRooms", "numberOfUnits", "numberOfAccommodationUnits"):
            value = node.get(key)
            if value not in (None, ""):
                return self._clean_text(str(value))
        description = self._clean_text(node.get("description"))
        match = UNIT_COUNT_RE.search(description)
        return match.group(1) if match else ""

    def _classification_from_text(self, text: str) -> str:
        """Extract classification wording that is actually present on the page."""
        clean_text = self._clean_text(text)
        if not clean_text:
            return ""

        labels = [
            label
            for label, pattern in CLASSIFICATION_PATTERNS
            if pattern.search(clean_text)
        ]
        labels = list(dict.fromkeys(labels))
        storey_match = STOREY_RE.search(clean_text)
        storeys = storey_match.group(1) if storey_match else ""

        if labels:
            classification = " | ".join(labels)
            if storeys and len(labels) == 1:
                classification = f"{classification} - {storeys}"
            return classification
        if storeys:
            return f"{storeys} Storeys"
        return ""

    def _evidence_text(self, node: dict) -> str:
        evidence_parts = []
        for key in ("name", "description", "telephone", "email"):
            value = self._clean_text(str(node.get(key, "")))
            if value:
                evidence_parts.append(value)
        return " | ".join(evidence_parts)[:700]

    def _records_from_microdata(
        self,
        page_url: str,
        soup: BeautifulSoup,
        page_title: str,
    ) -> list[RecordCandidate]:
        output: list[RecordCandidate] = []
        for block in soup.select("[itemscope]"):
            itemtype = self._clean_text(block.get("itemtype"))
            if not any(schema_type.lower() in itemtype.lower() for schema_type in SCHEMA_PROPERTY_TYPES):
                continue

            def prop(name: str) -> str:
                node = block.select_one(f'[itemprop="{name}"]')
                if node is None:
                    return ""
                return self._clean_text(
                    node.get("content") or node.get("href") or node.get_text(" ", strip=True)
                )

            street = prop("streetAddress")
            if not street:
                continue
            output.append(
                RecordCandidate(
                    building_name=prop("name"),
                    street_address=street,
                    city=prop("addressLocality"),
                    province=prop("addressRegion"),
                    postal_code=self._normalize_postal(prop("postalCode")),
                    country=prop("addressCountry"),
                    phone=prop("telephone"),
                    primary_email=prop("email").replace("mailto:", ""),
                    website=prop("url") or page_url,
                    building_classification=self._classification_from_text(
                        block.get_text(" ", strip=True)
                    ),
                    source_url=page_url,
                    source_page_title=page_title,
                    extraction_method="HTML microdata",
                    confidence=0.90,
                    evidence=self._clean_text(block.get_text(" ", strip=True))[:700],
                )
            )
        return output

    def _records_from_address_elements(
        self,
        page_url: str,
        soup: BeautifulSoup,
        page_title: str,
    ) -> list[RecordCandidate]:
        output: list[RecordCandidate] = []
        organization_name = self._organization_name(soup)
        for address_tag in soup.find_all("address"):
            text = self._clean_text(address_tag.get_text(" ", strip=True))
            if not text or not re.search(r"\d", text):
                continue
            heading = self._nearest_heading(address_tag) or self._page_heading(soup)
            street, city, province, postal = self._split_address(text)
            local_context = self._clean_text(
                (address_tag.parent or address_tag).get_text(" ", strip=True)
            )[:2500]
            output.append(
                RecordCandidate(
                    building_name=heading,
                    management_owner=organization_name,
                    street_address=street or text,
                    city=city,
                    province=province,
                    postal_code=postal,
                    country="Canada" if postal else "",
                    phone=self._first_match(PHONE_RE, text),
                    primary_email=self._first_match(EMAIL_RE, text),
                    website=page_url,
                    number_of_apartments=self._first_group(UNIT_COUNT_RE, local_context),
                    building_classification=self._classification_from_text(local_context),
                    source_url=page_url,
                    source_page_title=page_title,
                    extraction_method="HTML address element",
                    confidence=0.82,
                    evidence=text[:700],
                )
            )
        return output

    def _records_from_visible_text(
        self,
        page_url: str,
        soup: BeautifulSoup,
        page_title: str,
    ) -> list[RecordCandidate]:
        output: list[RecordCandidate] = []
        organization_name = self._organization_name(soup)
        page_heading = self._page_heading(soup)
        page_text = self._visible_text(soup)
        phones = list(dict.fromkeys(PHONE_RE.findall(page_text)))
        emails = list(dict.fromkeys(EMAIL_RE.findall(page_text)))
        unit_count = self._first_group(UNIT_COUNT_RE, page_text)

        seen_addresses: set[str] = set()
        for match in ADDRESS_LINE_RE.finditer(page_text):
            raw_address = self._clean_text(match.group(0))
            key = self._record_key("", raw_address)
            if not raw_address or key in seen_addresses:
                continue
            seen_addresses.add(key)
            street, city, province, postal = self._split_address(raw_address)
            local_context = page_text[max(0, match.start() - 600):match.end() + 600]
            local_unit_count = self._first_group(UNIT_COUNT_RE, local_context) or unit_count
            output.append(
                RecordCandidate(
                    building_name=page_heading,
                    management_owner=organization_name,
                    street_address=street or raw_address,
                    city=city,
                    province=province,
                    postal_code=postal,
                    country="Canada" if postal else "",
                    phone=phones[0] if phones else "",
                    primary_email=emails[0] if emails else "",
                    website=page_url,
                    number_of_apartments=local_unit_count,
                    building_classification=self._classification_from_text(local_context),
                    source_url=page_url,
                    source_page_title=page_title,
                    extraction_method="Visible-text pattern",
                    confidence=0.62 if postal else 0.52,
                    evidence=raw_address[:700],
                )
            )
        return output

    def _organization_name(self, soup: BeautifulSoup) -> str:
        for selector in (
            '[itemprop="legalName"]', '[itemprop="name"]',
            'meta[property="og:site_name"]', 'meta[name="application-name"]',
        ):
            node = soup.select_one(selector)
            if node:
                value = self._clean_text(node.get("content") or node.get_text(" ", strip=True))
                if value:
                    return value[:250]
        title = self._page_title(soup)
        if " | " in title:
            return self._clean_text(title.rsplit(" | ", 1)[-1])
        return ""

    def _nearest_heading(self, tag: Tag) -> str:
        previous = tag.find_previous(["h1", "h2", "h3", "h4"])
        return self._clean_text(previous.get_text(" ", strip=True) if previous else "")

    def _split_address(self, text: str) -> tuple[str, str, str, str]:
        text = self._clean_text(text)
        postal_match = POSTAL_RE.search(text)
        postal = self._normalize_postal(postal_match.group(0) if postal_match else "")
        before_postal = text[: postal_match.start()].strip(" ,") if postal_match else text

        province = ""
        province_match = re.search(
            r"(?:,|\s)\s*(ON|Ontario|QC|Quebec|Québec|BC|Alberta|AB|MB|Manitoba|NB|NS|PE|NL|SK)\s*$",
            before_postal,
            re.IGNORECASE,
        )
        if province_match:
            province = province_match.group(1)
            before_postal = before_postal[: province_match.start()].strip(" ,")

        segments = [self._clean_text(part) for part in re.split(r"\s*,\s*", before_postal) if self._clean_text(part)]
        street = segments[0] if segments else before_postal
        city = segments[-1] if len(segments) >= 2 else ""
        return street, city, province, postal

    @staticmethod
    def _normalize_postal(value) -> str:
        value = re.sub(r"\s+", "", str(value or "")).upper()
        if len(value) == 6:
            return f"{value[:3]} {value[3:]}"
        return value

    def _country_value(self, value) -> str:
        if isinstance(value, dict):
            value = value.get("name") or value.get("@id") or ""
        return self._clean_text(str(value or ""))

    @staticmethod
    def _first_match(pattern: re.Pattern, text: str) -> str:
        match = pattern.search(text)
        return match.group(0) if match else ""

    @staticmethod
    def _first_group(pattern: re.Pattern, text: str) -> str:
        match = pattern.search(text)
        return match.group(1) if match else ""

    def _deduplicate_records(self, records: Iterable[RecordCandidate]) -> list[RecordCandidate]:
        winners: dict[str, RecordCandidate] = {}
        for record in records:
            record.website = record.website or record.source_url
            record.postal_code = self._normalize_postal(record.postal_code)
            key = self._record_key(record.building_name, record.street_address or record.source_url)
            existing = winners.get(key)
            if existing is None:
                winners[key] = record
                continue
            winners[key] = self._merge_records(existing, record)
        return sorted(
            winners.values(),
            key=lambda item: (
                self._clean_text(item.building_name).lower(),
                self._clean_text(item.street_address).lower(),
                item.source_url,
            ),
        )

    def _record_key(self, name: str, address_or_url: str) -> str:
        normalized_name = re.sub(r"[^a-z0-9]", "", self._clean_text(name).lower())
        normalized_location = re.sub(r"[^a-z0-9]", "", self._clean_text(address_or_url).lower())
        return f"{normalized_name}|{normalized_location}"

    @staticmethod
    def _merge_records(first: RecordCandidate, second: RecordCandidate) -> RecordCandidate:
        preferred = first if first.confidence >= second.confidence else second
        other = second if preferred is first else first
        for field_name in (
            "building_name", "management_owner", "street_address", "address_line_2",
            "city", "province", "postal_code", "country", "phone", "primary_email",
            "website", "number_of_apartments", "building_classification",
            "source_url", "source_page_title",
            "extraction_method", "evidence",
        ):
            if not getattr(preferred, field_name) and getattr(other, field_name):
                setattr(preferred, field_name, getattr(other, field_name))
        preferred.confidence = max(first.confidence, second.confidence)
        return preferred

    # ------------------------------------------------------------------
    # Progress reporting
    # ------------------------------------------------------------------

    def _notify(self, report: ScanReport, current_url: str, outcome: str) -> None:
        if self.progress_callback is None:
            return
        self.progress_callback(
            {
                "current_url": current_url,
                "outcome": outcome,
                "pages_processed": len(report.pages),
                "records_found": len(report.records),
                "blocked_count": len(report.blocked_urls),
                "error_count": len(report.errors),
                "max_pages": self.options.max_pages,
            }
        )

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(timezone.utc).isoformat(timespec="seconds")


def scan_website(
    website_url: str,
    options: ScanOptions | None = None,
    progress_callback: ProgressCallback | None = None,
) -> ScanReport:
    """Convenience entry point used by the Streamlit interface."""
    scanner = FullSiteScanner(
        start_url=website_url,
        options=options,
        progress_callback=progress_callback,
    )
    return scanner.scan()

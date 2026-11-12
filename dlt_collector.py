#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DLT Statistics Excel Collector — interactive terminal application.

Install: python -m pip install -r requirements.txt
Run:     python dlt_collector.py
"""

from __future__ import annotations

import csv
import hashlib
import random
import re
import threading
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

try:
    import truststore

    truststore.inject_into_ssl()
except ImportError:
    pass

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from textual import work
from textual.app import App, ComposeResult
from textual.containers import Horizontal, Vertical
from textual.widgets import (
    Button,
    Checkbox,
    DataTable,
    Footer,
    Header,
    Input,
    ProgressBar,
    RichLog,
    Select,
    Static,
)
from urllib3.util.retry import Retry

BASE_URL = "https://web.dlt.go.th/statistics/"
DOWNLOAD_URL = BASE_URL + "load_file_select_new_car.php"

CATEGORIES = {
    "1.1": ("Registration & Tax", "m1_stat_car.php", 4, "01_Registration_Tax"),
    "1.2": ("Vehicle Registration", "m1_new_car.php", 1, "02_Vehicle_Registration"),
    "1.3": ("New Registration", "m1_new_car_red.php", 2, "03_New_Registration"),
    "1.4": ("Accumulated Registration", "m1_collect_new_car.php", 3, "04_Accumulated_Registration"),
}

MANIFEST_COLUMNS = [
    "timestamp", "category", "topic", "subtopic", "dropdown",
    "data_file", "source_url", "final_url", "file_path",
    "file_size", "sha256", "status", "error",
]
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
INVALID_PATH = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class Job:
    category: str
    topic: str
    subtopic: str
    dropdown: str
    data_file: str
    t: int
    folder: Path
    source_type: str = "dropdown"

    @property
    def title(self) -> str:
        return f"{self.topic} / {self.subtopic}" if self.subtopic else self.topic

    @property
    def key(self) -> str:
        return f"{self.category}:{self.t}:{self.data_file}"


class DownloadStopped(Exception):
    pass


def safe_name(text: str, max_length: int = 50) -> str:
    text = re.sub(r"\s+", " ", str(text)).strip().lstrip("- ")
    text = INVALID_PATH.sub("_", text).rstrip(" .")
    return text[:max_length].rstrip(" .") or "unknown"


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def filename_from_disposition(header: str) -> str | None:
    match = re.search(r"filename\*\s*=\s*(?:UTF-8'')?([^;]+)", header, re.I)
    if match:
        return re.split(r"[/\\]", unquote(match.group(1).strip().strip('"')))[-1]
    match = re.search(r'filename\s*=\s*(?:"([^"]+)"|([^;]+))', header, re.I)
    if match:
        return re.split(r"[/\\]", (match.group(1) or match.group(2)).strip())[-1]
    return None


def excel_extension(path: Path, content_type: str = "", disposition: str = "") -> str | None:
    if not path.is_file() or path.stat().st_size == 0:
        return None
    with path.open("rb") as stream:
        first = stream.read(512)

    if first.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        return ".xls"

    if first.startswith(b"PK") and zipfile.is_zipfile(path):
        try:
            with zipfile.ZipFile(path) as archive:
                names = set(archive.namelist())
                if "xl/workbook.bin" in names:
                    return ".xlsb"
                if "xl/workbook.xml" in names:
                    types = (
                        archive.read("[Content_Types].xml").decode("utf-8", "replace")
                        if "[Content_Types].xml" in names else ""
                    )
                    return ".xlsm" if "macroEnabled" in types else ".xlsx"
        except (OSError, zipfile.BadZipFile, KeyError):
            return None

    # Official site can also publish an HTML table with an .xls filename.
    original = (filename_from_disposition(disposition) or "").lower()
    if (original.endswith(".xls") and "excel" in content_type.lower()
            and (b"<html" in first.lower() or b"<table" in first.lower())):
        return ".xls"
    return None


def sheets_export_url(final_url: str, html: str) -> str | None:
    match = re.search(
        r"https?://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)",
        final_url + "\n" + html.replace("\\/", "/"),
    )
    if match:
        return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=xlsx"
    return None


def create_session(insecure: bool = False) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36",
        "Accept": "*/*",
        "Accept-Language": "th-TH,th;q=0.9,en;q=0.8",
        "Referer": BASE_URL,
    })
    retry = Retry(
        total=3, backoff_factor=1.3,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"], respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.verify = not insecure
    if insecure:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    return session


def parse_rows(rows, category: str, topic: str | None = None, topic_index: int = 0) -> list[Job]:
    _, _, t, category_folder = CATEGORIES[category]
    jobs = []
    for index, row in enumerate(rows, 1):
        name_element = row.select_one("div.col-lg-8")
        if name_element is None:
            continue
        label = name_element.get_text(" ", strip=True).lstrip("- ").strip()
        dropdown = row.select_one("select")
        button = row.select_one("button[value]")
        if dropdown is None and button is None:
            continue

        if topic is None:
            current_topic, subtopic = label, ""
            folder = Path(category_folder) / f"{index:02d}_{safe_name(label)}"
        else:
            current_topic, subtopic = topic, label
            folder = (
                Path(category_folder) / f"{topic_index:02d}_{safe_name(topic)}"
                / f"{index:02d}_{safe_name(label)}"
            )

        if dropdown is not None:
            for option in dropdown.find_all("option"):
                file_id = str(option.get("value", "")).strip()
                value_label = option.get_text(" ", strip=True)
                if file_id and value_label:
                    jobs.append(Job(
                        category, current_topic, subtopic, value_label, file_id, t,
                        folder / safe_name(value_label),
                    ))
        elif button is not None:
            file_id = str(button.get("value", "")).strip()
            if file_id:
                jobs.append(Job(
                    category, current_topic, subtopic, "Direct Download", file_id, t,
                    folder / "Direct_Download", "button",
                ))
    return jobs


def parse_catalog(html: str, category: str) -> list[Job]:
    soup = BeautifulSoup(html, "html.parser")
    accordion = soup.select("#accordion > .card")
    if not accordion:
        return parse_rows(soup.select("div.row.border"), category)

    jobs = []
    for index, card in enumerate(accordion, 1):
        heading = card.select_one(".card-header span.l")
        if heading is not None:
            jobs.extend(parse_rows(
                card.select("div.row.border"), category,
                heading.get_text(" ", strip=True), index,
            ))
    return jobs


def discover_category(session: requests.Session, category: str) -> list[Job]:
    _, page, _, _ = CATEGORIES[category]
    response = session.get(
        BASE_URL + page, params={"tmp": random.random() * 10000}, timeout=(20, 90),
    )
    response.raise_for_status()
    response.encoding = "utf-8"
    jobs = parse_catalog(response.text, category)
    if not jobs:
        raise RuntimeError(f"No download entries found: {category}")
    return jobs


def existing_excel(folder: Path, file_id: str) -> Path | None:
    if folder.exists():
        for file in folder.glob(f"DLT_{file_id}_*"):
            if file.is_file() and file.suffix.lower() in EXCEL_EXTS and excel_extension(file):
                return file
    return None


def stream_file(
    session: requests.Session, url: str, params: dict | None,
    destination: Path, stop: threading.Event,
) -> tuple[str, str, str]:
    with session.get(
        url, params=params, timeout=(20, 180), stream=True, allow_redirects=True,
    ) as response:
        response.raise_for_status()
        final_url = response.url
        ctype = response.headers.get("Content-Type", "")
        disposition = response.headers.get("Content-Disposition", "")
        with destination.open("wb") as out:
            for chunk in response.iter_content(chunk_size=128 * 1024):
                if stop.is_set():
                    raise DownloadStopped()
                if chunk:
                    out.write(chunk)
    return final_url, ctype, disposition


def download_one(
    session: requests.Session, job: Job, root: Path,
    stop: threading.Event, force: bool,
) -> dict:
    folder = root / job.folder
    folder.mkdir(parents=True, exist_ok=True)
    params = {"t": job.t, "tmp": random.random() * 10000, "data_file": job.data_file}
    prepared = requests.Request("GET", DOWNLOAD_URL, params=params).prepare()
    result = {
        "timestamp": now_iso(), "category": job.category, "topic": job.topic,
        "subtopic": job.subtopic, "dropdown": job.dropdown, "data_file": job.data_file,
        "source_url": prepared.url, "final_url": "", "file_path": "",
        "file_size": 0, "sha256": "", "status": "", "error": "",
    }
    old = existing_excel(folder, job.data_file)
    if old and not force:
        result.update(status="skipped", file_path=str(old), file_size=old.stat().st_size)
        return result

    temporary = folder / f"DLT_{job.data_file}.part"
    try:
        if stop.is_set():
            raise DownloadStopped()
        final_url, ctype, disposition = stream_file(
            session, DOWNLOAD_URL, params, temporary, stop,
        )
        result["final_url"] = final_url
        extension = excel_extension(temporary, ctype, disposition)

        if extension is None:
            with temporary.open("rb") as stream:
                preview = stream.read(1024 * 1024).decode("utf-8", "replace")
            export_url = sheets_export_url(final_url, preview)
            if export_url and export_url != final_url:
                final_url, ctype, disposition = stream_file(
                    session, export_url, None, temporary, stop,
                )
                result["final_url"] = final_url
                extension = excel_extension(temporary, ctype, disposition)

        if extension is None:
            debug_dir = root / "_debug"
            debug_dir.mkdir(parents=True, exist_ok=True)
            debug_file = debug_dir / f"{job.category.replace('.', '_')}_{job.data_file}.response"
            temporary.replace(debug_file)
            result.update(
                status="not_excel",
                error=f"Non-Excel response: {ctype}; {result['final_url']}",
            )
            return result

        original = filename_from_disposition(disposition) or "statistics"
        filename = f"DLT_{job.data_file}_{safe_name(Path(original).stem)}{extension}"
        destination = folder / filename
        temporary.replace(destination)
        if force and old and old != destination:
            old.unlink(missing_ok=True)
        result.update(
            status="downloaded", file_path=str(destination),
            file_size=destination.stat().st_size, sha256=sha256(destination),
        )
        return result

    except DownloadStopped:
        result["status"] = "stopped"
        return result
    except requests.exceptions.SSLError as exc:
        result.update(status="ssl_error", error=str(exc))
        return result
    except Exception as exc:
        result.update(status="error", error=f"{type(exc).__name__}: {exc}")
        return result
    finally:
        temporary.unlink(missing_ok=True)


class DLTCollectorApp(App):
    TITLE = "DLT Statistics Collector"
    SUB_TITLE = "Excel Downloader | Terminal UI"

    CSS = """
    Screen { background: #0c111b; color: #e5eaf3; }
    Header, Footer { background: #16243a; }
    #main { height: 1fr; padding: 1 2; }
    #hero {
        height: 3; background: #16243a; color: #8fbaff;
        content-align: center middle; text-style: bold;
        border: round #2a4770; margin-bottom: 1;
    }
    #controls { height: auto; border: round #2a4770; padding: 1; margin-bottom: 1; }
    #settings { height: 3; }
    #category { width: 1fr; margin-right: 1; }
    #output { width: 2fr; }
    .setting { width: 1fr; margin-right: 1; }
    #buttons { height: 3; margin-top: 1; }
    #buttons Button { margin-right: 1; min-width: 14; }
    #scan { background: #2563eb; }
    #download_all { background: #15803d; }
    #download_one { background: #0f766e; }
    #stop { background: #a32929; }
    #options { height: 3; margin-top: 1; }
    #options Checkbox { margin-right: 3; }
    #stats { height: 3; margin-bottom: 1; }
    .metric {
        width: 1fr; height: 3; background: #16243a;
        border: round #2a4770; content-align: center middle;
        margin-right: 1; text-style: bold;
    }
    #catalog_panel { height: 1fr; border: round #2a4770; margin-bottom: 1; }
    .panel_title {
        height: 1; background: #16243a; color: #8fbaff;
        text-style: bold; padding-left: 1;
    }
    #catalog { height: 1fr; background: #101927; }
    DataTable > .datatable--cursor { background: #284c76; }
    #progress_panel { height: 5; border: round #2a4770; padding: 0 1; margin-bottom: 1; }
    #status { height: 1; color: #a7c5f3; }
    #progress { height: 2; }
    #log_panel { height: 10; border: round #2a4770; }
    #log { height: 1fr; background: #0c111b; padding: 0 1; }
    Input, Select { background: #152337; border: tall #304d72; }
    """

    BINDINGS = [
        ("ctrl+r", "discover", "Discover"),
        ("ctrl+d", "download_all", "Download all"),
        ("ctrl+s", "stop", "Stop"),
        ("ctrl+q", "quit", "Quit"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.jobs: list[Job] = []
        self.visible_jobs: list[Job] = []
        self.selected_job: Job | None = None
        self.stop_event = threading.Event()
        self.busy = False
        self.statuses: dict[str, str] = {}
        self.row_keys: dict[str, str] = {}
        self.counts = {"downloaded": 0, "skipped": 0, "error": 0}

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with Vertical(id="main"):
            yield Static("DLT STATISTICS  |  REGISTRATION DATA COLLECTOR", id="hero")
            with Vertical(id="controls"):
                with Horizontal(id="settings"):
                    yield Select(
                        [
                            ("ALL CATEGORIES", "ALL"),
                            ("1.1 Registration & Tax", "1.1"),
                            ("1.2 Vehicle Registration", "1.2"),
                            ("1.3 New Registration", "1.3"),
                            ("1.4 Accumulated Registration", "1.4"),
                        ],
                        value="ALL", allow_blank=False, id="category",
                    )
                    yield Input(value="DLT_Statistics", placeholder="Output folder", id="output")
                with Horizontal():
                    yield Input(value="0", placeholder="Limit (0 = all)", id="limit", classes="setting")
                    yield Input(value="0.8", placeholder="Delay (seconds)", id="delay", classes="setting")
                with Horizontal(id="options"):
                    yield Checkbox("Force redownload", id="force")
                    yield Checkbox("Disable SSL verification (unsafe)", id="insecure")
                with Horizontal(id="buttons"):
                    yield Button("DISCOVER", id="scan")
                    yield Button("DOWNLOAD ALL", id="download_all")
                    yield Button("DOWNLOAD ROW", id="download_one")
                    yield Button("STOP", id="stop")
            with Horizontal(id="stats"):
                yield Static("CATALOG: 0", id="stat_catalog", classes="metric")
                yield Static("SUCCESS: 0", id="stat_success", classes="metric")
                yield Static("SKIPPED: 0", id="stat_skipped", classes="metric")
                yield Static("FAILED: 0", id="stat_failed", classes="metric")
            with Vertical(id="catalog_panel"):
                yield Static(" DOWNLOAD CATALOG  |  Choose a row for individual download", classes="panel_title")
                yield DataTable(id="catalog", cursor_type="row", zebra_stripes=True)
            with Vertical(id="progress_panel"):
                yield Static("READY  |  Click DISCOVER to begin", id="status")
                yield ProgressBar(total=100, show_eta=True, show_percentage=True, id="progress")
            with Vertical(id="log_panel"):
                yield Static(" ACTIVITY LOG", classes="panel_title")
                yield RichLog(id="log", highlight=True, markup=True, wrap=True, auto_scroll=True)
        yield Footer()

    def on_mount(self) -> None:
        table = self.query_one("#catalog", DataTable)
        for heading, key, width in [
            ("#", "number", 6),
            ("CATEGORY", "category", 10),
            ("TOPIC / SUBTOPIC", "topic", 58),
            ("DROPDOWN", "dropdown", 22),
            ("FILE ID", "data_file", 12),
            ("STATUS", "status", 14),
        ]:
            table.add_column(heading, key=key, width=width)
        self.query_one("#stop", Button).disabled = True
        self.write_log("Ready. Install truststore for Windows SSL support.", "info")

    def write_log(self, message: str, level: str = "info") -> None:
        colors = {"info": "cyan", "success": "green", "warning": "yellow", "error": "red"}
        color = colors.get(level, "white")
        stamp = datetime.now().strftime("%H:%M:%S")
        # Escape Rich markup in untrusted server responses.
        from rich.markup import escape
        self.query_one("#log", RichLog).write(
            f"[dim]{stamp}[/dim] [{color}]{escape(str(message))}[/{color}]"
        )

    def set_status(self, message: str) -> None:
        self.query_one("#status", Static).update(message)

    def set_busy(self, value: bool) -> None:
        self.busy = value
        for name in ("scan", "download_all", "download_one"):
            self.query_one(f"#{name}", Button).disabled = value
        self.query_one("#category", Select).disabled = value
        self.query_one("#stop", Button).disabled = not value

    def update_metrics(self) -> None:
        for widget, value in [
            ("stat_catalog", f"CATALOG: {len(self.visible_jobs)}"),
            ("stat_success", f"SUCCESS: {self.counts['downloaded']}"),
            ("stat_skipped", f"SKIPPED: {self.counts['skipped']}"),
            ("stat_failed", f"FAILED: {self.counts['error']}"),
        ]:
            self.query_one(f"#{widget}", Static).update(value)

    def refresh_table(self) -> None:
        selected = self.query_one("#category", Select).value
        self.visible_jobs = [
            job for job in self.jobs if selected == "ALL" or job.category == selected
        ]
        table = self.query_one("#catalog", DataTable)
        table.clear()
        self.row_keys.clear()
        self.selected_job = None
        for index, job in enumerate(self.visible_jobs):
            row_id = str(index)
            table.add_row(
                str(index + 1), job.category, job.title, job.dropdown,
                job.data_file, self.statuses.get(job.key, "READY"),
                key=row_id,
            )
            self.row_keys[job.key] = row_id
        self.update_metrics()

    def set_row_status(self, job: Job, status: str) -> None:
        self.statuses[job.key] = status.upper()
        row_id = self.row_keys.get(job.key)
        if row_id is not None:
            try:
                self.query_one("#catalog", DataTable).update_cell(
                    row_id, "status", status.upper(),
                )
            except Exception:
                pass

    def on_select_changed(self, event: Select.Changed) -> None:
        if event.select.id == "category" and self.is_mounted:
            self.refresh_table()

    def on_data_table_row_highlighted(self, event: DataTable.RowHighlighted) -> None:
        try:
            index = int(str(event.row_key.value))
            self.selected_job = self.visible_jobs[index]
        except (IndexError, ValueError, TypeError, AttributeError):
            self.selected_job = None

    def on_data_table_row_selected(self, event: DataTable.RowSelected) -> None:
        self.on_data_table_row_highlighted(event)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        actions = {
            "scan": self.action_discover,
            "download_all": self.action_download_all,
            "download_one": self.download_selected,
            "stop": self.action_stop,
        }
        action = actions.get(event.button.id)
        if action is not None:
            action()

    def action_discover(self) -> None:
        if self.busy:
            return
        self.stop_event.clear()
        self.set_busy(True)
        self.query_one("#progress", ProgressBar).update(total=4, progress=0)
        self.set_status("DISCOVERING DLT STATISTICS...")
        self.write_log("Scanning all 4 categories...")
        insecure = self.query_one("#insecure", Checkbox).value
        self.discover_worker(insecure)

    @work(thread=True, group="scan", exclusive=True)
    def discover_worker(self, insecure: bool) -> None:
        session = create_session(insecure)
        jobs: list[Job] = []
        errors = 0
        try:
            for number, category in enumerate(CATEGORIES, 1):
                if self.stop_event.is_set():
                    break
                self.call_from_thread(self.write_log, f"Scanning {category}...")
                try:
                    found = discover_category(session, category)
                    jobs.extend(found)
                    self.call_from_thread(
                        self.write_log, f"{category}: {len(found)} entries", "success",
                    )
                except Exception as exc:
                    errors += 1
                    self.call_from_thread(
                        self.write_log, f"{category}: {type(exc).__name__}: {exc}", "error",
                    )
                self.call_from_thread(
                    self.query_one("#progress", ProgressBar).update,
                    progress=number,
                )
            unique = list({job.key: job for job in jobs}.values())
            self.call_from_thread(self.finish_discovery, unique, errors)
        finally:
            session.close()

    def finish_discovery(self, jobs: list[Job], errors: int) -> None:
        self.jobs = jobs
        self.refresh_table()
        self.set_busy(False)
        state = "STOPPED" if self.stop_event.is_set() else "READY"
        self.set_status(f"CATALOG {state}: {len(jobs)} entries")
        self.write_log(
            f"Discovery finished: {len(jobs)} entries, {errors} category errors.",
            "warning" if errors else "success",
        )

    def action_download_all(self) -> None:
        if not self.busy:
            self.start_download(list(self.visible_jobs))

    def download_selected(self) -> None:
        if self.busy:
            return
        if self.selected_job is None:
            self.write_log("Highlight a catalog row first.", "warning")
            return
        self.start_download([self.selected_job])

    def start_download(self, jobs: list[Job]) -> None:
        if not jobs:
            self.write_log("No jobs found. Click DISCOVER first.", "warning")
            return
        try:
            limit = int(self.query_one("#limit", Input).value)
            delay = float(self.query_one("#delay", Input).value)
            if limit < 0 or delay < 0:
                raise ValueError("Limit and delay must not be negative")
        except ValueError:
            self.write_log("Invalid limit or delay value.", "error")
            return

        output_name = self.query_one("#output", Input).value.strip()
        if not output_name:
            self.write_log("Output folder cannot be empty.", "error")
            return
        if limit:
            jobs = jobs[:limit]
        output = Path(output_name).expanduser().resolve()
        force = self.query_one("#force", Checkbox).value
        insecure = self.query_one("#insecure", Checkbox).value

        self.counts = {"downloaded": 0, "skipped": 0, "error": 0}
        self.update_metrics()
        self.stop_event.clear()
        self.set_busy(True)
        self.query_one("#progress", ProgressBar).update(total=len(jobs), progress=0)
        self.write_log(f"Starting {len(jobs)} downloads into {output}")
        if insecure:
            self.write_log("SSL verification is disabled (unsafe).", "warning")
        self.download_worker(jobs, output, force, insecure, delay)

    @work(thread=True, group="download", exclusive=True)
    def download_worker(
        self, jobs: list[Job], output: Path, force: bool,
        insecure: bool, delay: float,
    ) -> None:
        session = create_session(insecure)
        processed = 0
        try:
            output.mkdir(parents=True, exist_ok=True)
            manifest = output / "manifest.csv"
            header_needed = not manifest.exists() or manifest.stat().st_size == 0
            with manifest.open("a", newline="", encoding="utf-8-sig") as stream:
                writer = csv.DictWriter(stream, fieldnames=MANIFEST_COLUMNS)
                if header_needed:
                    writer.writeheader()
                    stream.flush()
                for index, job in enumerate(jobs, 1):
                    if self.stop_event.is_set():
                        break
                    self.call_from_thread(
                        self.set_status,
                        f"DOWNLOADING {index}/{len(jobs)} | {job.category} | ID={job.data_file}",
                    )
                    self.call_from_thread(self.set_row_status, job, "DOWNLOADING")
                    result = download_one(session, job, output, self.stop_event, force)
                    writer.writerow(result)
                    stream.flush()
                    processed += 1
                    self.call_from_thread(
                        self.handle_result, job, result, processed, len(jobs),
                    )
                    if result["status"] in ("stopped", "ssl_error"):
                        break
                    if index < len(jobs) and self.stop_event.wait(delay):
                        break
        except Exception as exc:
            self.call_from_thread(self.write_log, f"Worker failed: {exc}", "error")
        finally:
            session.close()
            self.call_from_thread(self.finish_download, processed, len(jobs))

    def handle_result(self, job: Job, result: dict, done: int, total: int) -> None:
        status = result["status"]
        self.set_row_status(job, status)
        if status in ("downloaded", "skipped"):
            self.counts[status] += 1
            self.write_log(
                f"[{done}/{total}] {status.upper()} {job.category} "
                f"{job.dropdown} | {result['file_size']:,} bytes",
                "success" if status == "downloaded" else "warning",
            )
        elif status == "stopped":
            self.write_log("Download stopped.", "warning")
        else:
            self.counts["error"] += 1
            self.write_log(
                f"[{done}/{total}] {status.upper()} ID={job.data_file}: {result['error']}",
                "error",
            )
        self.query_one("#progress", ProgressBar).update(progress=done)
        self.update_metrics()

    def finish_download(self, done: int, total: int) -> None:
        self.set_busy(False)
        self.set_status(
            f"FINISHED {done}/{total} | SUCCESS {self.counts['downloaded']} | "
            f"SKIPPED {self.counts['skipped']} | FAILED {self.counts['error']}"
        )
        self.write_log(f"Run finished: {done}/{total} processed.", "success")

    def action_stop(self) -> None:
        if self.busy:
            self.stop_event.set()
            self.set_status("STOP REQUESTED")
            self.write_log("Stop requested.", "warning")


if __name__ == "__main__":
    DLTCollectorApp().run()

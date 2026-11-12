
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DLT STATISTICS COLLECTOR
Professional Terminal UI

INSTALL:
    python -m pip install -U textual rich requests beautifulsoup4 truststore

RUN:
    python dlt_collector.py

SHORTCUTS:
    Ctrl+R  Refresh catalog
    Ctrl+D  Download all visible files
    Ctrl+S  Stop
    Ctrl+Q  Quit
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
from urllib3.util.retry import Retry

from rich.markup import escape
from rich.text import Text

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
    Label,
    ProgressBar,
    RichLog,
    Select,
    Static,
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_URL = "https://web.dlt.go.th/statistics/"

DOWNLOAD_URL = BASE_URL + "load_file_select_new_car.php"

CATEGORIES = {
    "1.1": {
        "name": "Registration & Tax",
        "endpoint": "m1_stat_car.php",
        "type_id": 4,
        "directory": "01_Registration_Tax",
    },
    "1.2": {
        "name": "Vehicle Registration",
        "endpoint": "m1_new_car.php",
        "type_id": 1,
        "directory": "02_Vehicle_Registration",
    },
    "1.3": {
        "name": "New Vehicle Registration",
        "endpoint": "m1_new_car_red.php",
        "type_id": 2,
        "directory": "03_New_Registration",
    },
    "1.4": {
        "name": "Accumulated Registration",
        "endpoint": "m1_collect_new_car.php",
        "type_id": 3,
        "directory": "04_Accumulated_Registration",
    },
}

MANIFEST_COLUMNS = [
    "timestamp",
    "category",
    "topic",
    "subtopic",
    "period",
    "file_id",
    "request_url",
    "final_url",
    "file_path",
    "file_size",
    "sha256",
    "status",
    "error",
]

EXCEL_EXTENSIONS = {
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
}

INVALID_FILENAME = re.compile(
    r'[<>:"/\\|?*\x00-\x1f]'
)

DEFAULT_DELAY = 0.8


# ============================================================
# DATA MODEL
# ============================================================

@dataclass(frozen=True)
class DownloadJob:
    category: str
    topic: str
    subtopic: str
    period: str
    file_id: str
    type_id: int
    directory: Path

    @property
    def description(self) -> str:
        if self.subtopic:
            return f"{self.topic} / {self.subtopic}"
        return self.topic

    @property
    def key(self) -> str:
        return (
            f"{self.category}:"
            f"{self.type_id}:"
            f"{self.file_id}"
        )


class DownloadInterrupted(Exception):
    pass


# ============================================================
# UTILITIES
# ============================================================

def safe_filename(
    value: str,
    max_length: int = 54,
) -> str:

    value = re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip().lstrip("- ")

    value = INVALID_FILENAME.sub(
        "_",
        value,
    ).rstrip(" .")

    return value[:max_length].rstrip(" .") or "unknown"


def timestamp() -> str:
    return datetime.now().astimezone().isoformat(
        timespec="seconds"
    )


def sha256_file(path: Path) -> str:

    digest = hashlib.sha256()

    with path.open("rb") as file:
        for chunk in iter(
            lambda: file.read(1024 * 1024),
            b"",
        ):
            digest.update(chunk)

    return digest.hexdigest()


def extract_filename(
    disposition: str,
) -> str | None:

    match = re.search(
        r"filename\*\s*=\s*(?:UTF-8'')?([^;]+)",
        disposition,
        re.I,
    )

    if match:
        value = unquote(
            match.group(1).strip().strip('"')
        )
        return re.split(r"[/\\]", value)[-1]

    match = re.search(
        r'filename\s*=\s*(?:"([^"]+)"|([^;]+))',
        disposition,
        re.I,
    )

    if match:
        value = (
            match.group(1)
            or match.group(2)
        ).strip()

        return re.split(r"[/\\]", value)[-1]

    return None


def detect_excel(
    path: Path,
    content_type: str = "",
    disposition: str = "",
) -> str | None:

    if not path.is_file():
        return None

    if path.stat().st_size == 0:
        return None

    with path.open("rb") as file:
        signature = file.read(512)

    if signature.startswith(
        bytes.fromhex("D0CF11E0A1B11AE1")
    ):
        return ".xls"

    if signature.startswith(b"PK"):

        if not zipfile.is_zipfile(path):
            return None

        try:
            with zipfile.ZipFile(path) as archive:

                names = set(archive.namelist())

                if "xl/workbook.bin" in names:
                    return ".xlsb"

                if "xl/workbook.xml" in names:

                    types = ""

                    if "[Content_Types].xml" in names:
                        types = archive.read(
                            "[Content_Types].xml"
                        ).decode(
                            "utf-8",
                            errors="replace",
                        )

                    if "macroEnabled" in types:
                        return ".xlsm"

                    return ".xlsx"

        except (
            zipfile.BadZipFile,
            OSError,
            KeyError,
        ):
            return None

    original = (
        extract_filename(disposition)
        or ""
    ).lower()

    if (
        original.endswith(".xls")
        and "excel" in content_type.lower()
        and (
            b"<html" in signature.lower()
            or b"<table" in signature.lower()
        )
    ):
        return ".xls"

    return None


def google_sheets_export(
    url: str,
    page: str,
) -> str | None:

    match = re.search(
        r"https?://docs\.google\.com/"
        r"spreadsheets/d/([A-Za-z0-9_-]+)",
        url + "\n" + page.replace("\\/", "/"),
    )

    if not match:
        return None

    return (
        "https://docs.google.com/"
        f"spreadsheets/d/{match.group(1)}"
        "/export?format=xlsx"
    )


# ============================================================
# HTTP SESSION
# ============================================================

def create_session(
    insecure: bool = False,
) -> requests.Session:

    session = requests.Session()

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/157.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": (
            "th-TH,th;q=0.9,en;q=0.8"
        ),
        "Referer": BASE_URL,
    })

    retry = Retry(
        total=3,
        connect=3,
        read=3,
        status=3,
        backoff_factor=1.3,
        status_forcelist=[
            429,
            500,
            502,
            503,
            504,
        ],
        allowed_methods=["GET"],
        respect_retry_after_header=True,
    )

    adapter = HTTPAdapter(
        max_retries=retry
    )

    session.mount("https://", adapter)
    session.mount("http://", adapter)

    session.verify = not insecure

    if insecure:
        import urllib3

        urllib3.disable_warnings(
            urllib3.exceptions.InsecureRequestWarning
        )

    return session


# ============================================================
# CATALOG PARSER
# ============================================================

def parse_rows(
    rows,
    category: str,
    topic: str | None = None,
    topic_number: int = 0,
) -> list[DownloadJob]:

    config = CATEGORIES[category]
    jobs = []

    for row_number, row in enumerate(
        rows,
        start=1,
    ):

        title = row.select_one(
            "div.col-lg-8"
        )

        if title is None:
            continue

        label = title.get_text(
            " ",
            strip=True,
        ).lstrip("- ").strip()

        dropdown = row.select_one("select")
        button = row.select_one("button[value]")

        if dropdown is None and button is None:
            continue

        if topic is None:

            current_topic = label
            current_subtopic = ""

            directory = (
                Path(config["directory"])
                / f"{row_number:02d}_{safe_filename(label)}"
            )

        else:

            current_topic = topic
            current_subtopic = label

            directory = (
                Path(config["directory"])
                / f"{topic_number:02d}_{safe_filename(topic)}"
                / f"{row_number:02d}_{safe_filename(label)}"
            )

        if dropdown is not None:

            for option in dropdown.find_all("option"):

                file_id = str(
                    option.get("value", "")
                ).strip()

                period = option.get_text(
                    " ",
                    strip=True,
                )

                if not file_id or not period:
                    continue

                jobs.append(
                    DownloadJob(
                        category=category,
                        topic=current_topic,
                        subtopic=current_subtopic,
                        period=period,
                        file_id=file_id,
                        type_id=config["type_id"],
                        directory=(
                            directory / safe_filename(period)
                        ),
                    )
                )

        elif button is not None:

            file_id = str(
                button.get("value", "")
            ).strip()

            if not file_id:
                continue

            jobs.append(
                DownloadJob(
                    category=category,
                    topic=current_topic,
                    subtopic=current_subtopic,
                    period="Direct download",
                    file_id=file_id,
                    type_id=config["type_id"],
                    directory=directory / "Direct_Download",
                )
            )

    return jobs


def parse_catalog(
    html: str,
    category: str,
) -> list[DownloadJob]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    accordions = soup.select(
        "#accordion > .card"
    )

    if not accordions:
        return parse_rows(
            soup.select("div.row.border"),
            category,
        )

    jobs = []

    for number, card in enumerate(
        accordions,
        start=1,
    ):

        heading = card.select_one(
            ".card-header span.l"
        )

        if heading is None:
            continue

        jobs.extend(
            parse_rows(
                rows=card.select("div.row.border"),
                category=category,
                topic=heading.get_text(
                    " ",
                    strip=True,
                ),
                topic_number=number,
            )
        )

    return jobs


def fetch_category(
    session: requests.Session,
    category: str,
) -> list[DownloadJob]:

    endpoint = CATEGORIES[category]["endpoint"]

    response = session.get(
        BASE_URL + endpoint,
        params={
            "tmp": random.random() * 10000
        },
        timeout=(20, 90),
    )

    response.raise_for_status()
    response.encoding = "utf-8"

    jobs = parse_catalog(
        response.text,
        category,
    )

    if not jobs:
        raise RuntimeError(
            f"No downloadable entries: {category}"
        )

    return jobs


# ============================================================
# DOWNLOAD ENGINE
# ============================================================

def find_existing(
    directory: Path,
    file_id: str,
) -> Path | None:

    if not directory.exists():
        return None

    for path in directory.glob(
        f"DLT_{file_id}_*"
    ):

        if (
            path.is_file()
            and path.suffix.lower() in EXCEL_EXTENSIONS
            and detect_excel(path)
        ):
            return path

    return None


def stream_download(
    session: requests.Session,
    url: str,
    parameters: dict | None,
    temporary: Path,
    stop_event: threading.Event,
) -> tuple[str, str, str]:

    with session.get(
        url,
        params=parameters,
        stream=True,
        allow_redirects=True,
        timeout=(20, 180),
    ) as response:

        response.raise_for_status()

        final_url = response.url

        content_type = response.headers.get(
            "Content-Type",
            "",
        )

        disposition = response.headers.get(
            "Content-Disposition",
            "",
        )

        with temporary.open("wb") as file:

            for chunk in response.iter_content(
                chunk_size=128 * 1024
            ):

                if stop_event.is_set():
                    raise DownloadInterrupted()

                if chunk:
                    file.write(chunk)

    return final_url, content_type, disposition


def download_file(
    session: requests.Session,
    job: DownloadJob,
    root: Path,
    stop_event: threading.Event,
    force: bool,
) -> dict:

    directory = root / job.directory

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    params = {
        "t": job.type_id,
        "tmp": random.random() * 10000,
        "data_file": job.file_id,
    }

    request_url = requests.Request(
        "GET",
        DOWNLOAD_URL,
        params=params,
    ).prepare().url

    result = {
        "timestamp": timestamp(),
        "category": job.category,
        "topic": job.topic,
        "subtopic": job.subtopic,
        "period": job.period,
        "file_id": job.file_id,
        "request_url": request_url,
        "final_url": "",
        "file_path": "",
        "file_size": 0,
        "sha256": "",
        "status": "",
        "error": "",
    }

    existing = find_existing(
        directory,
        job.file_id,
    )

    if existing and not force:

        result.update(
            status="skipped",
            file_path=str(existing),
            file_size=existing.stat().st_size,
        )

        return result

    temporary = (
        directory / f"DLT_{job.file_id}.part"
    )

    try:

        if stop_event.is_set():
            raise DownloadInterrupted()

        (
            final_url,
            content_type,
            disposition,
        ) = stream_download(
            session,
            DOWNLOAD_URL,
            params,
            temporary,
            stop_event,
        )

        result["final_url"] = final_url

        extension = detect_excel(
            temporary,
            content_type,
            disposition,
        )

        if extension is None:

            with temporary.open("rb") as file:

                html = file.read(
                    1024 * 1024
                ).decode(
                    "utf-8",
                    errors="replace",
                )

            export_url = google_sheets_export(
                final_url,
                html,
            )

            if export_url and export_url != final_url:

                (
                    final_url,
                    content_type,
                    disposition,
                ) = stream_download(
                    session,
                    export_url,
                    None,
                    temporary,
                    stop_event,
                )

                result["final_url"] = final_url

                extension = detect_excel(
                    temporary,
                    content_type,
                    disposition,
                )

        if extension is None:

            debug_directory = root / "_debug"

            debug_directory.mkdir(
                parents=True,
                exist_ok=True,
            )

            debug_file = (
                debug_directory
                / (
                    f"{job.category.replace('.', '_')}_"
                    f"{job.file_id}.response"
                )
            )

            temporary.replace(debug_file)

            result.update(
                status="not_excel",
                error=(
                    f"Unexpected response: "
                    f"{content_type}; {final_url}"
                ),
            )

            return result

        original = (
            extract_filename(disposition)
            or "statistics"
        )

        filename = (
            f"DLT_{job.file_id}_"
            f"{safe_filename(Path(original).stem)}"
            f"{extension}"
        )

        destination = directory / filename

        temporary.replace(destination)

        if (
            force
            and existing
            and existing != destination
        ):
            existing.unlink(missing_ok=True)

        result.update(
            status="downloaded",
            file_path=str(destination),
            file_size=destination.stat().st_size,
            sha256=sha256_file(destination),
        )

        return result

    except DownloadInterrupted:

        result["status"] = "stopped"
        return result

    except requests.exceptions.SSLError as error:

        result.update(
            status="ssl_error",
            error=str(error),
        )

        return result

    except Exception as error:

        result.update(
            status="error",
            error=f"{type(error).__name__}: {error}",
        )

        return result

    finally:

        temporary.unlink(missing_ok=True)


# ============================================================
# TERMINAL APPLICATION
# ============================================================

class DLTApp(App):

    TITLE = "DLT Statistics"
    SUB_TITLE = "Data Collection Workspace"

    CSS = """
    Screen {
        background: #0d141b;
        color: #dce5e9;
    }

    Header {
        background: #17232e;
        color: #dce5e9;
    }

    Footer {
        background: #17232e;
        color: #9aabb6;
    }

    #workspace {
        height: 1fr;
        padding: 0 2 1 2;
    }

    /* BRAND */

    #brand {
        height: 3;
        margin-top: 1;
        background: #17232e;
        border-left: thick #ac9b79;
        padding: 0 2;
        content-align: left middle;
        color: #e9e7df;
        text-style: bold;
    }

    /* CONFIGURATION */

    #configuration {
        height: 5;
        margin-top: 1;
    }

    #category_group {
        width: 39;
        height: 5;
        margin-right: 2;
    }

    #output_group {
        width: 1fr;
        height: 5;
    }

    .input_label {
        height: 1;
        color: #93a5b1;
        text-style: bold;
        padding-left: 1;
    }

    #output {
        width: 100%;
        height: 3;
    }

    /* CATEGORY SELECT: SINGLE BORDER, VISIBLE LABEL */

    #category {
        width: 100%;
        height: 3;
        min-height: 3;
        max-height: 3;
        border: tall #304352;
        padding: 0;
        background: #16232e;
        color: #ffffff;
    }

    #category:focus {
        border: tall #ac9b79;
    }

    #category > SelectCurrent {
        width: 100%;
        height: 1;
        min-height: 1;
        max-height: 1;
        border: none !important;
        padding: 0 1;
        background: #16232e;
        color: #ffffff;
    }

    #category > SelectCurrent > #label {
        width: 1fr;
        height: 1;
        color: #ffffff;
        background: transparent;
        text-style: bold;
    }

    #category > SelectCurrent > .arrow {
        color: #dce5e9;
        background: transparent;
    }

    #category > SelectOverlay {
        background: #172631;
        color: #ffffff;
        border: tall #405666;
        max-height: 12;
    }

    /* ACTION BUTTONS */

    #actions {
        height: 3;
        margin-top: 1;
        align: left middle;
    }

    #actions Button {
        height: 3;
        min-height: 3;
        margin-right: 1;
        content-align: center middle;
        text-style: bold;
    }

    #refresh {
        width: 19;
        min-width: 19;
        background: #30485a;
    }

    #download_all {
        width: 22;
        min-width: 22;
        background: #23685f;
    }

    #download_selected {
        width: 23;
        min-width: 23;
        background: #285361;
    }

    #stop {
        width: 12;
        min-width: 12;
        background: #61383b;
    }

    #settings_button {
        width: 15;
        min-width: 15;
        background: #263846;
    }

    Button {
        background: #263846;
        color: #e6ebed;
    }

    Button:hover {
        background: #344e60;
    }

    Button:focus {
        background: #3c596c;
        color: #ffffff;
    }

    Button:disabled {
        opacity: 0.45;
    }

    /* ADVANCED OPTIONS */

    #advanced {
        display: none;
        height: 5;
        margin-top: 1;
        padding: 0 1;
        background: #15212b;
        border-left: solid #405262;
    }

    #advanced_title {
        height: 1;
        color: #ac9b79;
        text-style: bold;
    }

    #advanced_row {
        height: 3;
        align: left middle;
    }

    .small_label {
        width: auto;
        height: 3;
        padding: 0 1;
        content-align: left middle;
        color: #9fb0ba;
    }

    #limit {
        width: 10;
        margin-right: 2;
    }

    #delay {
        width: 10;
        margin-right: 2;
    }

    #force {
        width: 22;
        margin-right: 2;
    }

    #insecure {
        width: 37;
    }

    /* INPUTS */

    Input {
        background: #16232e;
        color: #e5ebed;
        border: tall #304352;
    }

    Input:focus {
        border: tall #ac9b79;
    }

    Checkbox {
        color: #aabac4;
    }

    /* SEARCH */

    #search_row {
        height: 3;
        margin-top: 1;
        align: left middle;
    }

    #search_label {
        width: 11;
        height: 3;
        content-align: left middle;
        color: #93a5b1;
        text-style: bold;
    }

    #search {
        width: 1fr;
        height: 3;
    }

    /* SUMMARY */

    #summary {
        height: 3;
        margin-top: 1;
        background: #17232e;
        border: solid #293c49;
        padding: 0 2;
        content-align: left middle;
        color: #d7e0e4;
    }

    /* CATALOG */

    #catalog_panel {
        height: 1fr;
        min-height: 10;
        margin-top: 1;
        background: #111c25;
        border: solid #293c49;
    }

    #catalog_header {
        height: 2;
        padding: 0 1;
        background: #17232e;
        align: left middle;
    }

    #catalog_title {
        width: 1fr;
        height: 2;
        content-align: left middle;
        color: #c8b892;
        text-style: bold;
    }

    #catalog_info {
        width: 1fr;
        height: 2;
        content-align: right middle;
        color: #91a5b2;
    }

    #table {
        height: 1fr;
        background: #111c25;
    }

    DataTable > .datatable--header {
        background: #1e2e3a;
        color: #d9e2e7;
        text-style: bold;
    }

    DataTable > .datatable--cursor {
        background: #345265;
        color: #ffffff;
    }

    DataTable > .datatable--odd-row {
        background: #14212b;
    }

    /* PROGRESS */

    #progress_panel {
        height: 4;
        margin-top: 1;
        padding: 0 1;
        border: solid #293c49;
        background: #101a22;
    }

    #status {
        height: 1;
        color: #c7d3da;
    }

    #progress {
        height: 2;
    }

    /* ACTIVITY */

    #activity_panel {
        height: 8;
        margin-top: 1;
        border: solid #293c49;
        background: #101a22;
    }

    #activity_title {
        height: 1;
        padding-left: 1;
        color: #c8b892;
        text-style: bold;
        background: #17232e;
    }

    #activity {
        height: 1fr;
        padding: 0 1;
        background: #101a22;
    }
    """

    BINDINGS = [
        ("ctrl+r", "discover", "Refresh"),
        ("ctrl+d", "download_all", "Download"),
        ("ctrl+s", "stop", "Stop"),
        ("ctrl+q", "quit", "Quit"),
    ]

    # ========================================================
    # STATE
    # ========================================================

    def __init__(self):

        super().__init__()

        self.all_jobs: list[DownloadJob] = []
        self.visible_jobs: list[DownloadJob] = []

        self.selected_job: DownloadJob | None = None

        self.job_status: dict[str, str] = {}
        self.row_keys: dict[str, str] = {}

        self.stop_event = threading.Event()

        self.busy = False

        self.counters = {
            "downloaded": 0,
            "skipped": 0,
            "failed": 0,
        }

    # ========================================================
    # LAYOUT
    # ========================================================

    def compose(self) -> ComposeResult:

        yield Header(show_clock=True)

        with Vertical(id="workspace"):

            yield Static(
                "DLT  /  STATISTICS COLLECTOR"
                "                                      "
                "VEHICLE REGISTRATION DATA",
                id="brand",
            )

            # CATEGORY / OUTPUT

            with Horizontal(id="configuration"):

                with Vertical(id="category_group"):

                    yield Label(
                        "DATA CATEGORY",
                        classes="input_label",
                    )

                    # compact=True removes SelectCurrent's
                    # built-in border. The outer #category
                    # element provides the only border.

                    yield Select(
                        options=[
                            (
                                "All categories",
                                "ALL",
                            ),
                            (
                                "1.1  Registration & Tax",
                                "1.1",
                            ),
                            (
                                "1.2  Vehicle Registration",
                                "1.2",
                            ),
                            (
                                "1.3  New Registration",
                                "1.3",
                            ),
                            (
                                "1.4  Accumulated Registration",
                                "1.4",
                            ),
                        ],
                        value="ALL",
                        allow_blank=False,
                        compact=True,
                        id="category",
                    )

                with Vertical(id="output_group"):

                    yield Label(
                        "OUTPUT DIRECTORY",
                        classes="input_label",
                    )

                    yield Input(
                        value="DLT_Statistics",
                        placeholder="Output directory",
                        id="output",
                    )

            # ACTIONS

            with Horizontal(id="actions"):

                yield Button(
                    "REFRESH CATALOG",
                    id="refresh",
                )

                yield Button(
                    "DOWNLOAD ALL",
                    id="download_all",
                )

                yield Button(
                    "DOWNLOAD SELECTED",
                    id="download_selected",
                )

                yield Button(
                    "STOP",
                    id="stop",
                )

                yield Button(
                    "OPTIONS",
                    id="settings_button",
                )

            # ADVANCED OPTIONS

            with Vertical(id="advanced"):

                yield Static(
                    "DOWNLOAD OPTIONS",
                    id="advanced_title",
                )

                with Horizontal(id="advanced_row"):

                    yield Static(
                        "LIMIT",
                        classes="small_label",
                    )

                    yield Input(
                        value="0",
                        id="limit",
                    )

                    yield Static(
                        "DELAY (S)",
                        classes="small_label",
                    )

                    yield Input(
                        value=str(DEFAULT_DELAY),
                        id="delay",
                    )

                    yield Checkbox(
                        "Force redownload",
                        id="force",
                    )

                    yield Checkbox(
                        "Allow invalid SSL (unsafe)",
                        id="insecure",
                    )

            # SEARCH

            with Horizontal(id="search_row"):

                yield Static(
                    "SEARCH",
                    id="search_label",
                )

                yield Input(
                    placeholder=(
                        "Filter by dataset, period or file ID..."
                    ),
                    id="search",
                )

            # SUMMARY

            yield Static(
                "AVAILABLE   0        "
                "DOWNLOADED   0        "
                "SKIPPED   0        "
                "FAILED   0",
                id="summary",
            )

            # TABLE

            with Vertical(id="catalog_panel"):

                with Horizontal(id="catalog_header"):

                    yield Static(
                        "FILE CATALOG",
                        id="catalog_title",
                    )

                    yield Static(
                        "Loading catalog...",
                        id="catalog_info",
                    )

                yield DataTable(
                    id="table",
                    cursor_type="row",
                    zebra_stripes=True,
                )

            # PROGRESS

            with Vertical(id="progress_panel"):

                yield Static(
                    "READY  /  Waiting for catalog",
                    id="status",
                )

                yield ProgressBar(
                    total=1,
                    show_eta=False,
                    show_percentage=True,
                    id="progress",
                )

            # LOG

            with Vertical(id="activity_panel"):

                yield Static(
                    "ACTIVITY LOG",
                    id="activity_title",
                )

                yield RichLog(
                    id="activity",
                    markup=True,
                    highlight=False,
                    wrap=True,
                    auto_scroll=True,
                )

        yield Footer()

    # ========================================================
    # INITIALIZATION
    # ========================================================

    def on_mount(self) -> None:

        table = self.query_one(
            "#table",
            DataTable,
        )

        columns = [
            ("#", "number", 6),
            ("SECTION", "category", 10),
            ("DATASET", "dataset", 55),
            ("PERIOD / OPTION", "period", 22),
            ("FILE ID", "file_id", 12),
            ("STATUS", "status", 15),
        ]

        for title, key, width in columns:

            table.add_column(
                title,
                key=key,
                width=width,
            )

        self.query_one(
            "#stop",
            Button,
        ).disabled = True

        self.write_log(
            "Application initialized.",
        )

        self.write_log(
            "Connecting to DLT catalog...",
        )

        self.call_later(
            self.action_discover
        )

    # ========================================================
    # UI HELPERS
    # ========================================================

    def write_log(
        self,
        message: str,
        level: str = "info",
    ) -> None:

        colors = {
            "info": "#a9bbc7",
            "success": "#76c9ad",
            "warning": "#dec18e",
            "error": "#ec9f9f",
        }

        color = colors.get(
            level,
            "#a9bbc7",
        )

        current_time = datetime.now().strftime(
            "%H:%M:%S"
        )

        self.query_one(
            "#activity",
            RichLog,
        ).write(
            f"[dim]{current_time}[/dim]  "
            f"[color={color}]"
            f"{escape(str(message))}"
            f"[/color]"
        )

    def set_status(
        self,
        message: str,
    ) -> None:

        self.query_one(
            "#status",
            Static,
        ).update(message)

    def set_busy(
        self,
        busy: bool,
    ) -> None:

        self.busy = busy

        for name in (
            "refresh",
            "download_all",
            "download_selected",
        ):

            self.query_one(
                f"#{name}",
                Button,
            ).disabled = busy

        self.query_one(
            "#category",
            Select,
        ).disabled = busy

        self.query_one(
            "#stop",
            Button,
        ).disabled = not busy

    def set_progress(
        self,
        current: int,
        total: int,
    ) -> None:

        self.query_one(
            "#progress",
            ProgressBar,
        ).update(
            total=max(total, 1),
            progress=current,
        )

    def update_summary(self) -> None:

        summary = (
            f"AVAILABLE   {len(self.visible_jobs):,}"
            "        "
            f"DOWNLOADED   {self.counters['downloaded']:,}"
            "        "
            f"SKIPPED   {self.counters['skipped']:,}"
            "        "
            f"FAILED   {self.counters['failed']:,}"
        )

        self.query_one(
            "#summary",
            Static,
        ).update(summary)

    @staticmethod
    def formatted_status(
        status: str,
    ) -> Text:

        status = status.upper()

        colors = {
            "READY": "#9aafbd",
            "RUNNING": "#dec18e",
            "DOWNLOADED": "#76c9ad",
            "SKIPPED": "#aab8c1",
            "ERROR": "#ec9f9f",
            "NOT_EXCEL": "#ec9f9f",
            "SSL_ERROR": "#ec9f9f",
            "STOPPED": "#dec18e",
        }

        return Text(
            status,
            style=colors.get(
                status,
                "#9aafbd",
            ),
        )

    # ========================================================
    # TABLE
    # ========================================================

    def refresh_table(self) -> None:

        selected_category = self.query_one(
            "#category",
            Select,
        ).value

        search = self.query_one(
            "#search",
            Input,
        ).value.casefold().strip()

        self.visible_jobs = [
            job
            for job in self.all_jobs
            if (
                selected_category == "ALL"
                or job.category == selected_category
            )
            and (
                not search
                or search in " ".join(
                    [
                        job.category,
                        job.description,
                        job.period,
                        job.file_id,
                    ]
                ).casefold()
            )
        ]

        table = self.query_one(
            "#table",
            DataTable,
        )

        table.clear()

        self.row_keys.clear()
        self.selected_job = None

        for index, job in enumerate(
            self.visible_jobs
        ):

            row_key = str(index)

            table.add_row(
                str(index + 1),
                job.category,
                job.description,
                job.period,
                job.file_id,
                self.formatted_status(
                    self.job_status.get(
                        job.key,
                        "READY",
                    )
                ),
                key=row_key,
            )

            self.row_keys[job.key] = row_key

        if self.visible_jobs:
            self.selected_job = self.visible_jobs[0]

        self.query_one(
            "#catalog_info",
            Static,
        ).update(
            f"{len(self.visible_jobs):,} visible"
            "  /  "
            f"{len(self.all_jobs):,} total"
        )

        self.update_summary()

    def set_job_status(
        self,
        job: DownloadJob,
        status: str,
    ) -> None:

        status = status.upper()

        self.job_status[job.key] = status

        row_key = self.row_keys.get(job.key)

        if row_key is None:
            return

        try:
            self.query_one(
                "#table",
                DataTable,
            ).update_cell(
                row_key,
                "status",
                self.formatted_status(status),
            )

        except Exception:
            pass

    # ========================================================
    # EVENTS
    # ========================================================

    def on_select_changed(
        self,
        event: Select.Changed,
    ) -> None:

        if (
            event.select.id == "category"
            and self.is_mounted
        ):
            self.refresh_table()

    def on_input_changed(
        self,
        event: Input.Changed,
    ) -> None:

        if (
            event.input.id == "search"
            and self.is_mounted
        ):
            self.refresh_table()

    def on_data_table_row_highlighted(
        self,
        event: DataTable.RowHighlighted,
    ) -> None:

        try:
            index = int(
                str(event.row_key.value)
            )
            self.selected_job = self.visible_jobs[index]

        except (
            IndexError,
            ValueError,
            AttributeError,
            TypeError,
        ):
            self.selected_job = None

    def on_data_table_row_selected(
        self,
        event: DataTable.RowSelected,
    ) -> None:

        try:
            index = int(
                str(event.row_key.value)
            )

            self.selected_job = self.visible_jobs[index]

            self.set_status(
                "SELECTED  /  "
                f"{self.selected_job.category}  /  "
                f"{self.selected_job.period}  /  "
                f"ID {self.selected_job.file_id}"
            )

        except (
            IndexError,
            ValueError,
            AttributeError,
            TypeError,
        ):
            pass

    def on_button_pressed(
        self,
        event: Button.Pressed,
    ) -> None:

        button_id = event.button.id

        if button_id == "refresh":
            self.action_discover()

        elif button_id == "download_all":
            self.action_download_all()

        elif button_id == "download_selected":
            self.action_download_selected()

        elif button_id == "stop":
            self.action_stop()

        elif button_id == "settings_button":
            self.toggle_settings()

    def toggle_settings(self) -> None:

        panel = self.query_one(
            "#advanced",
            Vertical,
        )

        panel.display = not panel.display

        button = self.query_one(
            "#settings_button",
            Button,
        )

        button.label = (
            "CLOSE OPTIONS"
            if panel.display
            else "OPTIONS"
        )

    # ========================================================
    # DISCOVERY
    # ========================================================

    def action_discover(self) -> None:

        if self.busy:
            return

        self.stop_event.clear()
        self.set_busy(True)

        self.set_progress(0, 4)

        self.set_status(
            "DISCOVERING  /  "
            "Reading available files from DLT..."
        )

        insecure = self.query_one(
            "#insecure",
            Checkbox,
        ).value

        self.discover_worker(insecure)

    @work(
        thread=True,
        group="discovery",
        exclusive=True,
    )
    def discover_worker(
        self,
        insecure: bool,
    ) -> None:

        session = create_session(insecure)

        collected = []
        errors = 0

        try:

            for index, category in enumerate(
                CATEGORIES,
                start=1,
            ):

                if self.stop_event.is_set():
                    break

                self.call_from_thread(
                    self.write_log,
                    f"Scanning category {category}...",
                )

                try:

                    jobs = fetch_category(
                        session,
                        category,
                    )

                    collected.extend(jobs)

                    self.call_from_thread(
                        self.write_log,
                        (
                            f"{category}  /  "
                            f"{len(jobs)} files discovered"
                        ),
                        "success",
                    )

                except Exception as error:

                    errors += 1

                    self.call_from_thread(
                        self.write_log,
                        (
                            f"{category}  /  "
                            f"{type(error).__name__}: {error}"
                        ),
                        "error",
                    )

                self.call_from_thread(
                    self.set_progress,
                    index,
                    4,
                )

            unique = list(
                {
                    job.key: job
                    for job in collected
                }.values()
            )

            self.call_from_thread(
                self.discovery_finished,
                unique,
                errors,
            )

        finally:
            session.close()

    def discovery_finished(
        self,
        jobs: list[DownloadJob],
        errors: int,
    ) -> None:

        self.all_jobs = jobs

        self.refresh_table()
        self.set_busy(False)

        if self.stop_event.is_set():

            self.set_status(
                "DISCOVERY STOPPED"
            )

            return

        self.set_status(
            f"READY  /  "
            f"{len(jobs):,} files indexed"
        )

        if errors:

            self.write_log(
                f"Catalog loaded with "
                f"{errors} category errors.",
                "warning",
            )

        else:

            self.write_log(
                f"Catalog ready. "
                f"{len(jobs):,} files available.",
                "success",
            )

    # ========================================================
    # DOWNLOAD ACTIONS
    # ========================================================

    def action_download_all(self) -> None:

        if self.busy:
            return

        self.start_download(
            list(self.visible_jobs)
        )

    def action_download_selected(self) -> None:

        if self.busy:
            return

        if not self.selected_job:

            self.write_log(
                "Select a file from the catalog first.",
                "warning",
            )
            return

        self.start_download(
            [self.selected_job]
        )

    def start_download(
        self,
        jobs: list[DownloadJob],
    ) -> None:

        if not jobs:

            self.write_log(
                "No files available for download.",
                "warning",
            )
            return

        try:

            limit = int(
                self.query_one(
                    "#limit",
                    Input,
                ).value
            )

            delay = float(
                self.query_one(
                    "#delay",
                    Input,
                ).value
            )

            if limit < 0 or delay < 0:
                raise ValueError()

        except ValueError:

            self.write_log(
                "Invalid limit or delay setting.",
                "error",
            )
            return

        output_value = self.query_one(
            "#output",
            Input,
        ).value.strip()

        if not output_value:

            self.write_log(
                "Output directory cannot be empty.",
                "error",
            )
            return

        if limit:
            jobs = jobs[:limit]

        output = Path(
            output_value
        ).expanduser().resolve()

        force = self.query_one(
            "#force",
            Checkbox,
        ).value

        insecure = self.query_one(
            "#insecure",
            Checkbox,
        ).value

        self.counters = {
            "downloaded": 0,
            "skipped": 0,
            "failed": 0,
        }

        self.update_summary()

        self.stop_event.clear()
        self.set_busy(True)

        self.set_progress(
            0,
            len(jobs),
        )

        self.set_status(
            f"DOWNLOADING  /  "
            f"0 of {len(jobs)} files"
        )

        self.write_log(
            f"Starting download of {len(jobs)} files."
        )

        self.download_worker(
            jobs,
            output,
            force,
            insecure,
            delay,
        )

    # ========================================================
    # DOWNLOAD WORKER
    # ========================================================

    @work(
        thread=True,
        group="downloads",
        exclusive=True,
    )
    def download_worker(
        self,
        jobs: list[DownloadJob],
        output: Path,
        force: bool,
        insecure: bool,
        delay: float,
    ) -> None:

        session = create_session(insecure)

        processed = 0

        try:

            output.mkdir(
                parents=True,
                exist_ok=True,
            )

            manifest = output / "manifest.csv"

            write_header = (
                not manifest.exists()
                or manifest.stat().st_size == 0
            )

            with manifest.open(
                "a",
                newline="",
                encoding="utf-8-sig",
            ) as file:

                writer = csv.DictWriter(
                    file,
                    fieldnames=MANIFEST_COLUMNS,
                )

                if write_header:
                    writer.writeheader()
                    file.flush()

                for index, job in enumerate(
                    jobs,
                    start=1,
                ):

                    if self.stop_event.is_set():
                        break

                    self.call_from_thread(
                        self.set_status,
                        (
                            f"DOWNLOADING  /  "
                            f"{index} of {len(jobs)}"
                            f"  /  File {job.file_id}"
                        ),
                    )

                    self.call_from_thread(
                        self.set_job_status,
                        job,
                        "RUNNING",
                    )

                    result = download_file(
                        session,
                        job,
                        output,
                        self.stop_event,
                        force,
                    )

                    writer.writerow(result)
                    file.flush()

                    processed += 1

                    self.call_from_thread(
                        self.handle_result,
                        job,
                        result,
                        processed,
                        len(jobs),
                    )

                    if result["status"] in (
                        "ssl_error",
                        "stopped",
                    ):
                        break

                    if (
                        index < len(jobs)
                        and self.stop_event.wait(delay)
                    ):
                        break

        except Exception as error:

            self.call_from_thread(
                self.write_log,
                (
                    f"Download session error: "
                    f"{type(error).__name__}: {error}"
                ),
                "error",
            )

        finally:

            session.close()

            self.call_from_thread(
                self.download_finished,
                processed,
                len(jobs),
            )

    # ========================================================
    # LIVE RESULTS
    # ========================================================

    def handle_result(
        self,
        job: DownloadJob,
        result: dict,
        processed: int,
        total: int,
    ) -> None:

        status = result["status"]

        self.set_job_status(
            job,
            status,
        )

        if status == "downloaded":

            self.counters["downloaded"] += 1

            self.write_log(
                (
                    f"{processed}/{total}  "
                    f"Saved file {job.file_id}  "
                    f"({result['file_size']:,} bytes)"
                ),
                "success",
            )

        elif status == "skipped":

            self.counters["skipped"] += 1

            self.write_log(
                (
                    f"{processed}/{total}  "
                    f"Skipped file {job.file_id}"
                ),
            )

        elif status == "stopped":

            self.write_log(
                "Download stopped.",
                "warning",
            )

        else:

            self.counters["failed"] += 1

            self.write_log(
                (
                    f"{processed}/{total}  "
                    f"Failed file {job.file_id}: "
                    f"{result['error']}"
                ),
                "error",
            )

        self.set_progress(
            processed,
            total,
        )

        self.update_summary()

    def download_finished(
        self,
        processed: int,
        total: int,
    ) -> None:

        self.set_busy(False)

        state = (
            "STOPPED"
            if self.stop_event.is_set()
            else "COMPLETE"
        )

        self.set_status(
            (
                f"{state}  /  "
                f"{processed} of {total} files  /  "
                f"{self.counters['downloaded']} saved  /  "
                f"{self.counters['skipped']} skipped  /  "
                f"{self.counters['failed']} failed"
            )
        )

        self.write_log(
            (
                f"Session finished: "
                f"{processed}/{total} files processed."
            ),
            "success",
        )

    # ========================================================
    # STOP
    # ========================================================

    def action_stop(self) -> None:

        if not self.busy:
            return

        self.stop_event.set()

        self.set_status(
            "STOP REQUESTED  /  "
            "Finishing current request..."
        )

        self.write_log(
            "Stop requested.",
            "warning",
        )


# ============================================================
# ENTRY POINT
# ============================================================

def main() -> None:
    DLTApp().run()


if __name__ == "__main__":
    main()

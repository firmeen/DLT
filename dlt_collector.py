
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DLT Statistics Collector
Professional Terminal UI

INSTALL:
    python -m pip install -r requirements.txt

RUN:
    python dlt_collector.py
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
    ProgressBar,
    RichLog,
    Select,
    Static,
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_URL = "https://web.dlt.go.th/statistics/"

DOWNLOAD_URL = (
    BASE_URL + "load_file_select_new_car.php"
)

CATEGORIES = {
    "1.1": (
        "Registration & Tax",
        "m1_stat_car.php",
        4,
        "01_Registration_Tax",
    ),
    "1.2": (
        "Vehicle Registration",
        "m1_new_car.php",
        1,
        "02_Vehicle_Registration",
    ),
    "1.3": (
        "New Vehicle Registration",
        "m1_new_car_red.php",
        2,
        "03_New_Registration",
    ),
    "1.4": (
        "Accumulated Registration",
        "m1_collect_new_car.php",
        3,
        "04_Accumulated_Registration",
    ),
}

MANIFEST_FIELDS = [
    "timestamp",
    "category",
    "topic",
    "subtopic",
    "dropdown",
    "data_file",
    "source_url",
    "final_url",
    "file_path",
    "file_size",
    "sha256",
    "status",
    "error",
]

EXCEL_SUFFIXES = {
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
}

INVALID_FILENAME = re.compile(
    r'[<>:"/\\|?*\x00-\x1f]'
)


# ============================================================
# DATA MODEL
# ============================================================

@dataclass(frozen=True)
class Job:
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


class InterruptedDownload(Exception):
    pass


# ============================================================
# UTILITIES
# ============================================================

def safe_part(
    value: str,
    length: int = 52,
) -> str:

    value = re.sub(
        r"\s+",
        " ",
        value,
    ).strip().lstrip("- ")

    value = INVALID_FILENAME.sub(
        "_",
        value,
    ).rstrip(" .")

    return (
        value[:length].rstrip(" .")
        or "unknown"
    )


def filename_header(
    value: str,
) -> str | None:

    match = re.search(
        r"filename\*\s*=\s*(?:UTF-8'')?([^;]+)",
        value,
        re.I,
    )

    if match:
        name = unquote(
            match.group(1).strip().strip('"')
        )
        return re.split(
            r"[/\\]",
            name,
        )[-1]

    match = re.search(
        r'filename\s*=\s*(?:"([^"]+)"|([^;]+))',
        value,
        re.I,
    )

    if match:
        name = (
            match.group(1)
            or match.group(2)
        ).strip()

        return re.split(
            r"[/\\]",
            name,
        )[-1]

    return None


def identify_excel(
    path: Path,
    mime: str = "",
    disposition: str = "",
) -> str | None:

    if (
        not path.is_file()
        or path.stat().st_size == 0
    ):
        return None

    with path.open("rb") as stream:
        signature = stream.read(512)

    if signature.startswith(
        bytes.fromhex("D0CF11E0A1B11AE1")
    ):
        return ".xls"

    if signature.startswith(b"PK"):
        if zipfile.is_zipfile(path):

            try:
                with zipfile.ZipFile(path) as archive:

                    names = set(
                        archive.namelist()
                    )

                    if "xl/workbook.bin" in names:
                        return ".xlsb"

                    if "xl/workbook.xml" in names:

                        types = ""

                        if "[Content_Types].xml" in names:
                            types = archive.read(
                                "[Content_Types].xml"
                            ).decode(
                                "utf-8",
                                "replace",
                            )

                        if "macroEnabled" in types:
                            return ".xlsm"

                        return ".xlsx"

            except (
                OSError,
                zipfile.BadZipFile,
            ):
                return None

    original = (
        filename_header(disposition)
        or ""
    ).lower()

    if (
        original.endswith(".xls")
        and "excel" in mime.lower()
    ):
        if (
            b"<html" in signature.lower()
            or b"<table" in signature.lower()
        ):
            return ".xls"

    return None


def hash_file(path: Path) -> str:

    digest = hashlib.sha256()

    with path.open("rb") as stream:

        for chunk in iter(
            lambda: stream.read(1024 * 1024),
            b"",
        ):
            digest.update(chunk)

    return digest.hexdigest()


def spreadsheet_export(
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
# HTTP CLIENT
# ============================================================

def new_session(
    insecure: bool,
) -> requests.Session:

    session = requests.Session()

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/157.0.0.0 "
            "Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": (
            "th-TH,th;q=0.9,en;q=0.7"
        ),
        "Referer": BASE_URL,
    })

    retry = Retry(
        total=3,
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

    session.mount(
        "https://",
        adapter,
    )

    session.mount(
        "http://",
        adapter,
    )

    session.verify = not insecure

    if insecure:
        import urllib3

        urllib3.disable_warnings(
            urllib3.exceptions.InsecureRequestWarning
        )

    return session


# ============================================================
# CATALOG EXTRACTION
# ============================================================

def catalog_rows(
    elements,
    category: str,
    topic: str | None = None,
    topic_index: int = 0,
) -> list[Job]:

    _, _, type_id, category_dir = (
        CATEGORIES[category]
    )

    jobs = []

    for number, element in enumerate(
        elements,
        1,
    ):

        heading = element.select_one(
            "div.col-lg-8"
        )

        if heading is None:
            continue

        label = heading.get_text(
            " ",
            strip=True,
        ).lstrip("- ").strip()

        dropdown = element.select_one(
            "select"
        )

        button = element.select_one(
            "button[value]"
        )

        if dropdown is None and button is None:
            continue

        if topic is None:

            parent = label
            subtopic = ""

            directory = (
                Path(category_dir)
                / f"{number:02d}_{safe_part(label)}"
            )

        else:

            parent = topic
            subtopic = label

            directory = (
                Path(category_dir)
                / f"{topic_index:02d}_{safe_part(topic)}"
                / f"{number:02d}_{safe_part(label)}"
            )

        if dropdown is not None:

            for option in dropdown.find_all(
                "option"
            ):

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
                    Job(
                        category=category,
                        topic=parent,
                        subtopic=subtopic,
                        period=period,
                        file_id=file_id,
                        type_id=type_id,
                        directory=(
                            directory
                            / safe_part(period)
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
                Job(
                    category=category,
                    topic=parent,
                    subtopic=subtopic,
                    period="Direct download",
                    file_id=file_id,
                    type_id=type_id,
                    directory=(
                        directory
                        / "Direct_Download"
                    ),
                )
            )

    return jobs


def parse_catalog(
    html: str,
    category: str,
) -> list[Job]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    accordions = soup.select(
        "#accordion > .card"
    )

    if not accordions:

        return catalog_rows(
            soup.select("div.row.border"),
            category,
        )

    jobs = []

    for index, card in enumerate(
        accordions,
        1,
    ):

        title = card.select_one(
            ".card-header span.l"
        )

        if title is None:
            continue

        jobs.extend(
            catalog_rows(
                elements=card.select(
                    "div.row.border"
                ),
                category=category,
                topic=title.get_text(
                    " ",
                    strip=True,
                ),
                topic_index=index,
            )
        )

    return jobs


def fetch_catalog(
    session: requests.Session,
    category: str,
) -> list[Job]:

    _, endpoint, _, _ = CATEGORIES[category]

    response = session.get(
        BASE_URL + endpoint,
        params={
            "tmp": random.random() * 10000
        },
        timeout=(20, 90),
    )

    response.raise_for_status()
    response.encoding = "utf-8"

    found = parse_catalog(
        response.text,
        category,
    )

    if not found:
        raise ValueError(
            "No downloadable entries "
            f"in category {category}"
        )

    return found


# ============================================================
# DOWNLOAD ENGINE
# ============================================================

def existing_excel(
    directory: Path,
    file_id: str,
) -> Path | None:

    if not directory.exists():
        return None

    for path in directory.glob(
        f"DLT_{file_id}_*"
    ):

        if (
            path.suffix.lower() in EXCEL_SUFFIXES
            and identify_excel(path)
        ):
            return path

    return None


def retrieve(
    session: requests.Session,
    url: str,
    parameters: dict | None,
    destination: Path,
    stop: threading.Event,
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

        mime = response.headers.get(
            "Content-Type",
            "",
        )

        disposition = response.headers.get(
            "Content-Disposition",
            "",
        )

        with destination.open("wb") as stream:

            for data in response.iter_content(
                128 * 1024
            ):

                if stop.is_set():
                    raise InterruptedDownload()

                if data:
                    stream.write(data)

    return final_url, mime, disposition


def download_job(
    session: requests.Session,
    job: Job,
    root: Path,
    stop: threading.Event,
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

    url = requests.Request(
        "GET",
        DOWNLOAD_URL,
        params=params,
    ).prepare().url

    result = {
        "timestamp": (
            datetime.now()
            .astimezone()
            .isoformat(timespec="seconds")
        ),
        "category": job.category,
        "topic": job.topic,
        "subtopic": job.subtopic,
        "dropdown": job.period,
        "data_file": job.file_id,
        "source_url": url,
        "final_url": "",
        "file_path": "",
        "file_size": 0,
        "sha256": "",
        "status": "",
        "error": "",
    }

    old_file = existing_excel(
        directory,
        job.file_id,
    )

    if old_file and not force:

        result.update(
            status="skipped",
            file_path=str(old_file),
            file_size=old_file.stat().st_size,
        )

        return result

    temporary = (
        directory
        / f"DLT_{job.file_id}.part"
    )

    try:

        if stop.is_set():
            raise InterruptedDownload()

        final, mime, disposition = retrieve(
            session,
            DOWNLOAD_URL,
            params,
            temporary,
            stop,
        )

        result["final_url"] = final

        suffix = identify_excel(
            temporary,
            mime,
            disposition,
        )

        if suffix is None:

            with temporary.open("rb") as stream:

                page = stream.read(
                    1024 * 1024
                ).decode(
                    "utf-8",
                    "replace",
                )

            export = spreadsheet_export(
                final,
                page,
            )

            if export and export != final:

                final, mime, disposition = retrieve(
                    session,
                    export,
                    None,
                    temporary,
                    stop,
                )

                result["final_url"] = final

                suffix = identify_excel(
                    temporary,
                    mime,
                    disposition,
                )

        if suffix is None:

            debug = root / "_debug"

            debug.mkdir(
                parents=True,
                exist_ok=True,
            )

            debug_file = (
                debug
                / (
                    f"{job.category.replace('.', '_')}_"
                    f"{job.file_id}.response"
                )
            )

            temporary.replace(debug_file)

            result.update(
                status="not_excel",
                error=(
                    f"Unexpected file format: "
                    f"{mime}; {final}"
                ),
            )

            return result

        original = (
            filename_header(disposition)
            or "statistics"
        )

        filename = (
            f"DLT_{job.file_id}_"
            f"{safe_part(Path(original).stem)}"
            f"{suffix}"
        )

        output = directory / filename

        temporary.replace(output)

        if (
            force
            and old_file
            and old_file != output
        ):
            old_file.unlink(missing_ok=True)

        result.update(
            status="downloaded",
            file_path=str(output),
            file_size=output.stat().st_size,
            sha256=hash_file(output),
        )

        return result

    except InterruptedDownload:

        result["status"] = "stopped"
        return result

    except requests.exceptions.SSLError as exc:

        result.update(
            status="ssl_error",
            error=str(exc),
        )

        return result

    except Exception as exc:

        result.update(
            status="error",
            error=(
                f"{type(exc).__name__}: {exc}"
            ),
        )

        return result

    finally:

        temporary.unlink(
            missing_ok=True
        )


# ============================================================
# TERMINAL APPLICATION
# ============================================================

class DLTApp(App):

    TITLE = "DLT / Statistics"
    SUB_TITLE = "Vehicle Registration Data"

    CSS = """
    Screen {
        background: #0b1016;
        color: #dbe3e8;
    }

    Header {
        background: #111b26;
        color: #dbe3e8;
    }

    Footer {
        background: #111b26;
        color: #9aacbc;
    }

    #shell {
        height: 1fr;
        padding: 1 2 0 2;
    }

    #brand {
        height: 3;
        background: #14212c;
        border-left: thick #b6a27a;
        padding: 0 2;
        color: #e8e4d8;
        text-style: bold;
        content-align: left middle;
    }

    #toolbar {
        height: 3;
        margin-top: 1;
    }

    #actions {
        height: 3;
        margin-top: 1;
    }

    #options {
        height: 3;
        margin-top: 1;
    }

    .field-label {
        height: 3;
        width: auto;
        padding: 0 1;
        content-align: left middle;
        color: #95a5b1;
        text-style: bold;
    }

    #category {
        width: 31;
        margin-right: 2;
    }

    #output {
        width: 1fr;
        margin-right: 1;
    }

    #sync {
        width: 19;
    }

    #download_all {
        width: 22;
        margin-right: 1;
    }

    #download_selected {
        width: 24;
        margin-right: 1;
    }

    #stop {
        width: 12;
        margin-right: 2;
    }

    #search {
        width: 1fr;
    }

    #limit {
        width: 11;
        margin-right: 2;
    }

    #delay {
        width: 11;
        margin-right: 2;
    }

    #force {
        width: 23;
    }

    #insecure {
        width: 42;
    }

    Button {
        background: #203343;
        color: #e2e9ec;
        border: none;
    }

    Button:hover {
        background: #2b4558;
    }

    Button:focus {
        background: #345268;
        color: #ffffff;
    }

    Button:disabled {
        opacity: 0.45;
    }

    #sync {
        background: #354b61;
    }

    #download_all {
        background: #236b65;
    }

    #download_selected {
        background: #284b59;
    }

    #stop {
        background: #5a3335;
    }

    Input, Select {
        background: #131e28;
        color: #e6ecee;
        border: tall #273949;
    }

    Input:focus, Select:focus {
        border: tall #b6a27a;
    }

    Checkbox {
        color: #aebdc7;
    }

    #summary {
        height: 3;
        margin-top: 1;
        margin-bottom: 1;
    }

    .metric {
        width: 1fr;
        height: 3;
        background: #14212c;
        border: round #263949;
        text-style: bold;
        content-align: center middle;
        margin-right: 1;
        color: #cbd8df;
    }

    #metric_failed {
        margin-right: 0;
    }

    #catalog_panel {
        height: 1fr;
        min-height: 9;
        border: round #293a49;
        background: #101922;
    }

    #catalog_heading {
        height: 1;
        padding: 0 1;
    }

    #catalog_title {
        width: 1fr;
        color: #b6a27a;
        text-style: bold;
    }

    #catalog_note {
        width: 1fr;
        color: #8ba1b1;
        text-align: right;
    }

    #table {
        height: 1fr;
        background: #101922;
    }

    DataTable > .datatable--header {
        background: #1a2935;
        color: #c3d0d9;
        text-style: bold;
    }

    DataTable > .datatable--cursor {
        background: #294553;
        color: #ffffff;
    }

    DataTable > .datatable--odd-row {
        background: #121e28;
    }

    #progress_panel {
        height: 4;
        border: round #293a49;
        padding: 0 1;
        margin-top: 1;
    }

    #operation_status {
        height: 1;
        color: #c0ced5;
    }

    #progress {
        height: 2;
    }

    #log_panel {
        height: 8;
        border: round #293a49;
        margin-top: 1;
        background: #101922;
    }

    #log_title {
        height: 1;
        padding: 0 1;
        color: #b6a27a;
        text-style: bold;
    }

    #log {
        height: 1fr;
        padding: 0 1;
        background: #101922;
    }
    """

    BINDINGS = [
        ("ctrl+r", "discover", "Discover"),
        ("ctrl+d", "download_all", "Download all"),
        ("ctrl+s", "stop", "Stop"),
        ("ctrl+q", "quit", "Quit"),
    ]

    # --------------------------------------------------------
    # STATE
    # --------------------------------------------------------

    def __init__(self) -> None:

        super().__init__()

        self.all_jobs: list[Job] = []
        self.visible_jobs: list[Job] = []

        self.selected_job: Job | None = None

        self.status_by_key: dict[str, str] = {}
        self.row_by_key: dict[str, str] = {}

        self.busy = False

        self.stop_signal = threading.Event()

        self.counters = {
            "downloaded": 0,
            "skipped": 0,
            "failed": 0,
        }

    # --------------------------------------------------------
    # LAYOUT
    # --------------------------------------------------------

    def compose(self) -> ComposeResult:

        yield Header(show_clock=True)

        with Vertical(id="shell"):

            yield Static(
                "DLT   /   STATISTICS COLLECTION"
                "       Vehicle Registration · Excel Archive",
                id="brand",
            )

            # Row 1: Source and destination
            with Horizontal(id="toolbar"):

                yield Static(
                    "CATEGORY",
                    classes="field-label",
                )

                yield Select(
                    [
                        ("All categories", "ALL"),
                    ] + [
                        (
                            f"{number}  {info[0]}",
                            number,
                        )
                        for number, info in CATEGORIES.items()
                    ],
                    value="ALL",
                    allow_blank=False,
                    id="category",
                )

                yield Static(
                    "OUTPUT",
                    classes="field-label",
                )

                yield Input(
                    value="DLT_Statistics",
                    id="output",
                )

                yield Button(
                    "SYNC CATALOG",
                    id="sync",
                )

            # Row 2: Primary actions
            with Horizontal(id="actions"):

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

                yield Input(
                    placeholder=(
                        "Search topic, period or file ID..."
                    ),
                    id="search",
                )

            # Row 3: Advanced controls
            with Horizontal(id="options"):

                yield Static(
                    "LIMIT",
                    classes="field-label",
                )

                yield Input(
                    value="0",
                    id="limit",
                )

                yield Static(
                    "DELAY (S)",
                    classes="field-label",
                )

                yield Input(
                    value="0.8",
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

            # Summary cards
            with Horizontal(id="summary"):

                yield Static(
                    "AVAILABLE   0",
                    classes="metric",
                    id="metric_available",
                )

                yield Static(
                    "DOWNLOADED   0",
                    classes="metric",
                    id="metric_downloaded",
                )

                yield Static(
                    "SKIPPED   0",
                    classes="metric",
                    id="metric_skipped",
                )

                yield Static(
                    "FAILED   0",
                    classes="metric",
                    id="metric_failed",
                )

            # Main table
            with Vertical(id="catalog_panel"):

                with Horizontal(id="catalog_heading"):

                    yield Static(
                        "FILE CATALOG",
                        id="catalog_title",
                    )

                    yield Static(
                        "Select a row to download individually",
                        id="catalog_note",
                    )

                yield DataTable(
                    id="table",
                    cursor_type="row",
                    zebra_stripes=True,
                )

            # Progress
            with Vertical(id="progress_panel"):

                yield Static(
                    "READY   /   Sync the catalog to begin",
                    id="operation_status",
                )

                yield ProgressBar(
                    total=1,
                    show_eta=False,
                    show_percentage=True,
                    id="progress",
                )

            # Activity log
            with Vertical(id="log_panel"):

                yield Static(
                    "ACTIVITY",
                    id="log_title",
                )

                yield RichLog(
                    id="log",
                    markup=True,
                    highlight=False,
                    wrap=True,
                    auto_scroll=True,
                )

        yield Footer()

    # --------------------------------------------------------
    # INITIALIZATION
    # --------------------------------------------------------

    def on_mount(self) -> None:

        table = self.query_one(
            "#table",
            DataTable,
        )

        columns = [
            ("NO", "no", 5),
            ("SECTION", "category", 9),
            ("DATASET", "topic", 57),
            ("PERIOD / OPTION", "period", 22),
            ("FILE ID", "file_id", 10),
            ("STATUS", "status", 13),
        ]

        for label, key, width in columns:

            table.add_column(
                label,
                key=key,
                width=width,
            )

        self.query_one(
            "#stop",
            Button,
        ).disabled = True

        self.note(
            "Collector initialized. "
            "Choose SYNC CATALOG to load the source index."
        )

    # --------------------------------------------------------
    # UI UTILITIES
    # --------------------------------------------------------

    def note(
        self,
        message: str,
        kind: str = "info",
    ) -> None:

        colors = {
            "info": "#aabcc8",
            "ok": "#71c2b0",
            "warn": "#d6ba83",
            "error": "#e7a0a0",
        }

        color = colors.get(
            kind,
            "#aabcc8",
        )

        hour = datetime.now().strftime(
            "%H:%M:%S"
        )

        self.query_one(
            "#log",
            RichLog,
        ).write(
            f"[dim]{hour}[/dim] "
            f"[color={color}]"
            f"{escape(str(message))}"
            f"[/color]"
        )

    def set_operation(
        self,
        message: str,
    ) -> None:

        self.query_one(
            "#operation_status",
            Static,
        ).update(message)

    def set_busy(
        self,
        value: bool,
    ) -> None:

        self.busy = value

        for button in (
            "sync",
            "download_all",
            "download_selected",
        ):

            self.query_one(
                f"#{button}",
                Button,
            ).disabled = value

        self.query_one(
            "#category",
            Select,
        ).disabled = value

        self.query_one(
            "#stop",
            Button,
        ).disabled = not value

    def update_summary(self) -> None:

        values = {
            "metric_available": (
                f"AVAILABLE   {len(self.visible_jobs)}"
            ),
            "metric_downloaded": (
                f"DOWNLOADED   "
                f"{self.counters['downloaded']}"
            ),
            "metric_skipped": (
                f"SKIPPED   "
                f"{self.counters['skipped']}"
            ),
            "metric_failed": (
                f"FAILED   "
                f"{self.counters['failed']}"
            ),
        }

        for selector, value in values.items():

            self.query_one(
                f"#{selector}",
                Static,
            ).update(value)

    @staticmethod
    def status_text(status: str) -> Text:

        colors = {
            "READY": "#90a5b4",
            "RUNNING": "#d9c389",
            "DOWNLOADED": "#71c2b0",
            "SKIPPED": "#9daeb8",
            "ERROR": "#e7a0a0",
            "SSL_ERROR": "#e7a0a0",
            "NOT_EXCEL": "#e7a0a0",
            "STOPPED": "#d9c389",
        }

        return Text(
            status,
            style=colors.get(
                status,
                "#90a5b4",
            ),
        )

    # --------------------------------------------------------
    # CATALOG TABLE
    # --------------------------------------------------------

    def show_catalog(self) -> None:

        selected = self.query_one(
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
                selected == "ALL"
                or job.category == selected
            )
            and (
                not search
                or search in " ".join(
                    (
                        job.category,
                        job.description,
                        job.period,
                        job.file_id,
                    )
                ).casefold()
            )
        ]

        table = self.query_one(
            "#table",
            DataTable,
        )

        table.clear()

        self.row_by_key.clear()
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
                self.status_text(
                    self.status_by_key.get(
                        job.key,
                        "READY",
                    )
                ),
                key=row_key,
            )

            self.row_by_key[job.key] = row_key

        if self.all_jobs:

            catalog_note = (
                f"{len(self.visible_jobs):,} displayed"
                "  /  "
                f"{len(self.all_jobs):,} indexed"
            )

        else:

            catalog_note = (
                "Select a row to download individually"
            )

        self.query_one(
            "#catalog_note",
            Static,
        ).update(catalog_note)

        self.update_summary()

    def set_job_state(
        self,
        job: Job,
        status: str,
    ) -> None:

        status = status.upper()

        self.status_by_key[job.key] = status

        key = self.row_by_key.get(
            job.key
        )

        if key is None:
            return

        try:

            self.query_one(
                "#table",
                DataTable,
            ).update_cell(
                key,
                "status",
                self.status_text(status),
            )

        except Exception:
            pass

    # --------------------------------------------------------
    # EVENTS
    # --------------------------------------------------------

    def on_select_changed(
        self,
        event: Select.Changed,
    ) -> None:

        if (
            event.select.id == "category"
            and self.is_mounted
        ):
            self.show_catalog()

    def on_input_changed(
        self,
        event: Input.Changed,
    ) -> None:

        if (
            event.input.id == "search"
            and self.is_mounted
        ):
            self.show_catalog()

    def on_data_table_row_highlighted(
        self,
        event: DataTable.RowHighlighted,
    ) -> None:

        try:

            index = int(
                str(event.row_key.value)
            )

            self.selected_job = (
                self.visible_jobs[index]
            )

        except (
            AttributeError,
            TypeError,
            ValueError,
            IndexError,
        ):

            self.selected_job = None

    def on_data_table_row_selected(
        self,
        event: DataTable.RowSelected,
    ) -> None:

        self.on_data_table_row_highlighted(
            event
        )

    def on_button_pressed(
        self,
        event: Button.Pressed,
    ) -> None:

        handlers = {
            "sync": self.action_discover,
            "download_all": self.action_download_all,
            "download_selected": (
                self.action_download_selected
            ),
            "stop": self.action_stop,
        }

        handler = handlers.get(
            event.button.id
        )

        if handler:
            handler()

    # --------------------------------------------------------
    # CATALOG DISCOVERY
    # --------------------------------------------------------

    def action_discover(self) -> None:

        if self.busy:
            return

        self.stop_signal.clear()
        self.set_busy(True)

        self.query_one(
            "#progress",
            ProgressBar,
        ).update(
            total=4,
            progress=0,
        )

        self.set_operation(
            "INDEXING   /   "
            "Retrieving available files from DLT..."
        )

        self.note(
            "Connecting to DLT statistics catalog..."
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

        session = new_session(insecure)

        jobs = []
        failures = 0

        try:

            for i, category in enumerate(
                CATEGORIES,
                1,
            ):

                if self.stop_signal.is_set():
                    break

                try:

                    entries = fetch_catalog(
                        session,
                        category,
                    )

                    jobs.extend(entries)

                    self.call_from_thread(
                        self.note,
                        (
                            f"{category}  /  "
                            f"{len(entries)} entries indexed"
                        ),
                        "ok",
                    )

                except Exception as exc:

                    failures += 1

                    self.call_from_thread(
                        self.note,
                        (
                            f"{category}  /  "
                            f"{type(exc).__name__}: {exc}"
                        ),
                        "error",
                    )

                self.call_from_thread(
                    self.query_one(
                        "#progress",
                        ProgressBar,
                    ).update,
                    progress=i,
                )

            unique = list(
                {
                    job.key: job
                    for job in jobs
                }.values()
            )

            self.call_from_thread(
                self.discovery_done,
                unique,
                failures,
            )

        finally:

            session.close()

    def discovery_done(
        self,
        jobs: list[Job],
        failures: int,
    ) -> None:

        self.all_jobs = jobs

        self.show_catalog()
        self.set_busy(False)

        self.set_operation(
            f"CATALOG   /   "
            f"{len(jobs):,} files indexed"
            f" · {failures} category errors"
        )

        if failures:

            self.note(
                "Catalog finished with "
                f"{failures} failed categories.",
                "warn",
            )

        else:

            self.note(
                "Catalog ready: "
                f"{len(jobs):,} files "
                "across 4 categories.",
                "ok",
            )

    # --------------------------------------------------------
    # DOWNLOAD ACTIONS
    # --------------------------------------------------------

    def action_download_all(self) -> None:

        if self.busy:
            return

        self.launch_download(
            list(self.visible_jobs)
        )

    def action_download_selected(self) -> None:

        if self.busy:
            return

        if self.selected_job is None:

            self.note(
                "Highlight a row in "
                "the file catalog first.",
                "warn",
            )

            return

        self.launch_download(
            [self.selected_job]
        )

    def launch_download(
        self,
        jobs: list[Job],
    ) -> None:

        if not jobs:

            self.note(
                "No files selected. "
                "Sync the catalog first.",
                "warn",
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

            self.note(
                "LIMIT must be >= 0 "
                "and DELAY must be >= 0.",
                "error",
            )

            return

        name = self.query_one(
            "#output",
            Input,
        ).value.strip()

        if not name:

            self.note(
                "Output path must not be empty.",
                "error",
            )

            return

        if limit:
            jobs = jobs[:limit]

        root = Path(
            name
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

        self.stop_signal.clear()
        self.set_busy(True)

        self.query_one(
            "#progress",
            ProgressBar,
        ).update(
            total=len(jobs),
            progress=0,
        )

        self.set_operation(
            f"DOWNLOAD   /   "
            f"0 of {len(jobs)} processed"
        )

        self.note(
            f"Starting {len(jobs)} files"
            f" · output: {root}"
        )

        if insecure:

            self.note(
                "SSL certificate verification "
                "is disabled.",
                "warn",
            )

        self.download_worker(
            jobs,
            root,
            force,
            insecure,
            delay,
        )

    # --------------------------------------------------------
    # DOWNLOAD WORKER
    # --------------------------------------------------------

    @work(
        thread=True,
        group="downloads",
        exclusive=True,
    )
    def download_worker(
        self,
        jobs: list[Job],
        root: Path,
        force: bool,
        insecure: bool,
        delay: float,
    ) -> None:

        session = new_session(insecure)

        completed = 0

        try:

            root.mkdir(
                parents=True,
                exist_ok=True,
            )

            manifest = (
                root / "manifest.csv"
            )

            needs_header = (
                not manifest.exists()
                or manifest.stat().st_size == 0
            )

            with manifest.open(
                "a",
                newline="",
                encoding="utf-8-sig",
            ) as output:

                writer = csv.DictWriter(
                    output,
                    fieldnames=MANIFEST_FIELDS,
                )

                if needs_header:

                    writer.writeheader()
                    output.flush()

                for number, job in enumerate(
                    jobs,
                    1,
                ):

                    if self.stop_signal.is_set():
                        break

                    self.call_from_thread(
                        self.set_operation,
                        (
                            f"DOWNLOAD   /   "
                            f"{number} of {len(jobs)}"
                            f" · Section {job.category}"
                            f" · File {job.file_id}"
                        ),
                    )

                    self.call_from_thread(
                        self.set_job_state,
                        job,
                        "RUNNING",
                    )

                    result = download_job(
                        session,
                        job,
                        root,
                        self.stop_signal,
                        force,
                    )

                    writer.writerow(result)
                    output.flush()

                    completed += 1

                    self.call_from_thread(
                        self.download_result,
                        job,
                        result,
                        completed,
                        len(jobs),
                    )

                    if result["status"] in (
                        "ssl_error",
                        "stopped",
                    ):
                        break

                    if (
                        number < len(jobs)
                        and self.stop_signal.wait(delay)
                    ):
                        break

        except Exception as exc:

            self.call_from_thread(
                self.note,
                f"Download session error: {exc}",
                "error",
            )

        finally:

            session.close()

            self.call_from_thread(
                self.download_done,
                completed,
                len(jobs),
            )

    # --------------------------------------------------------
    # LIVE DOWNLOAD RESULTS
    # --------------------------------------------------------

    def download_result(
        self,
        job: Job,
        result: dict,
        completed: int,
        total: int,
    ) -> None:

        status = result["status"]

        self.set_job_state(
            job,
            status,
        )

        if status == "downloaded":

            self.counters["downloaded"] += 1

            self.note(
                f"{completed}/{total}"
                f"  SAVED  {job.file_id}"
                f" · {result['file_size']:,} bytes",
                "ok",
            )

        elif status == "skipped":

            self.counters["skipped"] += 1

            self.note(
                f"{completed}/{total}"
                f"  SKIPPED  {job.file_id}"
                " · already downloaded"
            )

        elif status == "stopped":

            self.note(
                "Download stopped by user.",
                "warn",
            )

        else:

            self.counters["failed"] += 1

            self.note(
                f"{completed}/{total}"
                f"  {status.upper()}"
                f"  {job.file_id}"
                f" · {result['error']}",
                "error",
            )

        self.query_one(
            "#progress",
            ProgressBar,
        ).update(
            progress=completed
        )

        self.update_summary()

    def download_done(
        self,
        completed: int,
        total: int,
    ) -> None:

        self.set_busy(False)

        status = (
            "STOPPED"
            if self.stop_signal.is_set()
            else "COMPLETE"
        )

        self.set_operation(
            f"{status}   /   "
            f"{completed} of {total}"
            f" · {self.counters['downloaded']} saved"
            f" · {self.counters['skipped']} skipped"
            f" · {self.counters['failed']} failed"
        )

        self.note(
            "Session finished: "
            f"{completed}/{total} entries processed.",
            "ok",
        )

    # --------------------------------------------------------
    # STOP
    # --------------------------------------------------------

    def action_stop(self) -> None:

        if not self.busy:
            return

        self.stop_signal.set()

        self.set_operation(
            "STOP REQUESTED   /   "
            "Finishing current request..."
        )

        self.note(
            "Stop requested.",
            "warn",
        )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    DLTApp().run()

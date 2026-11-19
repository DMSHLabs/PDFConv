"""Конвертация Markdown и LaTeX в PDF и DOCX через локальные pandoc и MiKTeX.

Модуль — единственная точка правды для конвертации: его используют GUI
(pdfconv.pyw), режим командной строки и внешняя обёртка C:\\Tools\\pandoc\\pdfbuild.py.

DOCX делает pandoc: формула становится объектом Word (OMML) и остаётся редактируемой.
PDF собирается через LaTeX-конвейер, который лечит три типовые беды:

* в pdflatex-преамбулу с кодировкой T2A добавляется cmap, иначе кириллицу в готовом
  PDF нельзя ни найти, ни скопировать (извлекается как мусор или как пустые глифы);
* поднимается \\emergencystretch: длинная неразрывная математика и идентификаторы
  иначе вылезают за правое поле и обрезаются краем листа;
* таблица шире текстового блока (частый случай для markdown-таблиц) на втором проходе
  набирается мельче, чтобы последний столбец не пропал.
"""
from __future__ import annotations

import functools
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from urllib.parse import unquote
from dataclasses import dataclass, field
from pathlib import Path

# Данные проекта (bin, lib, res) лежат в корне, а сам код — в src, поэтому различаем
# каталог кода и корень: в собранном exe (PyInstaller) корень — папка с exe, при запуске
# через Python — родитель каталога со скриптами.
CODE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else CODE_DIR.parent
BIN_DIR = PROJECT_DIR / "bin"
LIB_DIR = PROJECT_DIR / "lib"

# Локальные зависимости (pymupdf, tkinterdnd2) лежат рядом с программой, а не в
# глобальном site-packages: подключаем их и для командной строки, и для окна. В собранном
# exe они уже внутри, и лишний путь в sys.path только мешает импортам.
if not getattr(sys, "frozen", False) and LIB_DIR.is_dir() and str(LIB_DIR) not in sys.path:
    sys.path.insert(0, str(LIB_DIR))

SUPPORTED_SUFFIXES = {".md", ".markdown", ".tex", ".latex", ".docx"}
FORMATS = ("pdf", "docx", "md")
DEFAULT_FORMATS = ("pdf", "docx")

PANDOC_URL = "https://github.com/jgm/pandoc/releases/latest"
MIKTEX_URL = "https://miktex.org/download"
DND_URL = "https://pypi.org/project/tkinterdnd2/"
SCRIPT_ENGINE_URL = "https://miktex.org/kb/fix-script-engine-not-found"

PANDOC_LOCATIONS = [
    BIN_DIR / "pandoc" / "pandoc.exe",
    BIN_DIR / "pandoc.exe",
    Path(r"C:\Tools\pandoc\pandoc.exe"),
]
MIKTEX_LOCATIONS = [
    Path(r"C:\Program Files\MiKTeX\miktex\bin\x64"),
    Path(r"C:\Program Files (x86)\MiKTeX\miktex\bin\x64"),
    Path(r"C:\Program Files\MiKTeX\miktex\bin"),
    Path(r"C:\Program Files (x86)\MiKTeX\miktex\bin"),
    # установка «только для меня» ставит MiKTeX в профиль пользователя
    Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "MiKTeX" / "miktex" / "bin" / "x64",
]
# MiKTeX 26.5+: latexmk — Perl-скрипт, а интерпретатор Perl в дистрибутив MiKTeX не входит.
# Свежий Perl вендорится в bin\perl (Strawberry portable: perl\bin + perl\lib целиком,
# обновляется заменой каталога); каталог ищется и в PATH, и в стандартном месте
# portable-установки C:\Tools\strawberry-perl — как запасные варианты.
PERL_BUNDLED_DIR = BIN_DIR / "perl" / "bin"
PERL_KNOWN_DIR = Path(r"C:\Tools\strawberry-perl\perl\bin")
PERL_MIN_VERSION = (5, 16)  # старее не тянет latexmk из MiKTeX 26.5 (падает на модулях)

# Пакеты, которые конвейер подключает сам и которые нужны для кириллицы, переносов
# в \\texttt, подгонки таблиц и сборки markdown-маршрута. MiKTeX умеет доустанавливать их
# при первой сборке — но только если автоустановка разрешена.
REQUIRED_PACKAGES = {
    "cmap": "корректный текстовый слой PDF (поиск и копирование кириллицы)",
    "hyphenat": "перенос длинных идентификаторов в \\texttt",
    "etoolbox": "подгонка слишком широких таблиц",
    "fontspec": "сборка через xelatex (markdown-маршрут)",
    "unicode-math": "математика в markdown-маршруте (шаблон pandoc)",
    "geometry": "поля страницы (шаблон pandoc)",
    "xeCJK": "китайский и японский текст в markdown",
}

STRETCH = "\\setlength{\\emergencystretch}{8em}"
TABLE_FIT = (
    "\\usepackage{etoolbox}\n"
    "\\AtBeginEnvironment{longtable}{\\footnotesize\\setlength{\\tabcolsep}{4pt}}\n"
    "\\AtBeginEnvironment{tabular}{\\setlength{\\tabcolsep}{4pt}}"
)
# Вторая ступень: узких полей не хватило — уменьшаем ещё и шрифт таблицы.
TABLE_FIT_SMALL = (
    "\\AtBeginEnvironment{tabular}{\\footnotesize\\setlength{\\tabcolsep}{3pt}}\n"
    "\\AtBeginEnvironment{longtable}{\\scriptsize\\setlength{\\tabcolsep}{3pt}}"
)

# Технические предупреждения LaTeX, которые на результат не влияют: показываем одной
# строкой-примечанием, а не списком.
BENIGN_WARNINGS = (
    "Unknown CJK family",
    "Token not allowed in a PDF string",
    "fontenc already loaded",
    "Unused \\captionsetup",
    "Unused \\captionsetup",
    "Rerun to get",
    "Marginpar on page",
    "float specifier changed",
)
PANDOC_ARGS = [
    "--pdf-engine=xelatex",
    "-V", "mainfont=Times New Roman",
    "-V", "monofont=Consolas",
    "-V", "lang=ru",
    "-V", "geometry:a4paper",
    "-V", "geometry:margin=2.5cm",
    "-V", "fontsize=12pt",
    "-V", "CJKmainfont=Microsoft YaHei",
]

MAX_ISSUES = 6  # сколько подробностей показывать на файл, чтобы окно не разрасталось

# GUI-процесс консоли не имеет, и каждый дочерний консольный процесс (perl, latexmk,
# pandoc, утилиты MiKTeX) иначе открывает своё видимое окно консоли и держит его, пока
# идёт сборка. Флаг достаётся и внукам (latexmk → perl → pdflatex): они наследуют его
# скрытую консоль. В консольном CLI-режиме флаг просто ничего не меняет.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


# --- окружение ---------------------------------------------------------------

@dataclass
class Dependency:
    """Одна проверка окружения: что нашли или чего не хватает и как это исправить."""

    key: str
    title: str
    ok: bool
    detail: str
    fix: str = ""
    url: str = ""
    required: bool = True


@dataclass
class Result:
    """Итог конвертации одного файла."""

    source: Path
    requested: tuple[str, ...] = DEFAULT_FORMATS
    pdf: Path | None = None
    docx: Path | None = None
    md: Path | None = None
    pages: int | None = None
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def missing_formats(self) -> list[str]:
        return [name for name in self.requested if getattr(self, name) is None]

    @property
    def status(self) -> str:
        if self.errors or self.missing_formats:
            return "error"
        return "warnings" if self.warnings else "ok"


def _tool_env(extra_inputs: list[Path] | None = None) -> dict:
    """PATH с каталогами perl, MiKTeX и pandoc: дочерние процессы их не всегда наследуют.
    Perl идёт первым: обёртка latexmk в MiKTeX берёт первый perl.exe из PATH, и устаревший
    Perl из системного PATH иначе затирал бы свежий из пользовательского.

    extra_inputs — каталоги исходника для поиска LaTeX: latexmk собирает копию во временном
    каталоге, поэтому относительные \\includegraphics/\\input ищутся через TEXINPUTS/BIBINPUTS,
    а не рядом с исходным файлом."""
    env = dict(os.environ)
    parts = [str(path) for path in (find_perl_dir(), find_miktex(), find_pandoc_dir()) if path]
    parts.append(env.get("PATH", ""))
    env["PATH"] = os.pathsep.join(parts)
    if extra_inputs:
        dirs = os.pathsep.join(str(Path(path).resolve()) for path in extra_inputs)
        for var in ("TEXINPUTS", "BIBINPUTS"):
            env[var] = dirs + os.pathsep + env.get(var, "")
    return env


def find_pandoc() -> Path | None:
    for candidate in PANDOC_LOCATIONS:
        if candidate.exists():
            return candidate
    found = shutil.which("pandoc", path=tool_env_safe())
    return Path(found) if found else None


def tool_env_safe() -> str:
    return os.environ.get("PATH", "")


def find_pandoc_dir() -> Path | None:
    pandoc = find_pandoc()
    return pandoc.parent if pandoc else None


def find_miktex() -> Path | None:
    for candidate in MIKTEX_LOCATIONS:
        if (candidate / "latexmk.exe").exists():
            return candidate
    for name in ("latexmk", "xelatex", "pdflatex"):
        found = shutil.which(name, path=tool_env_safe())
        if found:
            return Path(found).parent
    return None


def resolve_tool(name: str) -> str | None:
    """Абсолютный путь к инструменту: bare-имя зависит от PATH, который не наследуется."""
    for directory in (find_miktex(), find_pandoc_dir()):
        if directory and (directory / (name + ".exe")).exists():
            return str(directory / (name + ".exe"))
    for directory in (find_miktex(), find_pandoc_dir()):
        if directory and (directory / name).exists():
            return str(directory / name)
    return shutil.which(name, path=tool_env_safe())


@functools.lru_cache(maxsize=1)
def find_perl_dir() -> Path | None:
    """Каталог с достаточно свежим perl.exe или None.

    latexmk в MiKTeX 26.5 — Perl-скрипт, и его обёртка берёт первый perl.exe из PATH.
    Если в системном PATH остался устаревший Perl (на одной из машин — ActiveState 5.8.8),
    свежий из пользовательского PATH он затирает: пользовательский PATH в Windows
    приписывается после системного. Поэтому сначала берём вендорный Perl из bin\\perl
    (проект самодостаточен и не зависит от PATH машины), затем сканируем весь PATH и в
    конце — стандартное место portable-установки; найденный каталог подставляем в начало
    PATH дочерних процессов.
    """
    directories: list[Path] = [PERL_BUNDLED_DIR]
    for raw in os.environ.get("PATH", "").split(os.pathsep):
        directory = Path(raw)
        if directory not in directories:
            directories.append(directory)
    if PERL_KNOWN_DIR not in directories:
        directories.append(PERL_KNOWN_DIR)
    for directory in directories:
        exe = directory / "perl.exe"
        if exe.exists() and _perl_version(exe) >= PERL_MIN_VERSION:
            return directory
    return None


def _perl_version(exe: Path) -> tuple[int, ...]:
    """Версия интерпретатора как (5, 42, 3) или () — запустить или разобрать не вышло."""
    try:
        run = subprocess.run([str(exe), "-e", "print $^V"], capture_output=True,
                             text=True, errors="replace", timeout=30,
                             creationflags=NO_WINDOW)
    except (OSError, subprocess.TimeoutExpired):
        return ()
    match = re.search(r"v(\d+)\.(\d+)(?:\.(\d+))?", run.stdout)
    return tuple(int(part) for part in match.groups() if part) if match else ()


def check_environment() -> list[Dependency]:
    """Проверки для панели «Окружение»: pandoc, MiKTeX, пакеты, drag&drop, pymupdf."""
    find_perl_dir.cache_clear()  # после установки Perl кнопка «проверить снова» должна его найти
    checks: list[Dependency] = []

    pandoc = find_pandoc()
    pandoc_in_app = pandoc is not None and BIN_DIR in pandoc.parents
    checks.append(Dependency(
        key="pandoc",
        title="pandoc (файлы .md и формат DOCX)",
        ok=pandoc is not None,
        detail=("найден: %s" % pandoc) if pandoc else
               "не найден — markdown-файлы и формат DOCX недоступны, .tex собирается в PDF",
        fix="" if pandoc else
            "Скачайте архив pandoc-*-windows-x86_64.zip и положите pandoc.exe в %s "
            "(или установите pandoc и добавьте его в PATH)" % (BIN_DIR / "pandoc"),
        url="" if pandoc else PANDOC_URL,
        required=True,
    ))
    if pandoc is not None and not pandoc_in_app:
        checks[-1].detail += " (вне каталога программы)"

    miktex = find_miktex()
    latexmk = resolve_tool("latexmk")
    checks.append(Dependency(
        key="miktex",
        title="MiKTeX (latexmk, xelatex, pdflatex)",
        ok=miktex is not None and latexmk is not None,
        detail=("найден: %s" % miktex) if miktex else
               "не найден ни в %s, ни в PATH" % MIKTEX_LOCATIONS[0],
        fix="" if miktex else
            "Установите MiKTeX (полный вариант Basic installer) — он даёт latexmk, xelatex "
            "и сам доустанавливает отсутствующие пакеты",
        url="" if miktex else MIKTEX_URL,
    ))

    missing_packages = missing_latex_packages()
    checks.append(Dependency(
        key="packages",
        title="Пакеты LaTeX: %s" % ", ".join(sorted(REQUIRED_PACKAGES)),
        ok=not missing_packages,
        detail=("все на месте") if not missing_packages else
               "не найдены: %s" % ", ".join(sorted(missing_packages)),
        fix="" if not missing_packages else
            "MiKTeX доустановит их при первой сборке, если включена автоустановка. "
            + miktex_advice(),
        url="" if not missing_packages else "https://docs.miktex.org/manual/miktex-console.html",
    ))

    autoinstall = miktex_autoinstall_state()
    if autoinstall is None:
        checks.append(Dependency(
            key="autoinstall", title="Автоустановка пакетов MiKTeX", ok=True,
            detail="состояние настройки определить не удалось — не критично, но первая сборка "
                   "может потребовать установки пакетов вручную", required=False))
    else:
        checks.append(Dependency(
            key="autoinstall", title="Автоустановка пакетов MiKTeX",
            ok=autoinstall,
            detail="включена — недостающие пакеты MiKTeX поставит сам" if autoinstall else
                   "выключена: на свежей установке первая сборка упадёт на отсутствующем пакете",
            fix="" if autoinstall else miktex_advice(),
            url="" if autoinstall else "https://docs.miktex.org/manual/miktex-console.html",
        ))

    # Отдельная проверка: здоровая установка не должна писать ничего лишнего в рабочий
    # каталог. Каталог вида «l» с логами означает включённое логирование MiKTeX с
    # относительным путём — из-за этого же движок может не запускаться вовсе.
    junk = stray_entries(scratch_dir())
    checks.append(Dependency(
        key="miktex-junk", title="MiKTeX: посторонние файлы в рабочем каталоге",
        ok=not junk,
        detail="чисто — лишних файлов MiKTeX не создаёт" if not junk else
               "обнаружено: %s" % ", ".join("'%s'" % name for name in junk),
        fix="" if not junk else miktex_logging_advice(junk),
        url="" if not junk else "https://docs.miktex.org/manual/miktex-console.html",
    ))

    # Наличие файлов ещё не значит, что LaTeX собирает: пробная сборка это проверяет.
    smoke_ok, smoke_detail, smoke_log, smoke_fix = latex_smoke_test()
    if smoke_ok:
        smoke_fix, smoke_url = "", ""
    elif smoke_fix:
        # Причина известна, и лечение уже сформулировано самой пробной сборкой.
        smoke_url = SCRIPT_ENGINE_URL if script_engine_missing(smoke_detail) else ""
    else:
        smoke_fix = ("Проверьте установку MiKTeX (возможно, она неполная — тогда переустановите) и "
                     + miktex_advice() + (". Подробный журнал: %s" % smoke_log if smoke_log else ""))
        smoke_url = MIKTEX_URL
    checks.append(Dependency(
        key="smoke",
        title="Пробная сборка LaTeX",
        ok=smoke_ok,
        detail=smoke_detail,
        fix=smoke_fix,
        url=smoke_url,
    ))

    try:
        import tkinterdnd2  # noqa: F401
        dnd_ok, dnd_detail = True, "перетаскивание файлов в окно работает"
    except Exception as error:  # noqa: BLE001
        dnd_ok = False
        dnd_detail = "не установлена (%s) — перетаскивание недоступно, пользуйтесь кнопкой" % error
    checks.append(Dependency(
        key="dragdrop",
        title="Перетаскивание файлов (tkinterdnd2)",
        ok=dnd_ok,
        detail=dnd_detail,
        fix="" if dnd_ok else "py -3 -m pip install --target \"%s\" tkinterdnd2" % LIB_DIR,
        url="" if dnd_ok else DND_URL,
        required=False,
    ))

    pdf_backend = active_pdf_backend()
    if pdf_backend:
        pdf_ok = True
        pdf_detail = "проверка готового PDF включена (%s)" % pdf_backend
    else:
        pdf_ok = False
        pdf_detail = ("не установлена (нет pymupdf, pdfminer.six и pypdf) — "
                      "проверка PDF выполняться не будет")
    checks.append(Dependency(
        key="pdfcheck",
        title="Проверка PDF (%s)" % (pdf_backend or "pymupdf/pdfminer/pypdf"),
        ok=pdf_ok,
        detail=pdf_detail,
        fix="" if pdf_ok else "py -3 -m pip install --target \"%s\" pypdf" % LIB_DIR,
        url="" if pdf_ok else "https://pypi.org/project/pypdf/",
        required=False,
    ))
    return checks


def missing_latex_packages() -> set[str]:
    """Каких из REQUIRED_PACKAGES нет в текущей установке MiKTeX."""
    kpsewhich = resolve_tool("kpsewhich")
    if not kpsewhich:
        return set()
    missing = set()
    for name, _ in REQUIRED_PACKAGES.items():
        result = subprocess.run([kpsewhich, name + ".sty"], capture_output=True, text=True,
                                env=_tool_env(), errors="replace", cwd=str(scratch_dir()),
                                creationflags=NO_WINDOW)
        if not result.stdout.strip():
            missing.add(name)
    return missing


def miktex_autoinstall_state() -> bool | None:
    """Разрешена ли в MiKTeX автоустановка пакетов: True/False, None — не удалось узнать.

    Без неё первая сборка на свежем MiKTeX падает: нужных пакетов ещё нет, а спросить
    разрешение MiKTeX в неинтерактивном запуске не может.
    """
    initexmf = resolve_tool("initexmf")
    if not initexmf:
        return None
    result = subprocess.run([initexmf, "--show-config-value=[MPM]AutoInstall"],
                            capture_output=True, text=True, env=_tool_env(), errors="replace",
                            cwd=str(scratch_dir()), creationflags=NO_WINDOW)
    if result.returncode != 0:
        return None
    value = result.stdout.strip().lower()
    if not value:
        return None
    return value in ("1", "yes", "true", "t", "2")


SMOKE_PDFLATEX = r"""\documentclass{article}
\usepackage{cmap}
\usepackage[T2A]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage[russian]{babel}
\usepackage{hyphenat}
\setlength{\emergencystretch}{8em}
\begin{document}
Проверка сборки: $E = mc^2$, \texttt{std::vector}, «кавычки».
\end{document}
"""

# Повторяет окружение markdown-маршрута: pandoc подключает fontspec, unicode-math и xeCJK.
SMOKE_XELATEX = r"""\documentclass{article}
\usepackage{fontspec}
\setmainfont{Times New Roman}
\usepackage{unicode-math}
\usepackage{xeCJK}
\setCJKmainfont{Microsoft YaHei}
\begin{document}
Проверка сборки: $E = mc^2$, 混合 текст.
\end{document}
"""


PDF_BACKENDS = ("pymupdf", "pdfminer", "pypdf")


def active_pdf_backend() -> str:
    """Первый доступный бэкенд проверки PDF: pymupdf, pdfminer или pypdf.

    Пустая строка — проверять нечем. Выбор по убыванию точности: pymupdf быстрый
    и точный, pdfminer лучше читает CID/CMAP-шрифты, pypdf — последний резерв.
    """
    for name in PDF_BACKENDS:
        try:
            if name == "pymupdf":
                import pymupdf  # noqa: F401
            elif name == "pdfminer":
                import pdfminer.high_level  # noqa: F401
            else:
                import pypdf  # noqa: F401
        except Exception:  # noqa: BLE001 — битые бинарные зависимости тоже сюда
            continue
        return name
    return ""


def pdf_document_text(pdf_path: Path) -> tuple[str, int] | None:
    """Текст и число страниц PDF; None — проверять нечем (нет ни одного бэкенда)."""
    backend = active_pdf_backend()
    if not backend:
        return None
    if backend == "pymupdf":
        import pymupdf
        document = pymupdf.open(pdf_path)
        try:
            return "".join(page.get_text() for page in document), document.page_count
        finally:
            document.close()
    if backend == "pdfminer":
        from pdfminer.high_level import extract_text
        from pdfminer.pdfpage import PDFPage
        with open(pdf_path, "rb") as stream:
            pages = sum(1 for _ in PDFPage.get_pages(stream))
        return extract_text(str(pdf_path)), pages
    from pypdf import PdfReader
    reader = PdfReader(str(pdf_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages), len(reader.pages)


def pdf_text_readable(pdf_path: Path, needle: str) -> bool | None:
    """Ищется ли слово в текстовом слое PDF; None — проверить нечем (нет бэкенда)."""
    try:
        data = pdf_document_text(pdf_path)
    except Exception:  # noqa: BLE001
        return None
    return None if data is None else needle in data[0]


def pdf_text_hint() -> str:
    """PDF собирается, а текст из него не читается: шрифтам не хватает обратного соответствия."""
    return ("PDF собирается, но кириллица в нём не ищется и не копируется — значит, шрифтам "
            "не прописано обратное соответствие символов. Для pdflatex-маршрута это лечит "
            "пакет cmap, подключённый ДО fontenc (конвейер делает это в prepare()); если текст "
            "идёт через xelatex, дело в шрифте: проверьте, что Times New Roman установлен")


def latex_smoke_test() -> tuple[bool, str, str, str]:
    """Собрать пробные документы — так проверяется работоспособность, а не наличие файлов.

    Возвращает (получилось, пояснение, путь к журналу неудачной сборки, готовый совет);
    совет пустой, если внятную формулировку даёт вызывающий код.
    """
    latexmk = resolve_tool("latexmk")
    if not latexmk:
        return False, "latexmk не найден — сборка PDF невозможна", "", ""

    problems, log_path, advice = [], "", ""
    for title, engine, source in (("pdflatex", "-pdf", SMOKE_PDFLATEX),
                                  ("xelatex", "-xelatex", SMOKE_XELATEX)):
        work_dir = Path(tempfile.mkdtemp(prefix="pdfconv-smoke-"))
        (work_dir / "main.tex").write_text(source, encoding="utf-8")
        try:
            run = subprocess.run(
                [latexmk, engine, "-interaction=nonstopmode", "main.tex"],
                cwd=str(work_dir), env=_tool_env(), capture_output=True, text=True,
                errors="replace", timeout=600, creationflags=NO_WINDOW)
        except subprocess.TimeoutExpired:
            problems.append("%s: сборка не завершилась за 10 минут" % title)
            continue
        produced = work_dir / "main.pdf"
        if not produced.exists():
            log = work_dir / "main.log"
            tail = log_tail(log)
            if not tail:
                # журнала нет: единственная зацепка — вывод самого latexmk
                tail = output_tail(run) or "PDF и журнал не созданы, latexmk ничего не вывел"
            problems.append("%s: %s" % (title, tail))
            if script_engine_missing(tail) and not advice:
                advice = script_engine_hint()
            if log.exists():
                log_path = str(log)
            continue
        # Собранный PDF ещё не значит, что из него читается текст: при неверном порядке
        # пакетов шрифты остаются без обратного соответствия и кириллица не ищется.
        if pdf_text_readable(produced, "Проверка") is False:
            problems.append("%s: PDF собрался, но кириллица в нём не ищется" % title)
            if not advice:
                advice = pdf_text_hint()
        shutil.rmtree(work_dir, ignore_errors=True)

    if problems:
        detail = "; ".join(problems[:2])
        if "журнал" in detail or "не вывел" in detail:
            detail = "%s. %s" % (detail, broken_miktex_hint())
        return False, detail, log_path, advice
    return True, "pdflatex и xelatex собирают тестовые документы", "", ""


def copy_log_beside(output_pdf: Path, log_path: Path) -> str:
    """Оставить журнал LaTeX рядом с результатом, чтобы было что смотреть при разборе."""
    if not log_path.exists():
        return ""
    target = output_pdf.with_suffix(".latex.log")
    try:
        shutil.copyfile(log_path, target)
    except OSError:
        return ""
    return str(target)


def log_tail(log_path: Path, lines: int = 3) -> str:
    """Самая полезная часть журнала: строки ошибок LaTeX, иначе последние осмысленные."""
    if not log_path.exists():
        return ""
    text = log_path.read_text(encoding="utf-8", errors="replace")
    errors = [" ".join(match.group(1).split())
              for match in re.finditer(r"^!\s*(.+)$", text, re.M)
              if not match.group(1).startswith("==>")]
    if not errors:
        candidates = [line.strip() for line in text.splitlines()
                      if "not found" in line.lower() or "error" in line.lower()]
        errors = candidates[-lines:]
    return "; ".join(errors[:lines])


def broken_miktex_hint() -> str:
    """Движок не стартует и не пишет журнал — типичный признак испорченной установки."""
    return ("Похоже, установка MiKTeX повреждена или обновление не завершилось: движок не "
            "стартует и даже журнал не создаётся. Частая причина — обновление MiKTeX без прав "
            "администратора (в логах MiKTeX видно «Отказано в доступе» на файлы в "
            "C:\\Program Files\\MiKTeX). Откройте MiKTeX Console от имени администратора, "
            "завершите обновление (Updates → Update now) или переустановите MiKTeX")


def script_engine_missing(text: str) -> bool:
    """Признак того, что latexmk не смог запуститься из-за Perl: интерпретатора нет вовсе
    или он устарел — тогда perl отвечает «File::Path version 2.08 required--this is only
    version …» на скрипте latexmk.pl."""
    lowered = text.lower()
    return ("script engine" in lowered or "script-engine-not-found" in lowered
            or ("latexmk.pl" in lowered and "required--this is only version" in lowered))


def script_engine_hint() -> str:
    """MiKTeX не находит подходящий интерпретатор скриптов: сам latexmk — Perl-скрипт."""
    return ("MiKTeX не смог запустить latexmk: это Perl-скрипт, а подходящего интерпретатора "
            "Perl нет — он не входит в состав MiKTeX 26.5+, и очень старый Perl (например, "
            "ActiveState 5.8) не подходит. Поставьте Strawberry Perl (годится portable-сборка "
            "без прав администратора) и добавьте его каталог perl\\bin в PATH, "
            "чтобы команда perl стала доступна программе")


def miktex_advice() -> str:
    return ("Включите автоустановку пакетов: MiKTeX Console → Settings → "
            "«Install missing packages on-the-fly: Yes», либо в командной строке "
            "initexmf --set-config-value=[MPM]AutoInstall=1, либо поставьте нужные пакеты "
            "вручную: mpm --install=<имя>")


# Инструменты MiKTeX (kpsewhich, initexmf) запускаем в служебном каталоге: если в MiKTeX
# включено логирование в файл, относительный путь к нему не должен пачкать папку программы.
_SCRATCH_DIR: Path | None = None


def scratch_dir() -> Path:
    global _SCRATCH_DIR
    if _SCRATCH_DIR is None:
        _SCRATCH_DIR = Path(tempfile.mkdtemp(prefix="pdfconv-tools-"))
    return _SCRATCH_DIR


# Что нормально появляется в рабочем каталоге сборки; всё остальное — следы посторонних
# настроек (например, логи MiKTeX, которые он пишет в относительный каталог).
EXPECTED_BUILD_FILES = {
    # всё, что штатно создают pdflatex/xelatex/latexmk и пакеты (babel, hyperref и т.п.)
    "main.tex", "main.log", "main.pdf", "main.xdv", "main.dvi", "main.ps", "main.aux",
    "main.out", "main.fls", "main.synctex.gz", "main.fdb_latexmk", "main.toc", "main.lof",
    "main.lot", "main.idx", "main.ind", "main.ilg", "main.bbl", "main.blg", "main.bcf",
    "main.run.xml", "main.nav", "main.snm", "main.vrb", "main.gls", "main.glo", "main.brf",
    "main.upa", "main.upb", "main.xwm", "main.prev_graphics", "main.synctex(busy)",
}


def stray_entries(work_dir: Path) -> list[str]:
    try:
        return sorted(item.name for item in work_dir.iterdir()
                      if item.name not in EXPECTED_BUILD_FILES)
    except OSError:
        return []


def output_tail(result, lines: int = 2) -> str:
    """Последние осмысленные строки вывода инструмента — когда журнала LaTeX нет."""
    if result is None:
        return ""
    text = ((result.stdout or "") + "\n" + (result.stderr or "")).strip()
    if not text:
        return ""
    interesting = [line.strip() for line in text.splitlines() if line.strip()]
    return " | ".join(interesting[-lines:])


def miktex_logging_advice(stray: list[str]) -> str:
    """MiKTeX пишет свои логи в относительный каталог — это ломает сборку и мусорит."""
    return ("MiKTeX создаёт в рабочем каталоге %s — похоже, в MiKTeX Console включено "
            "логирование MiKTeX-сообщений в файл, а путь к нему задан относительным. "
            "Откройте MiKTeX Console → Settings и либо выключите логирование в файл, либо "
            "укажите абсолютный путь (например, C:\\MiKTeX-Logs)"
            % ", ".join("'%s'" % name for name in stray))


# --- конвейер ----------------------------------------------------------------

def insert_after_documentclass(text: str, snippet: str) -> str:
    """Лямбда держит замену буквальной: "\\s" — недопустимая escape-последовательность."""
    return re.sub(
        r"(\\documentclass(?:\[[^\]]*\])?\{[^}]+\})",
        lambda match: match.group(1) + "\n" + snippet,
        text,
        count=1,
    )


# Выключная формула из одного aligned — так pandoc оформляет многострочные формулы markdown;
# в написанных руками .tex тот же блок встречается и в TeX-выключке $$...$$.
ALIGNED_DISPLAYS = (
    re.compile(r"\\\[\s*\\begin\{aligned\}(\[[^\]]*\])?(.*?)\\end\{aligned\}\s*\\\]", re.S),
    re.compile(r"\$\$\s*\\begin\{aligned\}(\[[^\]]*\])?(.*?)\\end\{aligned\}\s*\$\$", re.S),
)


def fix_inner_tags(text: str) -> str:
    r"""Перенести \tag из aligned в align*, где amsmath его принимает.

    Внутри aligned (и любого бокса) amsmath отвечает «\tag not allowed here» и сборка не
    даёт PDF: номер разрешён только в самой выключной формуле. Многострочная формула с
    номерами приходит из markdown именно как \[ \begin{aligned} ... \tag{6} \\ ... \end{aligned} \]
    (реальный документ — 37 пронумерованных формул, две из них внутри aligned), в .tex-входе
    встречается та же конструкция в $$...$$. Строки в align* те же, а номера остаются на полях.
    Вариант с необязательным аргументом ([t], [b]) не трогаем: у align* такого аргумента нет.
    """
    def convert(match):
        if match.group(1) or "\\tag" not in match.group(2):
            return match.group(0)
        return "\\begin{align*}" + match.group(2) + "\\end{align*}"
    for pattern in ALIGNED_DISPLAYS:
        text = pattern.sub(convert, text)
    return text


def prepare(text: str, add_cmap: bool) -> str:
    if add_cmap:
        # cmap должен идти ДО fontenc: загруженный после него, он предупреждает
        # «fontenc already loaded - some fonts may be unprocessed» и оставляет T2A-шрифты
        # без таблицы обратного соответствия — кириллица в PDF тогда не ищется и не копируется.
        patched = re.sub(
            r"(\\usepackage(?:\[[^\]]*\])?\{fontenc\})",
            lambda match: "\\usepackage{cmap}\n" + match.group(1),
            text,
            count=1,
        )
        text = patched if patched != text else insert_after_documentclass(text, "\\usepackage{cmap}")
    if "emergencystretch" not in text:
        text = insert_after_documentclass(text, STRETCH)
    if "hyphenat" not in text and "\\texttt" in text:
        text = insert_after_documentclass(text, "\\usepackage[htt]{hyphenat}")
    return text


def shrink_tabulars(text: str) -> str:
    """Обернуть каждую таблицу в resizebox, который сжимает её только если она шире текста."""
    return re.sub(
        r"(\\begin\{tabular\}.*?\\end\{tabular\})",
        lambda match: ("\\resizebox{\\ifdim\\width>\\textwidth \\textwidth\\else\\width\\fi}{!}{%\n"
                       + match.group(1) + "\n}"),
        text,
        flags=re.S,
    )


DISPLAY_MATH = re.compile(r"(?<!\\)\\\[(.+?)(?<!\\)\\\]", re.S)
TAG_IN_DISPLAY = re.compile(r"\\tag\*?\{([^{}]*)\}")


def shrink_displays(text: str) -> str:
    """Сжать выключные формулы шире текста.

    Выключную математику LaTeX не переносит и не сжимает сам — широкая формула уходит
    за поле. Каждая формула измеряется и масштабируется до \\textwidth, но только если
    шире (условие \\ifdim выполняется прямо в TeX, узкие формулы не трогаются).

    Формула с номером (\\tag) в бокс не помещается: amsmath разрешает \\tag только в самой
    выключной формуле, а в боксе пишет «\\tag not allowed here» и сборка не даёт PDF (так
    не собирался документ с 37 пронумерованными формулами). Поэтому такая формула сначала
    измеряется в боксе с номером, переписанным обычным текстом, и в исходном виде (с \\tag)
    остаётся, если по ширине проходит; сжимается же только та, что шире текста.
    """
    def wrap(match):
        formula = match.group(1)
        tag = TAG_IN_DISPLAY.search(formula)
        if not tag:
            return ("\\begin{center}\n\\setbox0=\\hbox{$\\displaystyle " + formula + "$}%\n"
                    "\\ifdim\\wd0>\\textwidth \\resizebox{\\textwidth}{!}{\\usebox0}"
                    "\\else\\usebox0\\fi\n\\end{center}")
        if "$" in tag.group(1):
            return match.group(0)  # номер сам содержит формулу — в боксе её не переписать
        measured = TAG_IN_DISPLAY.sub(lambda hit: "\\qquad\\text{(%s)}" % hit.group(1), formula)
        return ("\\setbox0=\\hbox{$\\displaystyle " + measured + "$}%\n"
                "\\ifdim\\wd0>\\textwidth \\begin{center}"
                "\\resizebox{\\textwidth}{!}{\\usebox0}\\end{center}"
                "\\else\\[" + formula + "\\]\\fi")
    return DISPLAY_MATH.sub(wrap, text)


def ensure_resizebox_support(text: str) -> str:
    r"""Подключить graphicx, если дальше в тексте есть \resizebox.

    \resizebox даёт пакет graphicx, а шаблон pandoc его не подключает: без этого шаг
    сжатия широких формул и таблиц падал с «Undefined control sequence» (и за ним
    «Missing number, treated as zero»), то есть лечил одну беду другой. Пакет ставится
    только когда сжатие действительно понадобилось.
    """
    if "\\resizebox" in text and "graphicx" not in text:
        return insert_after_documentclass(text, "\\usepackage{graphicx}")
    return text


# Инлайн-формула в том виде, в каком её пишет pandoc (\(...\)) или человек ($...$).
INLINE_MATH = re.compile(r"\\\((.+?)\\\)|\$(?!\$)([^$\n]+?)\$(?!\$)", re.S)


def _breaks_after_commas(formula: str) -> str:
    r"""Разрешить разрыв после запятых верхнего уровня: \, \allowbreak или ,\allowbreak."""
    result: list[str] = []
    depth = 0
    index = 0
    while index < len(formula):
        char = formula[index]
        if char == "\\":  # \{ \} \, — команда, а не группа и не запятая
            result.append(formula[index:index + 2])
            index += 2
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth = max(0, depth - 1)
        elif char == "," and depth == 0:
            before = formula[index - 1] if index else ""
            after = formula[index + 1] if index + 1 < len(formula) else ""
            if not (before.isdigit() and after.isdigit()):  # 3,14 — это число, а не список
                result.append(",\\allowbreak")
                index += 1
                continue
        result.append(char)
        index += 1
    return "".join(result)


def soften_inline_math(text: str) -> str:
    r"""Разрешить переносы в инлайн-формулах — TeX сам математику не переносит.

    Строка с длинной цепочкой формул уходит за поле: после запятой в математике разрыв
    невозможен, а другого просвета в такой строке нет (реальный документ — 129 пт за поле).
    \allowbreak — нулевой штраф и ничего не печатает, но даёт TeX точку возможного разрыва.
    Запятые внутри групп и между цифрами не трогаем: там разрыв бесполезен или испортит вид.
    """
    def fix(match):
        formula = match.group(1) if match.group(1) is not None else match.group(2)
        delimiters = ("\\(", "\\)") if match.group(1) is not None else ("$", "$")
        return delimiters[0] + _breaks_after_commas(formula) + delimiters[1]
    return INLINE_MATH.sub(fix, text)


def severe_overfull(log_path: Path) -> list[float]:
    if not log_path.exists():
        return []
    log = log_path.read_text(encoding="utf-8", errors="replace")
    return [float(value) for value in re.findall(r"Overfull \\hbox \(([\d.]+)pt too wide\)", log)
            if float(value) > 20]


def parse_log(log_path: Path, result: Result) -> None:
    """Разбор лога LaTeX: ошибки, предупреждения и пояснения к ним."""
    if not log_path.exists():
        return
    log = log_path.read_text(encoding="utf-8", errors="replace")

    def core(message: str) -> str:
        # «сборка не дала PDF — первая ошибка: X» и сама строка X — это одно и то же
        return " ".join(message.split("первая ошибка: ")[-1].split()).lower()

    seen = {core(existing) for existing in result.errors}
    for match in re.finditer(r"^!\s*(.+)$", log, re.M):
        message = " ".join(match.group(1).split())
        # финальную строку «Fatal error» и уже показанные ошибки не дублируем
        if message.startswith("==>") or core(message) in seen:
            continue
        seen.add(core(message))
        result.errors.append(message)

    overfull = [float(value) for value in
                re.findall(r"Overfull \\hbox \(([\d.]+)pt too wide\)", log)]
    big = [value for value in overfull if value > 20]
    if big:
        result.warnings.append(
            "%d строк выходят за правое поле (до %.0f пт) — обычно это длинная формула "
            "или неразрывный идентификатор без переноса" % (len(big), max(big)))
    elif overfull:
        result.notes.append("%d строк немного заходят за поле (мелкие, до 20 пт)" % len(overfull))

    missing_chars = sorted(set(re.findall(r"Missing character: There is no (\S+)", log)))
    if missing_chars:
        shown = ", ".join(missing_chars[:5])
        result.warnings.append(
            "в шрифте нет глифов: %s%s — поставьте шрифт с этими символами или замените текст"
            % (shown, " и др." if len(missing_chars) > 5 else ""))

    latex_warnings, benign = [], 0
    for match in re.finditer(r"^(?:LaTeX|Package \S+|Class \S+) Warning: (.+)$", log, re.M):
        message = " ".join(match.group(1).split())
        if "Label(s) may have changed" in message or not re.search(r"\w", message):
            continue  # уходит после повторного прохода latexmk / строка-разделитель
        if any(pattern in message for pattern in BENIGN_WARNINGS):
            benign += 1
            continue
        latex_warnings.append(message)
    for message in latex_warnings[:3]:
        result.warnings.append(message)
    if len(latex_warnings) > 3:
        result.notes.append("ещё %d предупреждений LaTeX в журнале сборки"
                            % (len(latex_warnings) - 3))
    if benign:
        result.notes.append("%d технических предупреждений LaTeX (на результат не влияют)"
                            % benign)


def verify_pdf(pdf_path: Path, source_text: str, result: Result) -> None:
    """Проверить готовый PDF: страницы и живой текстовый слой."""
    try:
        data = pdf_document_text(pdf_path)
    except Exception as error:  # noqa: BLE001
        result.errors.append("не удалось открыть готовый PDF: %s" % error)
        return
    if data is None:
        return
    text, pages = data
    result.pages = pages
    if text.count("\uffff") > 20:
        result.warnings.append("текстовый слой PDF повреждён (глифы без Unicode): "
                               "текст не ищется и не копируется")
        return
    if not text.strip():
        result.warnings.append("в PDF нет текстового слоя: текст не ищется и не копируется")
        return
    has_cyrillic_source = bool(re.search(r"[А-Яа-яЁё]", source_text))
    if has_cyrillic_source and not re.search(r"[А-Яа-яЁё]", text):
        result.warnings.append("в PDF не найден кириллический текст: проверьте копирование")


def build_latex(text: str, output_pdf: Path, work_dir: Path, result: Result,
                input_dir: Path | None = None) -> None:
    """Собрать PDF из готового LaTeX-исходника: подготовка, сборка, повтор с подгонкой.

    input_dir — каталог исходного .tex: через TEXINPUTS latexmk находит относительные
    картинки и \\input, хотя сам собирает во временном каталоге."""
    latexmk = resolve_tool("latexmk")
    if not latexmk:
        result.errors.append("не найден latexmk — установите MiKTeX")
        return

    needs_cmap = "fontspec" not in text
    prepared = fix_inner_tags(prepare(text, needs_cmap))
    (work_dir / "main.tex").write_text(prepared, encoding="utf-8")
    engine = "-pdf" if needs_cmap else "-xelatex"

    last_run = None

    def run():
        nonlocal last_run
        last_run = subprocess.run(
            [latexmk, engine, "-interaction=nonstopmode", "main.tex"],
            cwd=str(work_dir), env=_tool_env([input_dir] if input_dir else None),
            capture_output=True, text=True, errors="replace", creationflags=NO_WINDOW)
        return last_run

    run()
    log_path = work_dir / "main.log"
    # Ступени подгонки таблиц: следующая применяется, только если предыдущая не помогла.
    if severe_overfull(log_path) and ("longtable" in text or "tabular" in text):
        current = (work_dir / "main.tex").read_text(encoding="utf-8")
        (work_dir / "main.tex").write_text(
            insert_after_documentclass(current, TABLE_FIT), encoding="utf-8")
        run()
        result.notes.append("таблица была шире текста — уменьшены поля таблицы")

    if severe_overfull(log_path) and "tabular" in text:
        current = (work_dir / "main.tex").read_text(encoding="utf-8")
        (work_dir / "main.tex").write_text(
            insert_after_documentclass(current, TABLE_FIT_SMALL), encoding="utf-8")
        run()
        result.notes.append("таблица всё ещё не помещалась — уменьшен шрифт таблицы")

    if severe_overfull(log_path) and "tabular" in text:
        # Последняя ступень: сжимаем сами таблицы до ширины текста (только если шире).
        current = (work_dir / "main.tex").read_text(encoding="utf-8")
        (work_dir / "main.tex").write_text(
            ensure_resizebox_support(shrink_tabulars(current)), encoding="utf-8")
        run()
        result.notes.append("таблицы сжаты до ширины текстового блока")

    if severe_overfull(log_path):
        current = (work_dir / "main.tex").read_text(encoding="utf-8")
        if DISPLAY_MATH.search(current):
            # Выключная формула шире текста не переносится и не сжимается сама:
            # измеряем и, если шире, масштабируем её до ширины текста (см. shrink_displays)
            before = severe_overfull(log_path)
            (work_dir / "main.tex").write_text(
                ensure_resizebox_support(shrink_displays(current)), encoding="utf-8")
            run()
            if severe_overfull(log_path) != before:
                result.notes.append("широкая выключная формула сжата до ширины текста")

    if severe_overfull(log_path):
        # Строка с длинной цепочкой формул: сжимать её нечем, зато можно разрешить
        # переносы внутри самих формул (см. soften_inline_math).
        current = (work_dir / "main.tex").read_text(encoding="utf-8")
        if INLINE_MATH.search(current):
            before = severe_overfull(log_path)
            (work_dir / "main.tex").write_text(soften_inline_math(current), encoding="utf-8")
            run()
            if severe_overfull(log_path) != before:
                result.notes.append("в формулах внутри строки разрешены переносы после запятых")

    produced = work_dir / "main.pdf"
    if not produced.exists():
        # Без самой ошибки в отчёте разбираться не с чем, поэтому показываем её и оставляем
        # журнал рядом с исходником: временный каталог всё равно удаляется.
        reason = log_tail(log_path) or output_tail(last_run) or             "журнал LaTeX пуст (движок не запускался?)"
        first = reason.split(";")[0].strip()
        result.errors.append("сборка не дала PDF — первая ошибка: %s" % first)
        if script_engine_missing(reason):
            result.errors.append(script_engine_hint())
        elif "not found" in reason.lower() or ".sty" in reason:
            result.errors.append(miktex_advice())
        kept = copy_log_beside(output_pdf, log_path)
        if kept:
            result.notes.append("полный журнал сборки: %s" % kept)
    else:
        try:
            shutil.copyfile(produced, output_pdf)
            result.pdf = output_pdf
        except OSError:
            result.errors.append(
                "не удалось записать %s: файл занят другой программой (открыт в "
                "просмотрщике?) — закройте его и повторите" % output_pdf.name)
    parse_log(log_path, result)
    if severe_overfull(log_path):
        result.notes.append(
            "подсказка: если за поле выходит таблица, оберните её в "
            "\\resizebox{\\textwidth}{!}{...} или задайте ей \\small/\\footnotesize; "
            "если формула — добавьте в неё \\allowbreak в месте разрыва")


def verify_docx(docx_path: Path, source_text: str, result: Result) -> None:
    """Проверить DOCX без Word: это zip с word/document.xml, оттуда берём текст и формулы."""
    import zipfile

    try:
        with zipfile.ZipFile(docx_path) as archive:
            if "word/document.xml" not in archive.namelist():
                result.errors.append("файл не похож на DOCX: внутри нет word/document.xml")
                return
            document = archive.read("word/document.xml").decode("utf-8", "replace")
    except (OSError, zipfile.BadZipFile) as error:
        result.errors.append("DOCX повреждён: %s" % error)
        return

    plain = " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", document))
    formulas = document.count("<m:oMath")
    result.notes.append("DOCX: формул Word — %d, таблиц — %d" % (formulas, document.count("<w:tbl>")))
    if not plain.strip():
        result.warnings.append("в DOCX не нашлось текста — проверьте документ в Word")
        return
    has_cyrillic_source = bool(re.search(r"[А-Яа-яЁё]", source_text))
    if has_cyrillic_source and not re.search(r"[А-Яа-яЁё]", plain):
        result.warnings.append("в DOCX не найден кириллический текст")
    source_math = len(re.findall(r"\$[^$]{1,400}?\$", source_text)) + source_text.count("$$") // 2
    if source_math > 3 and formulas == 0:
        result.warnings.append("в исходнике есть формулы, но в DOCX они не найдены — "
                               "проверьте, как они выглядит в Word")


def build_docx(source: Path, output_docx: Path, result: Result) -> None:
    """DOCX собирает pandoc: формулы становятся объектами Word, таблицы сохраняются."""
    pandoc = resolve_tool("pandoc")
    if not pandoc:
        result.errors.append("не найден pandoc — он нужен для формата DOCX")
        return
    if output_docx.exists():
        try:
            os.remove(output_docx)  # старый файл не должен сойти за свежий результат
        except OSError:
            result.errors.append(
                "не удалось обновить %s: файл занят другой программой (открыт в Word?) — "
                "закройте его и повторите" % output_docx.name)
            return
    run = subprocess.run(
        [pandoc, str(source), "-o", str(output_docx),
         "--resource-path", str(source.parent)],
        env=_tool_env(), capture_output=True, text=True, errors="replace",
        creationflags=NO_WINDOW)
    for line in run.stderr.splitlines():
        if line.startswith("[WARNING]"):
            result.warnings.append("pandoc: " + line.removeprefix("[WARNING]").strip())
    if run.returncode != 0 or not output_docx.exists():
        result.errors.append("pandoc не собрал DOCX: %s"
                             % (" ".join(run.stderr.split())[:300] or "файл не создан"))
        return
    result.docx = output_docx


def build_markdown(source: Path, output_pdf: Path, work_dir: Path, result: Result,
                   resource_extra: Path | None = None) -> None:
    pandoc = resolve_tool("pandoc")
    if not pandoc:
        result.errors.append("не найден pandoc — он нужен для файлов .md")
        return
    tex_path = work_dir / "doc.tex"
    resource = str(source.parent)
    if resource_extra is not None:
        resource += os.pathsep + str(resource_extra)
    pandoc_result = subprocess.run(
        [pandoc, str(source), "-s", "-t", "latex", "-o", str(tex_path),
         "--resource-path", resource] + PANDOC_ARGS,
        env=_tool_env(), capture_output=True, text=True, errors="replace",
        creationflags=NO_WINDOW)
    for line in pandoc_result.stderr.splitlines():
        if line.startswith("[WARNING]"):
            result.warnings.append("pandoc: " + line.removeprefix("[WARNING]").strip())
    if pandoc_result.returncode != 0 or not tex_path.exists():
        result.errors.append("pandoc не смог прочитать файл: %s"
                             % " ".join(pandoc_result.stderr.split())[:300])
        return
    build_latex(tex_path.read_text(encoding="utf-8"), output_pdf, work_dir, result)


# --- DOCX -> Markdown: восстановление формул ---------------------------------
# Транслятор MathType ("LaTeX (Cyrillic)", см. src/make_mathtype_translator.py)
# оставляет в docx обычный текст: инлайн "$...$", выключные "\[...\]" (многострочные
# размазаны по соседним абзацам), вокруг формул — служебные комментарии
# "% MathType!...". Здесь они собираются обратно в настоящую математику.

MATH_OPEN = chr(92) + "["
MATH_CLOSE = chr(92) + "]"
COMMENT_TOKEN = re.compile(r"% MathType!.*?!")


def _ast_text(inline: list) -> str:
    """Плоский текст инлайновых элементов pandoc-AST (для анализа абзацев)."""
    parts = []
    for element in inline:
        kind = element.get("t")
        if kind == "Str":
            parts.append(element["c"])
        elif kind in ("Space", "SoftBreak", "Tab"):
            parts.append(" ")
    return "".join(parts)


def restore_docx_math(blocks: list) -> tuple[list, list, int]:
    """Собрать текстовые формулы docx в математические узлы pandoc-AST.

    Возвращает (новые блоки, предупреждения, сколько формул восстановлено).
    """
    warnings: list[str] = []
    result: list = []
    restored = 0
    last_was_math = False
    index = 0
    while index < len(blocks):
        block = blocks[index]
        if block.get("t") != "Para":
            result.append(block)
            index += 1
            continue
        text = re.sub(COMMENT_TOKEN, "", _ast_text(block["c"])).strip()
        if not text or MATHYPE_TAIL.match(text):
            # пусто после чистки комментариев или осиротевший хвост "% MathType!End!2!1!":
            # абзац из одной OMML-формулы сохраняем; формулу считаем «последней»,
            # чтобы её номер в следующем абзаце не выглядел потерянной формулой
            if any(element.get("t") == "Math" for element in block["c"]):
                result.append(block)
                last_was_math = True
            index += 1
            continue
        if text.startswith("% MathType!"):
            index += 1  # служебный комментарий транслятора целиком
            continue

        # выключная формула: "\[" в этом или нескольких следующих абзацах
        if text.startswith(MATH_OPEN):
            parts, stop = [], index
            while stop < len(blocks) and blocks[stop].get("t") == "Para":
                part = re.sub(COMMENT_TOKEN, "", _ast_text(blocks[stop]["c"])).strip()
                if part:
                    parts.append(part)
                if MATH_CLOSE in part:
                    break
                stop += 1
            joined = " ".join(parts)
            match = re.match(re.escape(MATH_OPEN) + r"(.*)" + re.escape(MATH_CLOSE) + r"(.*)$",
                             joined, re.S)
            if match:
                result.append({"t": "Para", "c": [
                    {"t": "Math", "c": [{"t": "DisplayMath"}, match.group(1).strip()]}]})
                tail = re.sub(r"\s+", " ", match.group(2)).strip()
                if tail:
                    result.append({"t": "Para", "c": [{"t": "Str", "c": tail}]})
                restored += 1
                last_was_math = True
                index = stop + 1
                continue

        # формула "$...$" на весь абзац (номер формулы — хвостом)
        single = re.match(r"^\$(.+)\$(.*)$", text, re.S)
        if single and chr(92) in single.group(1):
            result.append({"t": "Para", "c": [
                {"t": "Math", "c": [{"t": "DisplayMath"}, single.group(1).strip()]}]})
            tail = re.sub(r"\s+", " ", single.group(2)).strip()
            if tail:
                result.append({"t": "Para", "c": [{"t": "Str", "c": tail}]})
            restored += 1
            last_was_math = True
            index += 1
            continue

        # инлайн "$...$" внутри обычного текста; абзацы с настоящей математикой
        # (OMML) не трогаем — их текст восстанавливать не нужно
        if "$" in text and not any(element.get("t") == "Math" for element in block["c"]):
            pieces = []
            for part in re.split(r"(\$[^$]+\$)", text):
                if len(part) > 2 and part.startswith("$") and part.endswith("$"):
                    pieces.append(("math", part.strip("$").strip()))
                elif part:
                    for chunk in re.split(r"([ \t]+)", part):
                        pieces.append(("space", "") if chunk.isspace() else ("str", chunk))
            inlines = []
            for kind, value in pieces:
                if kind == "space":
                    if inlines and inlines[-1].get("t") != "Space":
                        inlines.append({"t": "Space"})
                elif kind == "math":
                    inlines.append({"t": "Math", "c": [{"t": "InlineMath"}, value]})
                elif value:
                    inlines.append({"t": "Str", "c": value})
            while inlines and inlines[-1].get("t") == "Space":
                inlines.pop()
            if inlines:
                result.append({"t": "Para", "c": inlines})
                restored += 1
                last_was_math = True
                index += 1
                continue

        # одинокий номер формулы — след потерянной при пакетной конвертации MathType.
        # Номер СРАЗУ после восстановленной формулы — это нормальная вёрстка
        # (в исходниках номер живёт отдельным абзацем), а не потеря.
        if re.fullmatch(r"\(?\d+(?:\.\d+)?\)?", text) and not last_was_math:
            warnings.append(
                "в документе остался номер формулы (%s) без самой формулы — MathType, "
                "вероятно, не смог её перевести; сверьте с исходником" % text)

        result.append(block)
        last_was_math = False
        index += 1

    for block in result:
        if block.get("t") != "Para":
            continue
        for inline in block["c"]:
            if inline.get("t") == "Math":
                inline["c"][1] = cleanup_math_code(inline["c"][1])
    return result, warnings, restored


# Греческие команды -> юникод: под unicode-math (xelatex) команды вроде \alpha
# уводят глиф в текстовый шрифт, где математики нет («Missing character»), а
# буквальные греческие буквы рендерятся любым шрифтом.
GREEK_LETTERS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ", "sigma": "σ",
    "varsigma": "ς", "tau": "τ", "upsilon": "υ", "phi": "φ", "varphi": "φ",
    "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}


def cleanup_math_code(code: str) -> str:
    """Починить math-содержимое из docx: старый \\rm и греческие команды."""
    code = code.replace("\\rm{", "\\mathrm{")
    for name, ch in GREEK_LETTERS.items():
        code = re.sub(r"\\" + name + r"(?![a-zA-Z])", ch, code)
    # одиночная буква под \mathrm: unicode-math уводит её в текстовый шрифт
    code = re.sub(r"\\mathrm\{([^{}])\}", lambda m: m.group(1), code)
    return code


def merge_char_texts(markdown: str) -> str:
    """Склеить посимвольные \\text{н}\\text{а}\\text{к}\\text{р} в \\text{накр}."""
    pattern = re.compile(r"\\text\{([^{}]*)\}\\text\{([^{}]*)\}")
    previous = None
    while previous != markdown:
        previous = markdown
        markdown = pattern.sub(
            lambda match: "\\text{" + match.group(1) + match.group(2) + "}", markdown)
    return markdown


MARK_SPAN = re.compile(r"\[([^\]\n]+)\]\{\.mark\}")
# ссылка на картинку в markdown: обычная или в угловых скобках (путь с пробелами)
IMAGE_LINK = re.compile(r"!\[[^\]]*\]\(\s*(<[^>]*>|[^)\s]+)")
# то же, но с захватом всей ссылки и возможных атрибутов {width=...} — для замены
IMAGE_LINK_FULL = re.compile(r"!\[[^\]]*\]\(\s*(<[^>]*>|[^)\s]+)\s*\)(\{[^}]*\})?")
# хвост разорванного Word'ом комментария MathType вида "% MathType!End!2!1!"
MATHYPE_TAIL = re.compile(r"^\s*\d+!\d+!\s*$")


def strip_mark_spans(markdown: str) -> str:
    """Убрать выделения маркером [...]{.mark}, оставив текст.

    Для PDF-сборки: pandoc рисует их пакетом soul, а soul не работает с кириллицей
    под xelatex (в готовом PDF теряются глифы). В md-результате выделения сохраняются.
    """
    return MARK_SPAN.sub(lambda m: m.group(1), markdown)


def drop_mathtype_markup(markdown: str) -> str:
    """Выбросить служебные строки MathType: комментарии и их осиротевшие хвосты."""
    return "\n".join(
        line for line in markdown.split("\n")
        if not line.lstrip().startswith(("% MathType!", chr(92) + "% MathType!"))
        and not MATHYPE_TAIL.match(line))


def image_links(text: str) -> list[str]:
    """Пути картинок, на которые ссылается markdown (URL и data: пропускаются)."""
    links: list[str] = []
    for link in IMAGE_LINK.findall(text):
        link = link.strip("<>")
        if "://" in link or link.startswith("data:"):
            continue
        value = unquote(link)
        if value not in links:
            links.append(value)
    return links


def missing_images(source: Path) -> list[str]:
    """Картинки, на которые ссылается markdown, но которых нет рядом с файлом."""
    try:
        text = source.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return []
    missing: list[str] = []
    for link in image_links(text):
        path = Path(link)
        candidate = path if path.is_absolute() else source.parent / path
        if not candidate.is_file() and link not in missing:
            missing.append(link)
    return missing


def replace_missing_images(markdown: str, source_dir: Path) -> str:
    """Заменить ссылки на отсутствующие картинки текстом-заглушкой.

    LaTeX на отсутствующем файле картинки падает целиком, поэтому в PDF-версии
    вместо пропавшей картинки ставится пометка, а сборка продолжается.
    """
    def swap(match):
        link = match.group(1).strip("<>")
        if "://" in link or link.startswith("data:"):
            return match.group(0)
        path = Path(unquote(link))
        candidate = path if path.is_absolute() else source_dir / path
        if candidate.is_file():
            return match.group(0)
        return "*[картинка не найдена: %s]*" % link
    return IMAGE_LINK_FULL.sub(swap, markdown)


def copy_referenced_media(source: Path, work_dir: Path) -> None:
    """Скопировать картинки, на которые ссылается markdown, в каталог сборки.

    latexmk компилирует во временном каталоге и относительные пути к картинкам
    оттуда не находит, поэтому переносим их, сохраняя структуру путей.
    """
    try:
        text = source.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return
    for link in image_links(text):
        path = Path(link)
        if path.is_absolute():
            continue  # такие картинки latexmk найдёт по абсолютному пути как есть
        origin = source.parent / path
        if not origin.is_file():
            continue
        target = work_dir / path
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(origin, target)
        except OSError:
            continue


def docx_media_dir(stem: str) -> str:
    """Имя папки для картинок docx-документа: рядом с итоговым md, без пробелов
    (LaTeX давится \\includegraphics с пробелами в пути)."""
    return re.sub(r"\s+", "_", stem) + ".media"


def build_markdown_from_docx(source: Path, output_md: Path, result: Result) -> None:
    """DOCX -> Markdown с восстановлением формул; картинки уходят в <имя>.media."""
    pandoc = resolve_tool("pandoc")
    if not pandoc:
        result.errors.append("не найден pandoc — он нужен для файлов .docx")
        return
    # без пробелов в имени: LaTeX давится \\includegraphics с пробелами в пути
    media_dir = docx_media_dir(source.stem)
    json_run = subprocess.run(
        [pandoc, str(source.resolve()), "-t", "json", "--extract-media=" + media_dir],
        cwd=str(output_md.parent), env=_tool_env(), capture_output=True,
        creationflags=NO_WINDOW)
    if json_run.returncode != 0:
        result.errors.append("pandoc не смог прочитать DOCX: %s"
                             % json_run.stderr.decode("utf-8", "replace")[:300])
        return
    try:
        ast = json.loads(json_run.stdout.decode("utf-8"))
    except ValueError as error:
        result.errors.append("pandoc вернул нечитаемое дерево документа: %s" % error)
        return

    blocks, warnings, restored = restore_docx_math(ast.get("blocks", []))
    for warning in warnings:
        result.warnings.append(warning)
    ast["blocks"] = blocks

    md_run = subprocess.run(
        [pandoc, "-f", "json", "-t", "markdown", "--wrap=none"],
        input=json.dumps(ast, ensure_ascii=False), env=_tool_env(),
        capture_output=True, text=True, encoding="utf-8", creationflags=NO_WINDOW)
    if md_run.returncode != 0:
        result.errors.append("pandoc не собрал Markdown: %s"
                             % md_run.stderr[:300])
        return

    text = merge_char_texts(md_run.stdout)
    # комментарии MathType выбрасываются; Word иногда рвёт "% MathType!End!2!1!" на две
    # строки, поэтому отдельно ловим осиротевшие хвосты вида "2!1!"
    text = drop_mathtype_markup(text)
    # 𝛼, 𝜋 и прочие Unicode-символы math-алфавитов (U+1D400+) не есть в обычных шрифтах;
    # NFKC возвращает их к обычным греческим/латинским буквам
    text = unicodedata.normalize("NFKC", text)
    try:
        output_md.write_text(text.rstrip() + "\n", encoding="utf-8")
    except OSError as error:
        result.errors.append("не удалось записать Markdown: %s" % error)
        return
    result.md = output_md

    media_root = output_md.parent / media_dir
    images = len(list(media_root.rglob("*"))) if media_root.is_dir() else 0
    note = "DOCX: формул восстановлено — %d" % restored
    if images:
        note += ", картинок — %d (папка %s)" % (images, media_dir)
    result.notes.append(note)


def build_markdown_from_tex(source: Path, output_md: Path, result: Result) -> None:
    """LaTeX -> Markdown тем же pandoc (для формата md у .tex-исходников)."""
    pandoc = resolve_tool("pandoc")
    if not pandoc:
        result.errors.append("не найден pandoc — он нужен для формата md")
        return
    run = subprocess.run(
        [pandoc, str(source), "-f", "latex", "-t", "markdown", "--wrap=none",
         "-o", str(output_md)],
        env=_tool_env(), capture_output=True, text=True, errors="replace",
        creationflags=NO_WINDOW)
    if run.returncode != 0 or not output_md.exists():
        result.errors.append("pandoc не собрал Markdown: %s"
                             % " ".join(run.stderr.split())[:300])
        return
    result.md = output_md


def convert_file(source: Path, output_dir: Path | None = None,
                 formats: tuple[str, ...] = DEFAULT_FORMATS) -> Result:
    """Сконвертировать файл в выбранные форматы; результат рядом с исходником или в output_dir."""
    source = Path(source)
    requested = tuple(name for name in FORMATS if name in formats)
    result = Result(source=source, requested=requested or DEFAULT_FORMATS)

    suffix = source.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        result.errors.append("неподдерживаемый формат %s (нужны .md, .markdown, .tex, "
                             ".latex, .docx)" % (suffix or "без расширения"))
        return result
    if not source.exists():
        result.errors.append("файл не найден")
        return result

    targets = {name: (Path(output_dir) / (source.stem + "." + name)) if output_dir
               else source.with_suffix("." + name)
               for name in result.requested}
    # выходной DOCX не должен затирать сам исходник: normalization docx -> docx
    # пишется отдельной копией, оригинал остаётся нетронутым
    converted_copy = False
    if suffix == ".docx" and "docx" in targets and targets["docx"] == source:
        targets["docx"] = source.with_name(source.stem + ".converted.docx")
        converted_copy = True
    # у .docx исходник бинарный: текстом для проверок служит восстановленный markdown
    source_text = ""
    if suffix != ".docx":
        source_text = source.read_text(encoding="utf-8", errors="replace")
    # markdown может ссылаться на отсутствующие картинки: конвертируем, но говорим,
    # каких файлов не хватает — иначе потеря видна только по пустому месту в PDF
    if suffix in (".md", ".markdown"):
        missing = missing_images(source)
        if missing:
            shown = ", ".join(missing[:MAX_ISSUES])
            if len(missing) > MAX_ISSUES:
                shown += " и ещё %d" % (len(missing) - MAX_ISSUES)
            result.warnings.append(
                "не найдены картинки, на которые ссылается файл: %s — конвертация "
                "продолжена без них (проверьте, что папка с картинками лежит рядом с .md)"
                % shown)

    work_dir = Path(tempfile.mkdtemp(prefix="pdfconv-"))
    try:
        docx_md: Path | None = None
        if suffix == ".docx":
            docx_md = Path(targets.get("md", work_dir / (source.stem + ".md")))
            build_markdown_from_docx(source, docx_md, result)
            if docx_md.exists():
                source_text = docx_md.read_text(encoding="utf-8")
                if "md" in result.requested:
                    result.md = docx_md

        if "pdf" in targets:
            if suffix in (".md", ".markdown"):
                # картинки и чистка — как в docx-маршруте: latexmk собирает во временном
                # каталоге, а .mark-выделения через soul теряют кириллические глифы
                copy_referenced_media(source, work_dir)
                pdf_source = work_dir / (source.stem + ".for-pdf.md")
                pdf_text = drop_mathtype_markup(strip_mark_spans(source_text))
                pdf_text = replace_missing_images(pdf_text, source.parent)
                pdf_source.write_text(pdf_text, encoding="utf-8")
                build_markdown(pdf_source, targets["pdf"], work_dir, result)
            elif suffix == ".docx":
                if docx_md is not None and docx_md.exists():
                    # .mark-выделения убираем только из PDF-версии: soul ломает кириллицу
                    pdf_source = work_dir / (source.stem + ".for-pdf.md")
                    pdf_source.write_text(
                        strip_mark_spans(docx_md.read_text(encoding="utf-8")), encoding="utf-8")
                    media_src = docx_md.parent / docx_media_dir(source.stem)
                    target_media = work_dir / media_src.name
                    # если md собран во временном каталоге, картинки уже лежат там;
                    # копировать их сами в себя нельзя — WinError 32
                    if media_src.is_dir() and media_src != target_media:
                        shutil.copytree(media_src, target_media,
                                        dirs_exist_ok=True)  # latexmk ищет картинки от cwd
                    build_markdown(pdf_source, targets["pdf"], work_dir, result)
            else:
                # .tex: каталог исходника нужен для относительных картинок и \input
                build_latex(source_text, targets["pdf"], work_dir, result,
                            source.parent.resolve())
            if result.pdf:
                verify_pdf(result.pdf, source_text, result)
        if "docx" in targets:
            docx_input = docx_md if suffix == ".docx" and docx_md is not None \
                and docx_md.exists() else source
            build_docx(docx_input, targets["docx"], result)
            if result.docx:
                if converted_copy:
                    result.notes.append("исходный DOCX не изменялся; копия с формулами Word — %s"
                                        % result.docx.name)
                verify_docx(result.docx, source_text, result)
        if "md" in targets and result.md is None:
            if suffix in (".md", ".markdown"):
                if targets["md"] == source:
                    # исходник уже markdown: копия самого себя не нужна
                    result.md = source
                    result.notes.append("исходник уже в формате md — копия не создаётся")
                else:
                    shutil.copyfile(source, targets["md"])
                    result.md = targets["md"]
            elif suffix in (".tex", ".latex"):
                build_markdown_from_tex(source, targets["md"], result)
    except Exception as error:  # noqa: BLE001
        result.errors.append("непредвиденная ошибка: %s: %s" % (type(error).__name__, error))
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    for name in result.missing_formats:
        result.errors.append("%s не создан" % name.upper())
    return result


def summarize(results: list[Result]) -> list[str]:
    """Те самые три строки статуса плюс расшифровка."""
    total = len(results)
    clean = [item for item in results if item.status == "ok"]
    warned = [item for item in results if item.status == "warnings"]
    failed = [item for item in results if item.status == "error"]
    return [
        "Сконвертировано %d из %d без ошибок" % (len(clean), total),
        "Сконвертировано %d из %d с предупреждениями" % (len(warned), total),
        "Сконвертировано %d из %d с ошибками" % (len(failed), total),
    ]


def environment_report() -> str:
    """Текстовый отчёт об окружении (для режима командной строки)."""
    lines = []
    for item in check_environment():
        mark = "OK  " if item.ok else "НЕТ "
        lines.append("%s%s: %s" % (mark, item.title, item.detail))
        if not item.ok and item.fix:
            lines.append("      что сделать: %s" % item.fix)
        if not item.ok and item.url:
            lines.append("      откуда взять: %s" % item.url)
    return "\n".join(lines)


def parse_formats(value: str) -> tuple[str, ...]:
    """Разобрать список форматов из командной строки: pdf, docx, all."""
    if value.strip().lower() in ("all", "*"):
        return FORMATS
    names = tuple(part.strip().lower() for part in value.split(",") if part.strip())
    unknown = [name for name in names if name not in FORMATS]
    if unknown:
        raise SystemExit("Неизвестный формат: %s (доступны %s)" % (", ".join(unknown),
                                                                   ", ".join(FORMATS)))
    return tuple(name for name in FORMATS if name in names) or FORMATS


def utf8_console() -> None:
    """Печатать в UTF-8 и не падать на символах вне локальной кодировки.

    В Python до 3.15 (например, 3.14) при перенаправленном выводе берётся локаль Windows
    (cp1251), и «→» в подсказках даёт UnicodeEncodeError; в 3.15 UTF-8-режим включён по
    умолчанию. Настройка безвредна и для 3.15 — просто страхует от такого запуска.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:  # noqa: BLE001
                pass


def main(argv: list[str]) -> int:
    """Командная строка: --check (окружение) и конвертация перечисленных файлов."""
    utf8_console()
    formats = DEFAULT_FORMATS
    arguments: list[str] = []
    index = 0
    while index < len(argv):
        argument = argv[index]
        if argument in ("-f", "--formats"):
            index += 1
            if index >= len(argv):
                raise SystemExit("После %s нужен список форматов, например: --formats pdf,docx"
                                 % argument)
            formats = parse_formats(argv[index])
        elif argument.startswith("--formats="):
            formats = parse_formats(argument.split("=", 1)[1])
        else:
            arguments.append(argument)
        index += 1

    if not arguments or arguments[0] in ("-h", "--help"):
        print(__doc__)
        program = Path(sys.argv[0]).name or "convert.py"
        print("Использование: py -3 %s [--formats pdf,docx] [--check] файл.md [файл.tex ...]"
              % program)
        print("Форматы по умолчанию: %s" % ", ".join(DEFAULT_FORMATS))
        return 0
    if arguments[0] == "--check":
        print(environment_report())
        return 0

    results = [convert_file(Path(argument), formats=formats) for argument in arguments]
    for result in results:
        mark = {"ok": "OK", "warnings": "ВНИМАНИЕ", "error": "ОШИБКА"}[result.status]
        produced = [str(getattr(result, name)) for name in result.requested
                    if getattr(result, name)] or ["ничего не создано"]
        print("%s: %s -> %s" % (mark, result.source.name, ", ".join(produced)))
        for message in result.errors:
            print("    ошибка: %s" % message)
        for message in result.warnings:
            print("    предупреждение: %s" % message)
        for message in result.notes:
            print("    примечание: %s" % message)
    print()
    for line in summarize(results):
        print(line)
    return 1 if any(result.status == "error" for result in results) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
"""Подготовить проект на новой машине: зависимости и проверка состава.

Запуск: C:\\Tools\\Scripts\\Python314\\python.exe src\\bootstrap.py [--install]
  без --install  — только проверка и подсказки, ничего не меняет;
  с --install    — ставит недостающие python-зависимости в lib\\ тем же
                   интерпретатором, которым запущен скрипт (так колёса
                   всегда соответствуют версии Python).

Проверяет: tkinterdnd2, хотя бы один бэкенд проверки PDF (pymupdf/pdfminer/pypdf),
pandoc в bin\\pandoc, Perl в bin\\perl (или в PATH), MiKTeX (latexmk).
Код возврата 0 — всё на месте, 1 — чего-то не хватает.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

CODE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = CODE_DIR.parent
LIB_DIR = PROJECT_DIR / "lib"
BIN_DIR = PROJECT_DIR / "bin"

PIP_COMMAND = ("%s -m pip install --target \"%s\" -r requirements.txt"
               % (Path(sys.executable).name, LIB_DIR))
REQUIREMENTS = PROJECT_DIR / "requirements.txt"


def utf8_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:  # noqa: BLE001
                pass


def can_import(module: str) -> bool:
    try:
        __import__(module)
    except Exception:  # noqa: BLE001 — битые бинарные зависимости тоже сюда
        return False
    return True


def check(title: str, ok: bool, hint: str = "") -> bool:
    print("%s %s%s" % ("OK  " if ok else "НЕТ ", title, ("" if ok or not hint else " — " + hint)))
    return ok


def main(argv: list[str]) -> int:
    utf8_console()
    sys.path.insert(0, str(LIB_DIR))
    install = "--install" in argv

    print("Python: %s (%s)" % (sys.version.split()[0], sys.executable))
    print("Проект: %s" % PROJECT_DIR)
    print()

    problems = 0
    print("1. Python-зависимости в lib\\")
    missing = []
    if not can_import("tkinterdnd2"):
        missing.append("tkinterdnd2")
        problems += 1
        print("НЕТ  tkinterdnd2 (перетаскивание файлов) — %s" % PIP_COMMAND)
    else:
        print("OK   tkinterdnd2")
    backends = [name for name in ("pymupdf", "pdfminer", "pypdf") if can_import(name)]
    if backends:
        print("OK   проверка PDF: %s" % ", ".join(backends))
    else:
        missing.extend(["pymupdf", "pdfminer.six", "pypdf"])
        problems += 1
        print("НЕТ  бэкенд проверки PDF (pymupdf/pdfminer/pypdf) — %s" % PIP_COMMAND)
    if missing and install:
        command = [sys.executable, "-m", "pip", "install", "--target", str(LIB_DIR)]
        if REQUIREMENTS.exists():
            command += ["-r", str(REQUIREMENTS)]
        else:
            command += ["tkinterdnd2", "pymupdf", "pdfminer.six", "pypdf"]
        print("\nУстанавливаю: %s" % " ".join(command))
        code = subprocess.call(command)
        print("pip завершился с кодом %d" % code)
        if code != 0:
            problems += 1
    print()

    print("2. Внешние программы")
    pandoc = BIN_DIR / "pandoc" / "pandoc.exe"
    if not check("pandoc: %s" % pandoc, pandoc.is_file(),
                 "скачайте pandoc-*-windows-x86_64.zip с https://github.com/jgm/pandoc/releases "
                 "и распакуйте pandoc.exe в bin\\pandoc\\"):
        problems += 1
    perl = BIN_DIR / "perl" / "bin" / "perl.exe"
    if not check("Perl для latexmk: %s" % perl,
                 perl.is_file() or shutil.which("perl") is not None,
                 "скачайте portable Strawberry Perl со https://strawberryperl.com/ и положите "
                 "perl\\bin и perl\\lib в bin\\perl\\ (или добавьте Perl в PATH)"):
        problems += 1
    latexmk = shutil.which("latexmk")
    if not check("MiKTeX (latexmk): %s" % (latexmk or "не найден"),
                 latexmk is not None,
                 "установите MiKTeX с https://miktex.org/download (нужен для PDF)"):
        problems += 1
    print()

    if problems:
        print("Итог: не хватает %d компонент(ов) — читайте подсказки выше." % problems)
        print("После установки: C:\\Tools\\Scripts\\Python314\\python.exe src\\bootstrap.py  (снова проверить)")
        return 1
    print("Итог: всё на месте. Дальше: C:\\Tools\\Scripts\\Python314\\python.exe src\\selfcheck.py, затем pdfconv.cmd")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

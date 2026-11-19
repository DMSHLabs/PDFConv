"""Самопроверка проекта: окружение, пробная конвертация, наличие файлов.

Запуск: C:\\Tools\\Scripts\\Python314\\python.exe src/selfcheck.py
Код возврата 0 — всё в порядке; 1 — есть проблемы (подробности в выводе).

Ничего не меняет в проекте: пробная конвертация идёт во временном каталоге.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

CODE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = CODE_DIR.parent
sys.path.insert(0, str(CODE_DIR))
sys.path.insert(0, str(PROJECT_DIR / "lib"))  # зависимости, если запуск через Python

import convert  # noqa: E402

convert.utf8_console()

TEST_NAME = "Самопроверка"
TEST_MD = """# Самопроверка

Кириллица, формула $E = mc^2$ и таблица.

| Величина | Формула |
|----------|---------|
| скорость | $v = \\frac{s}{t}$ |
"""


def main() -> int:
    problems: list[str] = []
    print("=" * 70)
    print("1. Окружение")
    print("=" * 70)
    for check in convert.check_environment():
        print("  %s%s: %s" % ("OK  " if check.ok else "НЕТ ", check.title, check.detail))
        if not check.ok:
            if check.required:
                problems.append(check.title)
            if check.fix:
                print("       что сделать: %s" % check.fix)

    print()
    print("=" * 70)
    print("2. Пробная конвертация: markdown → PDF и DOCX")
    print("=" * 70)
    work = Path(tempfile.mkdtemp(prefix="selfcheck-"))
    source = work / (TEST_NAME + ".md")
    source.write_text(TEST_MD, encoding="utf-8")
    result = convert.convert_file(source, formats=("pdf", "docx"))

    for message in result.errors:
        print("  ошибка: %s" % message)
    for message in result.warnings:
        print("  предупреждение: %s" % message)
    for message in result.notes:
        print("  примечание: %s" % message)

    print("  статус: %s" % {"ok": "без ошибок", "warnings": "с предупреждениями",
                            "error": "с ошибками"}[result.status])
    if result.status == "error":
        problems.append("пробная конвертация")

    if result.pdf and result.pdf.exists():
        try:
            data = convert.pdf_document_text(result.pdf)
        except Exception:  # noqa: BLE001
            data = None
        if data is None:
            print("  PDF: создан, но проверить его нечем (нет pymupdf/pdfminer/pypdf)")
        else:
            text, pages = data
            searchable = TEST_NAME in text
            print("  PDF: %d стр., текстовый слой %s" % (
                pages, "ищется" if searchable else "НЕ ищется"))
            if not searchable:
                problems.append("текстовый слой PDF")

    if result.docx and result.docx.exists():
        import zipfile
        with zipfile.ZipFile(result.docx) as archive:
            document_xml = archive.read("word/document.xml").decode("utf-8", "replace")
        formulas = document_xml.count("<m:oMath")
        print("  DOCX: формул Word — %d, таблиц — %d" % (formulas, document_xml.count("<w:tbl>")))
        if formulas == 0:
            problems.append("формулы в DOCX")

    # DOCX -> MD: формула должна пережить круг (маршрут MathType-документов)
    try:
        probe_md = work / "probe.md"
        probe_md.write_text("# Проверка docx\n\nФормула $E = mc^2$ и текст.\n", encoding="utf-8")
        probe_docx = work / "probe.docx"
        make = convert.Result(source=probe_md, requested=("docx",))
        convert.build_docx(probe_md, probe_docx, make)
        back = convert.Result(source=probe_docx, requested=("md",))
        probe_out = work / "probe_out.md"
        convert.build_markdown_from_docx(probe_docx, probe_out, back)
        errors = make.errors + back.errors
        survived = (not errors and probe_out.exists()
                    and any(formula in probe_out.read_text(encoding="utf-8")
                            for formula in ("E = mc^2", "E = mc^{2}")))
        print("  %s docx -> md: формула пережила круг" % ("OK  " if survived else "НЕТ "))
        if not survived:
            print("       " + "; ".join(errors)[:200])
            problems.append("docx -> md")
    except Exception as error:  # noqa: BLE001
        print("  НЕТ  docx -> md: %s: %s" % (type(error).__name__, error))
        problems.append("docx -> md")
    shutil.rmtree(work, ignore_errors=True)

    print()
    print("=" * 70)
    print("3. Файлы проекта")
    print("=" * 70)
    expected = [
        ("PDFConv.exe", "самостоятельный вариант без Python (не обязателен)"),
        ("pdfconv.cmd", "запуск окна через Python"),
        ("bin/pandoc/pandoc.exe", "pandoc: нужен для .md и любого DOCX"),
        ("bin/perl/bin/perl.exe", "вендорный Perl для latexmk (MiKTeX 26.5+): обновляется "
                                  "заменой bin\\perl (Strawberry portable)"),
        ("lib", "python-зависимости для запуска через Python"),
        ("res/PDFConv.ico", "иконка окна и exe"),
        ("src/convert.py", "конвейеры сборки"),
        ("src/pdfconv.pyw", "окно"),
        ("README.md", "описание для пользователя"),
        ("AGENTS.md", "правила работы"),
        ("CONTEXT.md", "полное описание проекта и проверок"),
    ]
    for relative, description in expected:
        path = PROJECT_DIR / relative
        if path.exists():
            size = path.stat().st_size
            print("  OK   %-24s %s" % (relative, "%d МБ" % (size // 1048576) if size > 1048576
                                       else description))
        else:
            print("  НЕТ  %-24s %s" % (relative, description))
            if relative in ("pdfconv.cmd", "bin/pandoc/pandoc.exe", "src/convert.py"):
                problems.append("нет файла %s" % relative)

    print()
    print("=" * 70)
    if problems:
        print("ИТОГ: есть проблемы — %s" % "; ".join(problems))
        print("Подробности и что делать — в CONTEXT.md, раздел «Известные проблемы».")
    else:
        print("ИТОГ: всё на месте, конвертация проверена — проект готов к работе.")
    print("=" * 70)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
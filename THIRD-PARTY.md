# Сторонние компоненты и их лицензии

Проект распространяется под GNU GPL v3 или более поздней версии (см. `LICENSE`).
Ниже — всё, что проект использует, и лицензии этих компонентов.

## Что в репозитории и что ставится отдельно

| Компонент | Версия | Лицензия | Как используется |
|-----------|--------|----------|------------------|
| Python | 3.10+ (на машине 3.15) | Python Software Foundation License | интерпретатор для `pdfconv.cmd` и CLI |
| pandoc | 3.11 | GPL-2.0-or-later | `bin/pandoc/pandoc.exe` — конвертация `.md`/`.tex`/DOCX (скачивается отдельно, не в git) |
| Strawberry Perl | 5.42.3.1 (portable) | Artistic-1.0-Perl OR GPL-1.0-or-later | `bin/perl/` — `latexmk` в MiKTeX 26.5+ (скачивается отдельно, не в git) |
| MiKTeX | системный | собственные свободные лицензии | не распространяется; нужен для PDF |
| tkinterdnd2 | 0.6.3 | MIT | перетаскивание файлов (в `lib/`, не в git) |
| PyMuPDF (pymupdf) | 1.28.2 | **AGPL-3.0 или коммерческая Artifex** | необязательная проверка готового PDF (в `lib/`, не в git) |
| pdfminer.six | 20260107 | MIT | резервная проверка PDF (в `lib/`, не в git) |
| pypdf | 7.1.1 | BSD-3-Clause | резервная проверка PDF (в `lib/`, не в git) |
| cryptography | 50.0.2 | BSD-3-Clause OR Apache-2.0 | зависимость pdfminer (в `lib/`, не в git) |
| cffi | 2.1.1 | MIT-0 | зависимость cryptography (в `lib/`, не в git) |
| charset-normalizer | 3.5.2 | MIT | зависимость pdfminer (в `lib/`, не в git) |
| pycparser | 3.12 | BSD-3-Clause | зависимость cffi (в `lib/`, не в git) |
| Pillow | build-time | MIT-CMU | только `src/make_icon.py` при перерисовке иконки; в дистрибутив не входит |
| PyInstaller | 6.22.3 | GPL-2.0-or-later с исключением для собранных приложений | только пересборка exe; в дистрибутив не входит |
| Шрифты Times New Roman, Consolas | системные | проприетарные Microsoft | не распространяются; используются при сборке PDF |

## Совместимость с GPL

Все компоненты, которые попадают в репозиторий или в сборку exe, совместимы с GPL-3:
MIT, BSD-3-Clause, Apache-2.0 и MIT-0 — permissive; pandoc (GPL-2.0-or-later) и Perl
(Artistic/GPL-1.0+) — свободные лицензии с обратной совместимостью по версии.

## PyMuPDF и AGPL — важное

PyMuPDF распространяется под **AGPL-3.0** (или по коммерческой лицензии Artifex). В репозитории
его нет: `lib/` не версионируется. Проверка PDF работает и без него — через pdfminer.six (MIT)
или pypdf (BSD); лестница бэкендов выбирает первый доступный.

Если вы распространяете **бинарник, внутри которого есть pymupdf** (например, собранный
`PDFConv.exe` текущим `src\build-exe.cmd`), такой бинарник нужно распространять на условиях
**AGPL-3.0**, а не GPL-3. Для релиза под GPL-3 exe собирается без pymupdf.

## Исходники сторонних компонентов

- pandoc: <https://github.com/jgm/pandoc> (та же версия — тег `3.11`)
- Strawberry Perl: <https://strawberryperl.com/>; Perl — <https://www.perl.org/>
- PyMuPDF: <https://github.com/pymupdf/PyMuPDF> (AGPL-3.0)
- pdfminer.six: <https://github.com/pdfminer/pdfminer.six> (MIT)
- pypdf: <https://github.com/py-pdf/pypdf> (BSD-3-Clause)
- tkinterdnd2: <https://pypi.org/project/tkinterdnd2/> (MIT)

Если вы распространяете папку/архив, где лежат pandoc и Perl (GPL/Artistic), приложите их
лицензии и укажите эти ссылки на исходники — этого требует GPL-2 у pandoc и Artistic License
у Perl.

## MathType-транслятор

`res/LaTeX Cyrillic.tdl` — производная от `LaTeX.tdl` Design Science; их лицензия разрешает
производные «solely for use with MathType», поэтому файл **не публикуется** (исключён в
`.gitignore`). В репозитории есть только генератор `src/make_mathtype_translator.py` — он
воспроизводит файл локально.

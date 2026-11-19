@echo off
rem Собрать PDFConv.exe - самостоятельный файл: Python внутрь, pandoc ищется рядом с exe.
rem Запуск: src\build-exe.cmd   (нужен PyInstaller: "C:\Tools\Scripts\Python314\python.exe" -m pip install pyinstaller)
rem Результат: PDFConv.exe в корне проекта, рядом с bin\, lib\ и res\.
rem Все пути абсолютные: с --specpath относительные считались бы от каталога build.
set "ROOT=%~dp0.."
cd /d "%ROOT%" || exit /b 1

if not exist "%ROOT%\res\PDFConv.ico" (
  echo Рисую иконку...
  "C:\Tools\Scripts\Python314\python.exe" "%ROOT%\src\make_icon.py" || exit /b 1
)

echo Собираю PDFConv.exe (это занимает пару минут)...
rem --paths lib: tkinterdnd2 и pymupdf лежат в проекте, а не в site-packages сборщика,
rem поэтому PyInstaller надо показать, где их искать, иначе они не попадут в exe.
rem --exclude-module: --collect-all pymupdf тянет за собой необязательные зависимости
rem из глобального Python (pyarrow, numpy, pandas, PIL), конвейеру они не нужны -
rem без исключений exe выходит 85 МБ вместо 32 (см. HANDOFF, <Окно и exe>).
"C:\Tools\Scripts\Python314\python.exe" -m PyInstaller --noconfirm --onefile --noconsole ^
  --name PDFConv ^
  --icon "%ROOT%\res\PDFConv.ico" ^
  --distpath "%ROOT%" ^
  --workpath "%ROOT%\build" ^
  --specpath "%ROOT%\build" ^
  --paths "%ROOT%\lib" ^
  --collect-all tkinterdnd2 ^
  --exclude-module pyarrow ^
  --exclude-module numpy ^
  --exclude-module pandas ^
  --exclude-module PIL ^
  --exclude-module lxml ^
  --exclude-module pdfminer ^
  --exclude-module pypdf ^
  --exclude-module cryptography ^
  --exclude-module cffi ^
  --exclude-module charset_normalizer ^
  --exclude-module pycparser ^
  "%ROOT%\src\pdfconv.pyw" || exit /b 1

echo.
echo Готово: "%ROOT%\PDFConv.exe"
echo Рядом должны лежать bin\pandoc\pandoc.exe (для .md и DOCX) и, для PDF, установленный MiKTeX.

@echo off
rem Окно-конвертер Markdown и LaTeX в PDF и DOCX (без консольного окна).
rem Через установленный Python; самостоятельный вариант - PDFConv.exe рядом.
start "" "C:\Tools\Scripts\Python314\pythonw.exe" "%~dp0src\pdfconv.pyw" %*

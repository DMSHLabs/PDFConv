@echo off
rem Окно-конвертер Markdown и LaTeX в PDF и DOCX (без консольного окна).
rem Через установленный Python; самостоятельный вариант - PDFConv.exe рядом.
start "" pyw -3.14 "%~dp0src\pdfconv.pyw" %*

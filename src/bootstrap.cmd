@echo off
rem Подготовка на новой машине: py -3.14 src\bootstrap.py [--install]
py -3.14 "%~dp0bootstrap.py" %*

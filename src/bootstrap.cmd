@echo off
rem Подготовка на новой машине: py -3 src\bootstrap.py [--install]
py -3 "%~dp0bootstrap.py" %*
